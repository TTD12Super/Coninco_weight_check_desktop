"""Calculate contractor rows without removing the original workbook layout."""

from __future__ import annotations

from copy import copy
import csv
import json
import math
import re
from functools import lru_cache
from pathlib import Path
import shutil
import subprocess
import tempfile
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, range_boundaries

from .excel_reader import ExcelReader
from .formula_catalog import build_formula_catalog, match_reference_formula
from .formula_parser import (
    FormulaEvaluationError,
    FormulaParseError,
    _normalize_source,
    canonical,
    evaluate,
    extract_assignments,
    formula_block_result,
    formula_result,
    parse_formula,
    render_excel,
)
from .input_formats import detect_source_columns
from .product_formatter import ProductFormat, calculate_section
from .quantity_calculator import calculate_K


# Original contractor sheet positions, zero-based.
_SOURCE_COLUMNS = {
    "item_no": 1,
    "description": 2,
    "unit": 3,
    "unit_duplicate": 4,
    "W": 5,
    "H": 6,
    "W1": 7,
    "H1": 8,
    "W2": 9,
    "H2": 10,
    "L_M": 11,
    "R_D": 12,
    "a": 13,
    "t": 14,
    "Q": 15,
    "K": 16,
    "plan": 17,
    "up_down": 18,
    "vertical_axis": 19,
    "formula_or_insulation_detail": 22,
    "acoustic_fireproof_detail": 23,
    "note": 24,
}

_FORMULA_MATCH_COLORS = (
    "#E53935", "#8E44AD", "#FB8C00", "#43A047", "#D81B60", "#795548",
    "#9E9D24", "#26A69A", "#F4511E", "#AB47BC", "#7CB342", "#FF7043",
    "#FDD835", "#A1887F", "#EC407A", "#FFA726", "#66BB6A", "#29B6F6",
    "#42A5F5", "#81D4FA", "#5C6BC0", "#26C6DA", "#BA68C8", "#AED581",
)

_K_COLUMN = 16
_OUTER_S_COLUMN = 20
_INNER_S_COLUMN = 21
_FIREPROOF_S_COLUMN = 21

_K_OUTPUT = "K (tính bằng máy)"
_OUTER_S_OUTPUT = "S (cách nhiệt tính bằng máy)"
_INNER_S_OUTPUT = "SCC/STA (cách nhiệt trong tính bằng máy)"

_DISPLAY_HEADERS = {
    _K_OUTPUT: "K\n(tính bằng máy)",
    _OUTER_S_OUTPUT: "S\n(cách nhiệt tính bằng máy)",
    _INNER_S_OUTPUT: "SCC/STA\n(cách nhiệt trong tính bằng máy)",
}


def read_raw_workbook(path: str | Path) -> dict[str, pd.DataFrame]:
    """Read every used cell, including the original title and header rows."""
    reader = ExcelReader(path)
    return pd.read_excel(
        reader.path,
        sheet_name=None,
        engine=reader.engine,
        header=None,
        skiprows=0,
    )


def _cell(row: pd.Series, name: str, columns: dict[str, int]) -> Any:
    index = columns.get(name)
    return row.iloc[index] if index is not None and index < len(row) else None


def _normalized_row(row: pd.Series, columns: dict[str, int]) -> dict[str, Any]:
    normalized = {name: _cell(row, name, columns) for name in columns}
    normalized["section"] = calculate_section(normalized)
    return normalized


def _section_from_heading(row: pd.Series) -> str | None:
    """Read section headings such as the COMBINE template's group rows."""
    # Section labels live in the first description columns.  Restricting the
    # scan prevents header notes such as "chống cháy" from becoming a section.
    text = " ".join(str(value).lower() for value in row.iloc[:3].tolist() if _has_value(value))
    if "chống cháy" in text or "chong chay" in text:
        return "fireproof"
    if "tiêu âm" in text or "tieu am" in text:
        return "acoustic"
    if "ống gió thường" in text or "ong gio thuong" in text:
        return "standard"
    return None


def _has_section_formula_hint(row: dict[str, Any]) -> bool:
    text = " ".join(str(row.get(name, "")).lower() for name in (
        "formula_or_insulation_detail",
        "acoustic_fireproof_detail",
    ))
    return any(marker in text for marker in ("scn", "scc", "sta"))


def _has_value(value: Any) -> bool:
    return value is not None and not (not isinstance(value, str) and pd.isna(value)) and str(value).strip() != ""


def _is_numeric(value: Any) -> bool:
    if not _has_value(value):
        return False
    try:
        float(str(value).replace(",", ".").strip())
    except (TypeError, ValueError):
        return False
    return True


def _is_calculable(row: dict[str, Any]) -> bool:
    return _has_value(row["description"]) and any(
        _is_numeric(row[name])
        for name in ("W", "H", "W1", "H1", "W2", "H2", "L_M", "R_D")
    )


def _number(value: Any) -> float:
    """Convert an Excel number to a comparable value; blank cells equal zero."""
    if not _has_value(value):
        return 0.0
    try:
        number = float(str(value).replace(",", ".").strip())
    except (TypeError, ValueError):
        return 0.0
    return 0.0 if math.isnan(number) else number


def _text_value(value: Any) -> str:
    """Return a display-safe workbook value without exposing pandas NaN."""
    return "" if not _has_value(value) else str(value).replace("\n", " ").replace("\r", " ")


def _matches_original(original: Any, calculated: float) -> bool:
    """Compare displayed three-decimal values, allowing normal rounding noise."""
    if calculated is None:
        return False
    return abs(_number(original) - calculated) <= 0.0005


def _formula_targets(label: str) -> tuple[str, ...]:
    """Map contractor labels to the two workbook quantity columns.

    ``S``/``SCN`` are the outside insulation quantity.  ``SCC``, ``STA`` and
    ``SCC/TA`` are all different labels for the one inside-insulation column.
    The shared label is therefore deliberately returned only once.
    """
    label = label.upper()
    if label in {"SCC", "STA", "SCC/TA"}:
        return ("S_inner",)
    return ("S",)


def _formula_profile(row: dict[str, Any]) -> dict[str, Any]:
    """Read contractor labels only to choose the standard formula family.

    The expressions themselves are never used as the calculation authority.
    They only tell us whether the row contains fireproof, acoustic or shared
    inside-insulation work.
    """
    assignments: list[tuple[str, str, str]] = []
    for source_name in ("formula_or_insulation_detail", "acoustic_fireproof_detail"):
        for label, expression in extract_assignments(row.get(source_name)):
            if label.upper() in {"S", "SCN", "SCC", "STA", "SCC/TA"}:
                assignments.append((label.upper(), expression, source_name))

    labels = {label for label, _expression, _source in assignments}
    inner_labels = labels & {"SCC", "STA", "SCC/TA"}
    if "SCC/TA" in inner_labels:
        mode = "shared"
    elif "SCC" in inner_labels:
        mode = "fireproof"
    elif "STA" in inner_labels:
        mode = "acoustic"
    else:
        mode = "standard"

    # MFD/FD/MSD rows are the one deliberate exception.  When the ordinary,
    # fireproof and acoustic variants all say the inside quantity is included
    # in the valve, there is no SCC/STA/SCC-TA label to identify the family.
    # Only then may the SCN expression distinguish the reference family.
    description = str(row.get("description", "")).casefold()
    detail_text = " ".join(
        str(row.get(name, ""))
        for name in ("formula_or_insulation_detail", "acoustic_fireproof_detail")
    ).casefold()
    scn_expression = " ".join(
        expression
        for label, expression, _source in assignments
        if label == "SCN"
    )
    compact_scn = re.sub(r"\s+", "", scn_expression).casefold()
    is_mfd = any(marker in description for marker in ("fire damper", "van ngăn cháy", "mfd", "msd"))
    mfd_included_variant = (
        is_mfd
        and "bao gồm trong van" in detail_text
        and "SCN" in labels
        and not inner_labels
    )
    if mfd_included_variant:
        mode = "fireproof" if "a/25" in compact_scn or "4a" in compact_scn else "acoustic"
    return {
        "assignments": assignments,
        "labels": labels,
        "inner_labels": inner_labels,
        "mode": mode,
        "has_outer": bool(labels & {"S", "SCN"}),
        "mapping_exception": "mfd_included_variant" if mfd_included_variant else "",
    }


def _formula_category(mode: str, target: str) -> str:
    """Return the category used by the selected standard catalog row."""
    if target.upper() in {"SCC", "STA", "SCC/TA"}:
        return "Chống cháy" if target.upper() == "SCC" else "Tiêu âm"
    return "Cách nhiệt"


def _strip_parentheses(value: str) -> str:
    """Normalize only cosmetic parentheses/spacing for the explicit exception."""
    return re.sub(r"[\s()]", "", _normalize_source(value)).casefold()


def _same_parentheses_only(left: str, right: str) -> bool:
    return bool(left and right) and _strip_parentheses(left) == _strip_parentheses(right)


def _formula_variables(row: dict[str, Any]) -> dict[str, float]:
    """Build the variable context used by the row formula parser."""
    return {
        "W": _number(row.get("W")),
        "H": _number(row.get("H")),
        "W1": _number(row.get("W1")),
        "H1": _number(row.get("H1")),
        "W2": _number(row.get("W2")),
        "H2": _number(row.get("H2")),
        "L_M": _number(row.get("L_M")),
        "R_D": _number(row.get("R_D")),
        "a": _number(row.get("a")),
        "t": _number(row.get("t")),
        "Q": _number(row.get("Q")),
        "K": _number(row.get("K_for_formula", row.get("K"))),
    }


def _formula_references(columns: dict[str, int], source_row: int) -> dict[str, str]:
    """Map formula variables to their actual Excel cells for one row."""
    references = {}
    for name in ("W", "H", "W1", "H1", "W2", "H2", "L_M", "R_D", "a", "t", "Q", "K"):
        column = columns.get(name)
        if column is not None:
            references[name] = f"{get_column_letter(column + 1)}{source_row}"
    return references


def _repair_known_detail_expression(
    label: str,
    expression: str,
    product_type: str,
    section: str = "",
) -> tuple[str, str]:
    """Repair confirmed template typos while retaining the source text.

    These repairs are deliberately narrow: they only match the known
    reference-template patterns.  The contractor text remains unchanged in
    the audit table; only the calculation AST uses the repaired expression.
    """
    label_key = label.upper()
    section_key = section.lower()
    compact = re.sub(r"\s+", "", expression).upper()
    if product_type == "square_elbow" and label_key == "SCC/TA":
        broken = "(K*2*(((W+R)^2-R^2)+(2*(W+R)*H)+2*R*H))/1000000"
        if compact == broken:
            corrected = "(K*(2*((W+R)^2-R^2)+2*(W+R)*H+2*R*H))/1000000"
            return corrected, "Moved the leading 2 into the top/bottom term for square-elbow SCC/TA."

    # Do not rewrite a valid detail expression to imitate the contractor's
    # cached Excel value.  In particular, ``(W''/2)^2`` and ``W''/2^2`` are
    # different expressions in Excel.  The detail text is the calculation
    # source; a disagreement with the cached value must remain visible.
    #
    repairs: dict[tuple[str, str, str], tuple[str, str]] = {
        ("standard", "reducer", "S"): (
            "(K*((W+W')*SQRT(L^2+((H-H')/2)^2)+(H+H')*SQRT(L^2)+(W-W'/2)^2)+(W+H+25*4)*2*K*Q*100)/1000000",
            "Restored the standard reducer grouping from the source Excel formula.",
        ),
        ("standard", "square_round_reducer", "S"): (
            "(K*((W+(PI()*D)/4)*SQRT(L^2+(H-(PI()*D)/4)^2)+(H+(PI()*D)/4)*SQRT(L^2+(W-(PI()*D)/4)^2))+(W+H+25*4)*2*K*Q*100+PI()*(D+25*2)*K*Q*100)/1000000",
            "Restored the missing grouping in the standard square-round reducer formula.",
        ),
        ("standard", "takeoff", "S"): (
            "(K*((W+W')*L+H'*L+H*SQRT((W-W')^2+L^2)-W*H)+(W'+H'+25*4)*2*100*K*Q)/1000000",
            "Restored the standard takeoff closing parentheses.",
        ),
        ("fireproof", "rectangular_duct", "SCC"): (
            "((W+H+4a)*2*K*1000)/1000000",
            "Restored the fireproof rectangular-duct closing parenthesis.",
        ),
        ("fireproof", "reducer", "SCN"): (
            "((K*((W+W')*SQRT(L^2+((H-H')/2)^2)+(H+H')*SQRT(L^2+((W-W')/2)^2))+(W+H+4a+4t)*2*K*Q*100*K+(W+H+8a+4t)*2*K*Q*100*K)/1000000)*a/25",
            "Restored the fireproof reducer seam grouping and missing operator.",
        ),
        ("fireproof", "reducer", "SCC"): (
            "(K*((W+W'+4a)*SQRT(L^2+((H-H')/2)^2)+(H+H'+4a)*SQRT(L^2+(W-W'/2)^2)))/1000000",
            "Restored the source Excel grouping for the fireproof reducer formula.",
        ),
        ("fireproof", "square_round_reducer", "SCN"): (
            "((K*((W+(PI()*D)/4)*SQRT(L^2+(H-(PI()*D)/4)^2)+(H+(PI()*D)/4)*SQRT(L^2+(W-(PI()*D)/4)^2))+(W+H+4a+4t)*2*K*Q*100+(W+H+8a+4t)*2*K*Q*100+PI()*(D+4a+4t)*2*K*Q*100)/1000000)*a/25",
            "Restored the fireproof square-round reducer grouping.",
        ),
        ("fireproof", "square_round_reducer", "SCC"): (
            "(K*((W+2a+(PI()*(D+2a))/4)*SQRT(L^2+(H+2a-(PI()*(D+2a))/4)^2)+(H+2a+(PI()*(D+2a))/4)*SQRT(L^2+(W+2a-(PI()*(D+2a))/4)^2)))/1000000",
            "Restored the fireproof square-round reducer PI() operand and grouping.",
        ),
        ("fireproof", "square_elbow", "SCN"): (
            "((K*(2*((W+R)^2-R^2)+2*(W+R)*H+2*R*H)+(W+H+2a+4t)*2*K*Q*100+(W+H+8a+4t)*2*K*Q*100)/1000000)*a/25",
            "Restored the fireproof square-elbow seam grouping from the source Excel formula.",
        ),
        ("fireproof", "tee", "SCC"): (
            "(K*(S_V1+S_V2+S_V3-S_V4))/1000000",
            "Restored the fireproof tee closing parenthesis.",
        ),
        ("fireproof", "takeoff", "SCN"): (
            "((K*((W+W')*L+H'*L+H*SQRT((W-W')^2+L^2)-W*H)+(W'+H'+4a+4t)*2*100*K*Q+(W'+H'+8a+4t)*2*100*K*Q)/1000000)*a/25",
            "Restored the fireproof takeoff operator and parentheses.",
        ),
        ("fireproof", "takeoff", "SCC"): (
            "(K*((W+W'+4a)*L+(H'+2a)*L+(H+2a)*SQRT((W-W')^2+L^2)-W*H))/1000000",
            "Restored the fireproof takeoff closing parenthesis.",
        ),
        ("acoustic", "rectangular_duct", "SCC/TA"): (
            "((W+H)*2*K*1000)/1000000",
            "Restored the acoustic rectangular-duct closing parenthesis.",
        ),
        ("acoustic", "reducer", "SCC/TA"): (
            "(K*((W+W')*SQRT(L^2+((H-H')/2)^2)+(H+H')*SQRT(L^2)+(W-W'/2)^2))/1000000",
            "Restored the source Excel grouping for the acoustic reducer formula.",
        ),
        ("acoustic", "square_round_reducer", "SCC/TA"): (
            "(K*((W+(PI()*D)/4)*SQRT(L^2+(H-(PI()*D)/4)^2)+(H+(PI()*D)/4)*SQRT(L^2+(W-(PI()*D)/4)^2)))/1000000",
            "Restored the acoustic square-round reducer grouping.",
        ),
        ("acoustic", "tee", "SCN"): (
            "((W'+H'+4t)*2*100*K+(W'+H'+4*25+4t)*2*100*K*Q+(W''+H''+4t)*2*100*K+(W''+H''+4t+4a)*100*K*Q+(W+H+4t)*2*100*K+(W+H+4a+4t)*2*100*K*Q)/1000000",
            "Restored the acoustic tee missing operators between area blocks.",
        ),
        ("acoustic", "tee", "SCC/TA"): (
            "(K*(S_V1+S_V2+S_V3-S_V4))/1000000",
            "Restored the acoustic tee closing parenthesis.",
        ),
        ("acoustic", "elbow", "SCC/TA"): (
            "(K*(2*PI()*((W+R)^2-R^2)+2*PI()*H*(W+R+R))*L/360)/1000000",
            "Restored the acoustic elbow closing parenthesis and angle term.",
        ),
        ("acoustic", "fcu_reducer_box", "STA"): (
            "(K*((W+W'+4a)*SQRT((L+2a)^2+((H-H')/2)^2)+(H+H'+4a)*SQRT((L+2a)^2)+(W-W'/2)^2+(W+2a)*(H+2a)+(W'+2a)*(H'+2a)-W''*H''-t*D+(W''+H'')*2*50+(t+D)*2*50))/1000000",
            "Restored the acoustic FCU reducer-box quote and closing parentheses.",
        ),
    }
    repair = repairs.get((section_key, product_type, label_key))
    if repair:
        # Only apply the table to the known malformed patterns.  Valid future
        # formulas remain fully detail-driven instead of being overwritten.
        malformed = (
            compact.count("(") != compact.count(")")
            or 'H"' in compact
            or "Q(" in compact
        )
        if malformed:
            return repair
    # A missing closing parenthesis at the end of one assignment is safe to
    # repair for calculation: it cannot change the order of existing terms.
    # Keep the original expression untouched in ``detail_block`` and in the
    # formula comparison, so the contractor's source remains auditable.
    missing_closings = expression.count("(") - expression.count(")")
    if missing_closings > 0:
        return (
            f"{expression}{')' * missing_closings}",
            f"Auto-added {missing_closings} missing closing parenthesis(es) for calculation only.",
        )
    return expression, ""


def _formula_cells(raw_row: pd.Series, source_row: int) -> dict[str, float]:
    """Expose current-row Excel cells for formulas that explicitly use E6-like refs."""
    cells = {}
    for column, value in enumerate(raw_row.tolist(), start=1):
        cells[f"{get_column_letter(column)}{source_row}"] = _number(value)
    return cells


def _row_formula_audits(
    row: dict[str, Any],
    raw_row: pd.Series,
    source_columns: dict[str, int],
    source_row: int,
    product_type: str = "",
    reference_catalog: dict[str, Any] | None = None,
) -> dict[str, dict[str, Any]]:
    """Audit contractor text against the selected standard formula.

    This function intentionally evaluates the contractor block only for the
    comparison metadata.  The returned ``value`` and ``expanded_node`` always
    come from the reference catalog, so the contractor formula can never
    become the result written to the workbook.
    """
    audits: dict[str, dict[str, Any]] = {}
    profile = _formula_profile(row)
    variables = _formula_variables(row)
    cells = _formula_cells(raw_row, source_row)
    references = _formula_references(source_columns, source_row)
    description = str(row.get("description", ""))
    mode = profile["mode"]

    def selected(label: str) -> tuple[str, str, str]:
        label_key = label.upper()
        if label_key in {"S", "SCN"}:
            standard_target = "S" if label_key == "S" and mode == "standard" else "SCN"
        elif label_key == "SCC":
            standard_target = "SCC"
        elif label_key == "STA":
            standard_target = "STA"
        else:
            standard_target = "SCC/TA"
        return standard_target, _formula_category(mode, standard_target), label_key

    def reference_result(standard_target: str, category: str):
        # A few reference rows use STA while rectangular/round products use
        # the shared SCC/TA label for the same inside quantity. Treat them as
        # equivalent targets during lookup, but keep the contractor label in
        # the audit display.
        target_candidates = [standard_target]
        if standard_target == "STA":
            target_candidates.append("SCC/TA")
        elif standard_target == "SCC/TA":
            target_candidates.append("STA")
        if standard_target == "S":
            family = "standard"
        elif standard_target == "SCN":
            family = "fireproof" if mode == "standard" else "acoustic" if mode == "shared" else mode
        else:
            family = "fireproof" if standard_target == "SCC" else "acoustic"
        for candidate_target in target_candidates:
            reference = match_reference_formula(
                reference_catalog,
                product_type,
                candidate_target,
                description,
                category=category,
                family=family,
            )
            if not reference:
                continue
            try:
                assignments = extract_assignments(reference.get("formula_block", ""))
                repair_section = (
                    "fireproof" if candidate_target == "SCC"
                    else "acoustic" if candidate_target in {"STA", "SCC/TA"}
                    else "standard" if mode == "standard" else mode
                )
                repaired_assignments = []
                correction_notes = []
                for assignment_label, assignment_expression in assignments:
                    corrected, note = _repair_known_detail_expression(
                        assignment_label,
                        assignment_expression,
                        product_type,
                        repair_section,
                    )
                    repaired_assignments.append((assignment_label, corrected))
                    if note:
                        correction_notes.append(note)
                result = formula_block_result(
                    repaired_assignments,
                    variables,
                    cells=cells,
                    references=references,
                    target=candidate_target,
                )
                return reference, result, "", candidate_target, "; ".join(dict.fromkeys(correction_notes))
            except (FormulaParseError, FormulaEvaluationError, TypeError, ValueError) as error:
                # Keep trying the equivalent target if this row is malformed;
                # if no candidate works, return the first useful diagnostic.
                last_error = str(error)
                continue
        return None, None, "No matching usable standard formula in the reference catalog.", standard_target, ""

    def contractor_result(source_name: str, target: str):
        detail = row.get(source_name)
        assignments = extract_assignments(detail)
        try:
            result = formula_block_result(
                assignments,
                variables,
                cells=cells,
                references=references,
                target=target,
            )
            return result, ""
        except (FormulaParseError, FormulaEvaluationError, TypeError, ValueError) as error:
            return None, str(error)

    # Prefer the most specific label in each output column.  SCC/TA is one
    # shared inside quantity and is therefore emitted once.
    labels = profile["labels"]
    selected_labels: list[str] = []
    if "SCN" in labels:
        selected_labels.append("SCN")
    elif "S" in labels:
        selected_labels.append("S")
    if "SCC/TA" in labels:
        selected_labels.append("SCC/TA")
    elif "SCC" in labels:
        selected_labels.append("SCC")
    elif "STA" in labels:
        selected_labels.append("STA")

    for label in selected_labels:
        source_name = ""
        expression = ""
        for candidate_source in ("formula_or_insulation_detail", "acoustic_fireproof_detail"):
            assignments = extract_assignments(row.get(candidate_source))
            match = next(
                ((candidate_label, candidate_expression) for candidate_label, candidate_expression in assignments
                 if candidate_label.upper() == label),
                None,
            )
            if match:
                source_name = candidate_source
                expression = match[1]
                break
        if not source_name:
            continue

        standard_target, category, label_key = selected(label)
        target = _formula_targets(label)[0]
        detail_assignments = extract_assignments(row.get(source_name))
        detail_block = "\n".join(f"{item} = {value}" for item, value in detail_assignments)
        reference, standard, standard_error, resolved_target, standard_correction = reference_result(standard_target, category)
        standard_value = standard.get("value") if standard else None
        source_value_column = {
            "S": "insulation_quantity",
            "S_inner": "acoustic_quantity",
        }[target]
        source_value_index = source_columns.get(source_value_column)
        numeric_match = bool(
            standard
            and standard_value is not None
            and source_value_index is not None
            and _matches_original(raw_row.iloc[source_value_index], standard_value)
        )

        # The saved quantity is the verdict. Once it matches the standard
        # result, do not parse or compare the contractor formula at all.
        contractor = None
        contractor_error = ""
        contractor_value = None
        reference_status = ""
        formula_comparison = None
        if not numeric_match:
            contractor, contractor_error = contractor_result(source_name, label)
            contractor_value = contractor.get("value") if contractor else None
            reference_status = "unverifiable"
            if standard and contractor:
                contractor_node = contractor.get("expanded_nodes", {}).get(label_key)
                standard_node = standard.get("expanded_nodes", {}).get(resolved_target)
                same_structure = (
                    contractor_node is not None
                    and standard_node is not None
                    and canonical(contractor_node) == canonical(standard_node)
                )
                parentheses_only = _same_parentheses_only(
                    expression,
                    reference.get("formula", "") if reference else "",
                )
                if same_structure or (
                    parentheses_only
                    and standard_value is not None
                    and contractor_value is not None
                    and abs(float(standard_value) - float(contractor_value)) <= 0.0005
                ):
                    reference_status = "equivalent"
                else:
                    reference_status = "different"

            formula_comparison = {
                "status": (
                    "equivalent" if reference_status == "equivalent"
                    else "different" if reference_status == "different"
                    else "unverifiable"
                ),
                "original_formula_value": standard_value,
                "detail_formula_value": contractor_value,
                "difference": (
                    float(standard_value) - float(contractor_value)
                    if standard_value is not None and contractor_value is not None
                    else None
                ),
                "blocks": [],
            }
        audit = {
            "label": label,
            "expression": expression,
            "source": source_name,
            "excel_formula": standard.get("excel_formula", "") if standard else "",
            "value": standard_value,
            "contractor_value": contractor_value,
            "error": standard_error or contractor_error,
            "matches": numeric_match,
            "detail_block": detail_block,
            "reference_formula": reference.get("formula", "") if reference else "",
            "reference_formula_block": reference.get("formula_block", "") if reference else "",
            "reference_status": reference_status,
            "reference_target": resolved_target,
            "formula_comparison": formula_comparison,
            "correction": standard_correction,
        }
        if standard:
            audit["canonical"] = standard.get("canonical")
            audit["expanded_node"] = standard.get("expanded_nodes", {}).get(resolved_target)
            audit["standard_expanded_node"] = audit["expanded_node"]
        audits[target] = audit
    return audits


def _final_formula_references(
    source_columns: dict[str, int],
    source_row: int,
    original_insertions: tuple[tuple[int, str, str, int], ...],
) -> dict[str, str]:
    """Map source fields to their columns after result columns are inserted."""
    references = {}
    for name in ("W", "H", "W1", "H1", "W2", "H2", "L_M", "R_D", "a", "t", "Q", "K"):
        source_column = source_columns.get(name)
        if source_column is None:
            continue
        final_column = source_column + 1 + sum(
            1 for original_column, _label, _value_key, _width in original_insertions
            if original_column < source_column
        )
        references[name] = f"{get_column_letter(final_column)}{source_row}"
    return references


def _render_output_formula(
    values: dict[str, Any],
    value_key: str,
    references: dict[str, str],
) -> str:
    """Render a row's parsed detail formula for an inserted result cell."""
    audit = values.get("formula_audits", {}).get(value_key)
    if not audit:
        return ""
    try:
        # The inserted Excel formula must be the standard formula.  A bad
        # contractor expression may still be shown in the audit table, but it
        # must never replace the authoritative output.
        node = audit.get("standard_expanded_node") or audit.get("expanded_node")
        if node is None:
            return ""
        return f"={render_excel(node, references)}"
    except (FormulaParseError, FormulaEvaluationError, TypeError, ValueError):
        return ""


def calculate_raw_sheet(
    data: pd.DataFrame,
    section: str | None = None,
    product_type: str | None = None,
    reference_catalog: dict[str, Any] | None = None,
) -> tuple[list[dict[str, Any]], dict[int, dict[str, Any]], list[dict[str, Any]]]:
    """Calculate rows from the standard formula catalog.

    Contractor formulas are parsed only for classification/audit.  Numeric
    output is taken from the catalog (or the built-in standard calculator when
    no catalog was supplied), never from the contractor detail text.
    """
    calculated: list[dict[str, Any]] = []
    output_values: dict[int, dict[str, Any]] = {}
    errors: list[dict[str, Any]] = []
    source_columns = detect_source_columns(data)
    active_section: str | None = None

    for index, raw_row in data.iterrows():
        heading_section = _section_from_heading(raw_row)
        if heading_section:
            active_section = heading_section
        row = _normalized_row(raw_row, source_columns)
        if active_section and _is_calculable(row) and not _has_section_formula_hint(row):
            row["section"] = active_section
        if not _is_calculable(row):
            continue

        source_row = int(index) + 1
        try:
            has_k_components = any(_has_value(row.get(column)) for column in ("plan", "up_down", "vertical_axis"))
            calculated_k = calculate_K(row) if has_k_components else _number(row.get("K"))
            # K is independent only when at least one of the three component
            # columns is populated.  Otherwise the contractor K is retained
            # as the input required by the standard formula.
            row["K_By_function"] = calculated_k if has_k_components else None
            row["K_for_formula"] = calculated_k
            profile = _formula_profile(row)
            calculation_section = {
                "standard": "standard",
                "fireproof": "fireproof",
                "acoustic": "acoustic",
                "shared": "acoustic",
            }[profile["mode"]]
            product = ProductFormat.from_row(
                row,
                section=section or calculation_section,
                product_type=product_type,
            )
            # Built-in calculators are the safe fallback when a reference
            # catalog is not available. They are never fed contractor formula
            # text and are replaced by catalog values below when possible.
            result = product.calculate(row)
            formula_audits = _row_formula_audits(
                row,
                raw_row,
                source_columns,
                source_row,
                product.product_type,
                reference_catalog,
            )
            quantities = {
                "S": result.R,
                "S_inner": result.T,
            }
            for target, audit in formula_audits.items():
                if audit["value"] is not None:
                    quantities[target] = audit["value"]
            result = result.__class__(
                section=result.section,
                product_type=result.product_type,
                R=quantities["S"],
                S=quantities["S_inner"],
                T=quantities["S_inner"],
                unit=result.unit,
            )
            inputs = product.inputs_from(row)
            calculated_row = {
                "source_row": source_row,
                "description": product.description,
                "section": product.section,
                "product_type": product.product_type,
                "formula_mode": profile["mode"],
                "inputs": inputs.__dict__,
                "result": result.to_dict(),
            }
            calculated.append(calculated_row)
            output_values[index] = {
                "K": calculated_k,
                "S": result.R,
                "S_inner": result.T,
                "S_fireproof": result.T,
                "formula_audits": formula_audits,
                "matches": {
                    "K": _matches_original(row["K"], calculated_k) if has_k_components else True,
                    "S": _matches_original(raw_row.iloc[source_columns.get("insulation_quantity", _OUTER_S_COLUMN)], result.R),
                    "S_inner": _matches_original(raw_row.iloc[source_columns.get("acoustic_quantity", _INNER_S_COLUMN)], result.T),
                },
            }
            for target, audit in formula_audits.items():
                original_column = {
                    "S": "insulation_quantity",
                    "S_inner": "acoustic_quantity",
                }[target]
                if original_column in source_columns and audit["value"] is not None:
                    numeric_match = _matches_original(
                        raw_row.iloc[source_columns[original_column]],
                        audit["value"],
                    )
                else:
                    numeric_match = False
                # The quantity is the primary verdict. Formula parsing is
                # explanatory evidence only after the saved result is wrong;
                # equivalent results must not be marked red because the
                # contractor's description uses different grouping or has a
                # harmless formatting typo.
                audit["matches"] = numeric_match
                output_values[index]["matches"][target] = audit["matches"]
        except (TypeError, ValueError) as error:
            errors.append(
                {
                    "source_row": source_row,
                    "description": row.get("description", ""),
                    "message": str(error),
                }
            )

    return calculated, output_values, errors


def _header_row(data: pd.DataFrame) -> int:
    for index, row in data.iterrows():
        text = " ".join(str(value).lower() for value in row if _has_value(value))
        if "total of quantities" in text and ("khối lượng cách nhiệt" in text or "detailed quantities" in text):
            return int(index)
    return 0


def _insert_result_columns(
    data: pd.DataFrame,
    output_values: dict[int, dict[str, Any]],
) -> pd.DataFrame:
    """Insert calculated columns beside original K, S, and inner-S columns."""
    result = data.copy()
    header = _header_row(data)

    def values(key: str, label: str) -> list[Any]:
        output = [None] * len(data)
        output[header] = label
        for index, row_values in output_values.items():
            output[index] = row_values[key]
        return output

    # Insert from left to right using positions after previous insertions.
    result.insert(_K_COLUMN + 1, _K_OUTPUT, values("K", _K_OUTPUT))
    result.insert(_OUTER_S_COLUMN + 2, _OUTER_S_OUTPUT, values("S", _OUTER_S_OUTPUT))
    result.insert(_INNER_S_COLUMN + 3, _INNER_S_OUTPUT, values("S_inner", _INNER_S_OUTPUT))
    return result


def _soffice_path() -> str:
    """Find LibreOffice, used to convert legacy .xls files without losing layout."""
    executable = shutil.which("soffice")
    if executable:
        return executable

    bundled = (
        Path.home()
        / ".cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice"
    )
    if bundled.is_file():
        return str(bundled)
    raise FileNotFoundError(
        "Không tìm thấy LibreOffice (soffice) để chuyển file .xls sang .xlsx."
    )


def _xlsx_source(source: Path, temporary_directory: Path) -> Path:
    """Return an xlsx source while keeping the original .xls file untouched."""
    if source.suffix.lower() == ".xlsx":
        return source

    subprocess.run(
        [
            _soffice_path(),
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(temporary_directory),
            str(source),
        ],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    converted = temporary_directory / f"{source.stem}.xlsx"
    if not converted.is_file():
        raise FileNotFoundError(f"Không tạo được file xlsx tạm từ {source}.")
    return converted


def _shift_merged_ranges(worksheet, insert_at: int) -> None:
    """Move merge ranges around an inserted column.

    openpyxl moves cell values for ``insert_cols`` but does not update merged
    ranges.  Rebuilding the ranges here keeps the original header layout.
    """
    ranges = [str(cell_range) for cell_range in worksheet.merged_cells.ranges]
    for cell_range in ranges:
        worksheet.unmerge_cells(cell_range)

    for cell_range in ranges:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        if min_col >= insert_at:
            min_col += 1
            max_col += 1
        elif max_col >= insert_at:
            max_col += 1
        worksheet.merge_cells(
            start_row=min_row,
            start_column=min_col,
            end_row=max_row,
            end_column=max_col,
        )


def _copy_column_dimension(worksheet, source_column: int, target_column: int) -> None:
    source = worksheet.column_dimensions[get_column_letter(source_column)]
    target = worksheet.column_dimensions[get_column_letter(target_column)]
    target.width = source.width
    target.hidden = source.hidden
    target.bestFit = source.bestFit
    target.outlineLevel = source.outlineLevel
    target.collapsed = source.collapsed


def _copy_column_style(
    worksheet,
    source_column: int,
    target_column: int,
    red_text: bool = True,
    width: float | None = None,
) -> None:
    """Copy the original column's cell styles to a newly inserted column."""
    _copy_column_dimension(worksheet, source_column, target_column)
    if width is not None:
        worksheet.column_dimensions[get_column_letter(target_column)].width = width
    for row in range(1, worksheet.max_row + 1):
        source = worksheet.cell(row=row, column=source_column)
        target = worksheet.cell(row=row, column=target_column)
        if source.has_style:
            target._style = copy(source._style)
        target.number_format = source.number_format
        target.protection = copy(source.protection)
        if red_text:
            font = copy(target.font)
            font.color = "FFFF0000"
            target.font = font


def _set_new_header(worksheet, column: int, row: int, label: str) -> None:
    """Write a new header without breaking an existing merged cell."""
    cell = worksheet.cell(row=row, column=column)
    inside_merged_header = any(
        merged.min_row <= row <= merged.max_row
        and merged.min_col <= column <= merged.max_col
        and (merged.min_row, merged.min_col) != (row, column)
        for merged in worksheet.merged_cells.ranges
    )
    if isinstance(cell, MergedCell) or inside_merged_header:
        # In the reference template the top title band is merged across
        # columns S:AA.  Put a new result label in the second header row
        # directly; scanning farther down can land on a data cell and the
        # label may disappear when openpyxl saves the merged layout.
        candidate = worksheet.cell(row=row + 1, column=column)
        if not isinstance(candidate, MergedCell):
            cell = candidate
        else:
            for candidate_row in range(row + 2, worksheet.max_row + 1):
                candidate = worksheet.cell(row=candidate_row, column=column)
                if not isinstance(candidate, MergedCell):
                    cell = candidate
                    break
    cell.value = label
    font = copy(cell.font)
    font.bold = True
    font.color = "FFFF0000"
    cell.font = font
    alignment = copy(cell.alignment)
    alignment.horizontal = "center"
    alignment.vertical = "center"
    alignment.wrap_text = True
    cell.alignment = alignment
    current_height = worksheet.row_dimensions[row].height or 0
    worksheet.row_dimensions[row].height = max(current_height, 42)


def _formula_text(value: Any) -> str:
    """Return a compact one-line representation of an original cell formula."""
    if not isinstance(value, str) or not value.startswith("="):
        return ""
    return " ".join(value.replace("\r", " ").replace("\n", " ").split())


def _replace_formula_cells(node, cell_variables: dict[str, str]):
    """Use the contractor's dimension names when comparing source ASTs."""
    if node.kind == "cell":
        key = str(node.value).replace("$", "").upper()
        name = cell_variables.get(key)
        return type(node)("name", name) if name else node
    if node.kind in ("unary", "binary", "call"):
        return type(node)(
            node.kind,
            node.value,
            tuple(_replace_formula_cells(child, cell_variables) for child in node.children),
        )
    return node


def _formula_subtrees(node, display_node=None) -> dict[Any, Any]:
    """Index source formula subexpressions so detail blocks can be matched."""
    result = {}

    def visit(current, display):
        result.setdefault(canonical(current), display)
        for index, child in enumerate(current.children):
            visit(child, display.children[index])

    visit(node, display_node or node)
    return result


def _formula_match_span_index_uncached(
    expression: str,
    cell_variables: dict[str, str],
) -> dict[Any, list[dict[str, int]]]:
    """Index every token-aligned source fragment by its normalized AST.

    A whole formula block can differ while many smaller terms still match.
    Indexing only parenthesized fragments misses those terms, so include every
    parseable range between token boundaries as well.
    """
    source = expression.strip()
    leading = len(source) - len(source.lstrip("= "))
    formula = source[leading:]
    boundaries = {0, len(formula)}
    for match in re.finditer(
        r"[A-Za-z_][A-Za-z0-9_']*|\$?[A-Za-z]{1,3}\$?\d+|"
        r"\d+(?:[.,]\d*)?|[+\-*/^(),=<>×÷−]",
        formula,
    ):
        boundaries.add(match.start())
        boundaries.add(match.end())
    ordered_boundaries = sorted(boundaries)
    candidates = [
        (start, end)
        for index, start in enumerate(ordered_boundaries[:-1])
        for end in ordered_boundaries[index + 1:]
    ]
    opening = []
    for position, character in enumerate(formula):
        if character == "(":
            opening.append(position)
        elif character == ")" and opening:
            start = opening.pop()
            candidates.append((start, position + 1))
            # Parentheses are often just grouping; include their contents too.
            inner_start, inner_end = start + 1, position
            while inner_start < inner_end and formula[inner_start].isspace():
                inner_start += 1
            while inner_end > inner_start and formula[inner_end - 1].isspace():
                inner_end -= 1
            candidates.append((inner_start, inner_end))

    indexed: dict[Any, list[dict[str, int]]] = {}
    for start, end in candidates:
        if start >= end:
            continue
        fragment = formula[start:end]
        left_trim = len(fragment) - len(fragment.lstrip())
        right_trim = len(fragment.rstrip())
        parse_start = start + left_trim
        parse_end = start + right_trim
        if parse_start >= parse_end:
            continue
        try:
            candidate = _replace_formula_cells(
                parse_formula(formula[parse_start:parse_end]),
                cell_variables,
            )
            indexed.setdefault(canonical(candidate), []).append(
                {"start": leading + parse_start, "end": leading + parse_end}
            )
        except FormulaParseError:
            continue
    for key, spans in indexed.items():
        spans.sort(key=lambda span: (span["start"], -(span["end"] - span["start"])))
        selected = []
        for span in spans:
            if any(span["start"] >= item["start"] and span["end"] <= item["end"] for item in selected):
                continue
            selected.append(span)
        indexed[key] = selected
    return indexed


@lru_cache(maxsize=256)
def _formula_match_span_index_cached(
    expression: str,
    cell_variable_items: tuple[tuple[str, str], ...],
) -> dict[Any, list[dict[str, int]]]:
    return _formula_match_span_index_uncached(expression, dict(cell_variable_items))


def _formula_match_span_index(
    expression: str,
    cell_variables: dict[str, str],
) -> dict[Any, list[dict[str, int]]]:
    """Return cached token-aligned spans for one formula and variable map."""
    return _formula_match_span_index_cached(
        expression,
        tuple(sorted(cell_variables.items())),
    )


def _formula_match_spans(
    expression: str,
    target: Any,
    cell_variables: dict[str, str],
) -> list[dict[str, int]]:
    """Find source-text ranges whose parsed formula matches a canonical block."""
    return _formula_match_span_index(expression, cell_variables).get(target, [])


def _node_size(node) -> int:
    return 1 + sum(_node_size(child) for child in node.children)


def _matched_formula_components(
    source_formula: str,
    detail_expression: str,
    source_subtrees: dict[Any, Any],
    detail_node,
    cell_variables: dict[str, str],
    preferred_color: str = "",
    used_colors: set[str] | None = None,
) -> list[dict[str, Any]]:
    """Return maximal meaningful subexpressions shared by source and detail."""
    candidates = []
    source_index = _formula_match_span_index(source_formula, cell_variables)
    detail_index = _formula_match_span_index(detail_expression, {})

    def visit(node):
        key = canonical(node)
        source_spans = source_index.get(key, [])
        detail_spans = detail_index.get(key, [])
        if _node_size(node) >= 3 and key in source_subtrees and source_spans and detail_spans:
            candidates.append((min(span["start"] for span in detail_spans), key, source_spans, detail_spans))
            # Keep visiting children too.  A matching parent does not mean
            # every nested matching term should be hidden from the report.
        for child in node.children:
            visit(child)

    visit(detail_node)
    # The same normalized subtree can occur more than once. Keep one component
    # per distinct structure; its span arrays preserve every occurrence.
    unique = {}
    for start, key, source_spans, detail_spans in candidates:
        unique.setdefault(key, (start, source_spans, detail_spans))

    used_colors = used_colors if used_colors is not None else set()
    components = []
    for index, (_key, (start, source_spans, detail_spans)) in enumerate(sorted(unique.items(), key=lambda item: item[1][0])):
        preferred = preferred_color if preferred_color and len(unique) == 1 else ""
        color = preferred if preferred and preferred not in used_colors else next(
            (candidate for candidate in _FORMULA_MATCH_COLORS if candidate not in used_colors),
            _FORMULA_MATCH_COLORS[index % len(_FORMULA_MATCH_COLORS)],
        )
        used_colors.add(color)
        components.append({
            "color": color,
            "original_spans": source_spans,
            "detail_spans": detail_spans,
            "original_expression": source_formula[
                source_spans[0]["start"]:source_spans[0]["end"]
            ],
            "detail_expression": detail_expression[
                detail_spans[0]["start"]:detail_spans[0]["end"]
            ],
        })
    return components


def _formula_comparison(
    original_formula: str,
    detail_text: str,
    inputs: dict[str, Any],
    raw_row: pd.Series,
    source_columns: dict[str, int],
    source_row: int,
    value_worksheet=None,
) -> dict[str, Any]:
    """Compare formula structure and row results without guessing unmatched terms."""
    if not original_formula or not detail_text:
        return {"status": "unverifiable", "error": "Original formula or detail formula is missing.", "blocks": []}

    cells = _formula_cells(raw_row, source_row)
    if value_worksheet is not None:
        for reference in re.findall(r"\$?[A-Za-z]{1,3}\$?\d+", original_formula):
            address = reference.replace("$", "").upper()
            value = value_worksheet[address].value
            if _has_value(value):
                cells[address] = _number(value)
    references = _formula_references(source_columns, source_row)
    cell_variables = {
        address.replace("$", "").upper(): name
        for name, address in references.items()
    }
    input_fields = {
        "W": "w", "H": "h", "W1": "w1", "H1": "h1",
        "W2": "w2", "H2": "h2", "L_M": "L_M", "R_D": "R_D",
        "a": "a", "t": "t", "Q": "Q", "K": "K",
    }
    variables = {
        name: float(inputs.get(field, inputs.get(name)) or 0)
        for name, field in input_fields.items()
    }
    try:
        original_node = parse_formula(original_formula.lstrip("= ").strip())
        original_value = float(evaluate(original_node, {}, cells))
        normalized_original = _replace_formula_cells(original_node, cell_variables)
        source_subtrees = _formula_subtrees(normalized_original, original_node)
    except (FormulaParseError, FormulaEvaluationError, TypeError, ValueError) as error:
        return {
            "status": "unverifiable",
            "error": str(error),
            "original_formula_value": None,
            "blocks": [],
        }

    assignments = extract_assignments(detail_text)
    parsed_assignments = []
    parse_errors = {}
    parsed_nodes = {}
    for label, expression in assignments:
        key = label.upper()
        try:
            parsed_nodes[key] = parse_formula(expression)
            parsed_assignments.append((label, expression))
        except FormulaParseError as error:
            parse_errors[key] = str(error)

    output_names = {"S", "SCN", "SCC", "SCC/TA", "STA"}

    # Resolve each valid assignment independently; a malformed or unresolved
    # sibling must not hide formula blocks that can still be compared. Do not
    # use formula_block_result for intermediate labels: its target is reserved
    # for the final output labels above.
    resolved = []
    resolved_values: dict[str, float] = {}
    unresolved_errors = {}
    pending = list(parsed_assignments)
    while pending:
        remaining = []
        progress = False
        for assignment in pending:
            label, expression = assignment
            key = label.upper()
            try:
                single = formula_result(
                    expression,
                    {**variables, **resolved_values},
                    cells=cells,
                )
            except (FormulaParseError, FormulaEvaluationError, TypeError, ValueError) as error:
                unresolved_errors[key] = str(error)
                remaining.append(assignment)
                continue
            resolved.append(assignment)
            resolved_values[key] = single["value"]
            unresolved_errors.pop(key, None)
            progress = True
        if not progress:
            break
        pending = remaining

    label_results = {}
    if resolved:
        for label, _expression in resolved:
            if label.upper() not in output_names:
                continue
            label_results[label.upper()] = formula_block_result(
                resolved,
                variables,
                cells=cells,
                target=label,
            )

    output_label = next(
        (label.upper() for label, _expression in assignments
         if label.upper() in output_names and label.upper() in label_results),
        None,
    )
    detail_value = label_results[output_label]["value"] if output_label else None
    if detail_value is None:
        status = "unverifiable"
    elif canonical(normalized_original) == canonical(label_results[output_label]["expanded_nodes"][output_label]):
        status = "equivalent"
    elif abs(original_value - float(detail_value)) <= 0.0005:
        status = "same_value_structure_diff"
    else:
        status = "different"

    block_colors = {
        "S_V1": "#8E44AD",
        "S_V2": "#26A69A",
        "S_V3": "#F39C12",
        "S_V4": "#E74C3C",
        "S_V5": "#3498DB",
        "S_V6": "#2ECC71",
        "S_V7": "#D81B60",
        "S_V8": "#795548",
        "S_V9": "#8A9A0A",
        "S_V10": "#00ACC1",
        "S_V11": "#FDD835",
        "S_V12": "#FF7043",
    }
    used_colors: set[str] = set()
    blocks = []
    output_block_labels = {"S", "SCN", "SCC", "SCC/TA", "STA"}
    for label, expression in assignments:
        key = label.upper()
        is_output_block = key in output_block_labels
        node = parsed_nodes.get(key)
        result = label_results.get(key)
        expanded = result["expanded_nodes"].get(key, node) if result else node
        normalized = _replace_formula_cells(expanded, cell_variables) if expanded else None
        block_key = canonical(normalized) if normalized else None
        matched = source_subtrees.get(block_key) if block_key is not None else None
        color = block_colors.get(key, "") if matched is not None else ""
        if color in used_colors:
            color = next((candidate for candidate in _FORMULA_MATCH_COLORS if candidate not in used_colors), color)
        components = _matched_formula_components(
            original_formula,
            expression,
            source_subtrees,
            expanded,
            cell_variables,
            preferred_color=color,
            used_colors=used_colors,
        ) if expanded is not None and is_output_block else []
        if matched is not None and not is_output_block and not color:
            color = next(
                (candidate for candidate in _FORMULA_MATCH_COLORS if candidate not in used_colors),
                _FORMULA_MATCH_COLORS[len(used_colors) % len(_FORMULA_MATCH_COLORS)],
            )
            used_colors.add(color)
        if matched is not None and components:
            color = components[0]["color"]
        blocks.append({
            "label": label,
            "expression": expression,
            "value": result["values"].get(key) if result else resolved_values.get(key),
            "status": "matched" if matched is not None else "unmatched",
            "original_expression": render_excel(matched, {}) if matched is not None else "",
            "color": color,
            "original_spans": _formula_match_spans(original_formula, block_key, cell_variables) if color else [],
            "matched_components": components,
            "error": parse_errors.get(key) or unresolved_errors.get(key, ""),
        })

    errors = [*parse_errors.values(), *unresolved_errors.values()]
    comparison = {
        "status": status,
        "original_formula_value": original_value,
        "detail_formula_value": detail_value,
        "difference": original_value - float(detail_value) if detail_value is not None else None,
        "blocks": blocks,
    }
    if errors:
        comparison["error"] = "; ".join(dict.fromkeys(errors))
    return comparison


_CELL_REFERENCE_RE = re.compile(r"\$?[A-Za-z]{1,3}\$?\d+")
_REFERENCE_COLORS = (
    "FF1565C0",
    "FF2E7D32",
    "FFEF6C00",
    "FF6A1B9A",
    "FF00838F",
)


def _rich_formula_text(original: str, detail: str) -> CellRichText:
    """Make cell references visually distinct without changing the source text."""
    text = f"Original: {original or '(not available)'}\nDetail:   {detail or '(not available)'}"
    rich = CellRichText()
    cursor = 0
    colors: dict[str, str] = {}
    for match in _CELL_REFERENCE_RE.finditer(text):
        if match.start() > cursor:
            rich.append(text[cursor:match.start()])
        reference = match.group(0)
        key = reference.replace("$", "").upper()
        color = colors.setdefault(key, _REFERENCE_COLORS[len(colors) % len(_REFERENCE_COLORS)])
        rich.append(TextBlock(InlineFont(color=color, b=True), reference))
        cursor = match.end()
    if cursor < len(text):
        rich.append(text[cursor:])
    return rich


def _write_original_formula_table(
    worksheet,
    formula_worksheet,
    value_worksheet,
    data: pd.DataFrame,
    source_columns: dict[str, int],
    output_values: dict[int, dict[str, Any]],
    descriptions: dict[int, str],
) -> None:
    """Append only mismatched formula rows with original/detail text."""
    formula_rows = []
    source_value_columns = {
        "S": "insulation_quantity",
        "S_inner": "acoustic_quantity",
    }
    for raw_index, values in output_values.items():
        audits = values.get("formula_audits", {})
        mismatched = [audit for audit in audits.values() if not audit.get("matches", False)]
        if not mismatched:
            continue
        source_row = int(raw_index) + 1
        original_parts = []
        detail_parts = []
        original_values = []
        calculated_values = []
        differences = []
        for target, audit in audits.items():
            if audit.get("matches", False) and audit.get("reference_status") != "different":
                continue
            source_name = source_value_columns.get(target)
            original_formula = ""
            original_value: Any = None
            if source_name in source_columns:
                source_column = source_columns[source_name]
                original_formula = _formula_text(
                    formula_worksheet.cell(source_row, source_column + 1).value
                )
                original_value = value_worksheet.cell(source_row, source_column + 1).value
            # The comparison is standard-vs-contractor.  Keep the source
            # Excel formula/value below as an additional trace when present.
            original_parts.append(f"{target}: {audit.get('reference_formula') or original_formula}")
            detail_parts.append(f"{target}: {audit.get('expression', '')}")
            original_values.append(f"{target}={_text_value(original_value)}")
            calculated = audit.get("value")
            calculated_values.append(
                f"{target}={audit.get('error') or _text_value(calculated)}"
            )
            differences.append(
                f"{target}={_number(original_value) - _number(calculated)}"
                if calculated is not None and not audit.get("error")
                else f"{target}=error"
            )
        item_no = _text_value(data.iloc[raw_index].iloc[source_columns["item_no"]])
        formula_rows.append(
            (
                item_no or str(source_row),
                descriptions.get(raw_index, ""),
                _rich_formula_text(" | ".join(original_parts), " | ".join(detail_parts)),
                " | ".join(original_values),
                " | ".join(calculated_values),
                " | ".join(differences),
            )
        )

    if not formula_rows:
        return

    start_row = worksheet.max_row + 3
    headers = ["STT", "Name", "Formula audit", "Original value", "Calculated value", "Difference"]
    thin_gray = Side(style="thin", color="FFD9D9D9")
    header_fill = PatternFill("solid", fgColor="FFFFE5E5")
    for column, header in enumerate(headers, start=1):
        cell = worksheet.cell(start_row, column, header)
        cell.font = Font(bold=True, color="FF000000")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(top=thin_gray, bottom=thin_gray, left=thin_gray, right=thin_gray)

    for row_offset, values in enumerate(formula_rows, start=1):
        row_number = start_row + row_offset
        for column, value in enumerate(values, start=1):
            cell = worksheet.cell(row_number, column, value)
            cell.border = Border(top=thin_gray, bottom=thin_gray, left=thin_gray, right=thin_gray)
            cell.alignment = Alignment(
                horizontal="left" if column in (2, 3) else "center",
                vertical="center",
                wrap_text=False,
            )
        worksheet.row_dimensions[row_number].height = 38

    widths = {1: 8, 2: 42, 3: 160, 4: 22, 5: 22, 6: 22}
    for column, width in widths.items():
        worksheet.column_dimensions[get_column_letter(column)].width = width
    worksheet.auto_filter.ref = f"A{start_row}:F{start_row + len(formula_rows)}"


def _write_preview(
    destination: Path,
    data: pd.DataFrame,
    calculated: list[dict[str, Any]],
    output_values: dict[int, dict[str, Any]],
    formula_worksheet=None,
    value_worksheet=None,
    source_columns: dict[str, int] | None = None,
) -> Path:
    """Write a compact sidecar for the desktop table preview."""
    preview = destination.with_name(f"{destination.stem}_preview.tsv")
    source_columns = source_columns or detect_source_columns(data)
    headers = [
        "source_row", "item_no", "description", "unit", "unit_duplicate",
        "section", "product_type",
        "w", "h", "w1", "h1", "w2", "h2", "l_m", "r_d", "a", "t", "q",
        "plan", "up_down", "vertical_axis", "formula_detail", "acoustic_fireproof_detail", "note",
        "k_by_function", "contractor_k", "calculated_k", "k_match",
        "contractor_s", "calculated_s", "s_match", "contractor_fireproof_s", "calculated_fireproof_s", "fireproof_match", "contractor_inner_s",
        "calculated_inner_s", "inner_match",
        "formula_audits", "input_references",
    ]
    with preview.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.writer(stream, delimiter="\t", lineterminator="\n")
        writer.writerow(headers)
        for item in calculated:
            source_row = int(item["source_row"])
            raw_row = data.iloc[source_row - 1]
            values = output_values[source_row - 1]
            result = item["result"]["quantities"]
            inputs = item["inputs"]
            formula_audits = []
            source_value_columns = {
                "S": "insulation_quantity",
                "S_inner": "acoustic_quantity",
            }
            for target, audit in values.get("formula_audits", {}).items():
                source_name = source_value_columns.get(target)
                source_column = source_columns.get(source_name) if source_name else None
                original_formula = ""
                original_value = None
                if source_column is not None and formula_worksheet is not None:
                    original_formula = _formula_text(
                        formula_worksheet.cell(source_row, source_column + 1).value
                    )
                if source_column is not None and value_worksheet is not None:
                    original_value = value_worksheet.cell(source_row, source_column + 1).value
                calculated_value = audit.get("value")
                comparison = None
                if not audit.get("matches"):
                    comparison = audit.get("formula_comparison") or _formula_comparison(
                        original_formula,
                        audit.get("detail_block", ""),
                        inputs,
                        raw_row,
                        source_columns,
                        source_row,
                        value_worksheet,
                    )
                comparison_values = comparison or {}
                formula_audits.append({
                    "target": target,
                    "label": audit.get("label", target),
                    "original_formula": original_formula,
                    "detail_formula": audit.get("expression", ""),
                    "detail_block": audit.get("detail_block", ""),
                    "original_value": _number(original_value),
                    "original_formula_value": comparison_values.get("original_formula_value"),
                    "calculated_value": calculated_value,
                    "difference": (
                        _number(original_value) - _number(calculated_value)
                        if calculated_value is not None and not audit.get("error")
                        else None
                    ),
                    "matches": bool(audit.get("matches")),
                    "error": audit.get("error", ""),
                    "correction": audit.get("correction", ""),
                    "reference_formula": audit.get("reference_formula", ""),
                    "reference_formula_block": audit.get("reference_formula_block", ""),
                    "reference_status": audit.get("reference_status", ""),
                    "formula_comparison": comparison,
                })

            # ``formula_audits`` normally comes from a parsed detail formula.
            # K is different: it is calculated from Plan + Up/Down + Vertical
            # Axis, so it has no S/SCN/SCC/STA assignment to add to that list.
            # Also keep a visible fallback for an output mismatch whose detail
            # text could not be parsed.  Otherwise the comparison table can
            # show a red cell while Formula error review appears empty.
            audit_by_target = {str(audit["target"]): audit for audit in formula_audits}

            def source_formula(column_name: str) -> str:
                column = source_columns.get(column_name)
                if column is None or formula_worksheet is None:
                    return ""
                return _formula_text(formula_worksheet.cell(source_row, column + 1).value)

            def detail_text(*column_names: str) -> str:
                return "\n".join(
                    text
                    for text in (_text_value(raw_row.iloc[source_columns[name]]) for name in column_names if name in source_columns)
                    if text
                )

            mismatch_specs = (
                ("K", "K", "K", values.get("K"), source_columns.get("K"), "K = PLAN + UP-DOWN + VERTICAL AXIS"),
                ("S", "S", "insulation_quantity", result["R"], source_columns.get("insulation_quantity"), detail_text("formula_or_insulation_detail")),
                ("S_inner", "SCC / STA / SCC-TA", "acoustic_quantity", result["T"], source_columns.get("acoustic_quantity"), detail_text("acoustic_fireproof_detail", "formula_or_insulation_detail")),
            )
            for target, label, source_name, calculated_value, source_column, detail in mismatch_specs:
                if values["matches"].get(target, True) or target in audit_by_target:
                    continue
                original_value = (
                    _number(raw_row.iloc[source_column])
                    if source_column is not None
                    else 0.0
                )
                original = source_formula(source_name)
                if not original:
                    original = f"{label} source value = {_text_value(raw_row.iloc[source_column])}" if source_column is not None else "Original value not available"
                detail_block = detail or f"{label} calculated value = {_number(calculated_value)}"
                audit_by_target[target] = {
                    "target": target,
                    "label": label,
                    "original_formula": original,
                    "detail_formula": detail,
                    "detail_block": detail_block,
                    "original_value": original_value,
                    "calculated_value": _number(calculated_value),
                    "difference": original_value - _number(calculated_value),
                    "matches": False,
                    "error": "No parsed detail formula" if target != "K" and not detail else "",
                    "correction": "",
                }

            formula_audits = [audit_by_target[target] for target in ("K", "S", "S_inner") if target in audit_by_target]
            writer.writerow([
                source_row,
                _text_value(raw_row.iloc[source_columns["item_no"]]),
                _text_value(item["description"]),
                _text_value(raw_row.iloc[source_columns["unit"]]),
                _text_value(raw_row.iloc[source_columns.get("unit_duplicate", source_columns["unit"])]),
                item["section"],
                item["product_type"],
                inputs["w"],
                inputs["h"],
                inputs["w1"],
                inputs["h1"],
                inputs["w2"],
                inputs["h2"],
                inputs["L_M"],
                inputs["R_D"],
                inputs["a"],
                inputs["t"],
                inputs["Q"],
                _text_value(raw_row.iloc[source_columns["plan"]]),
                _text_value(raw_row.iloc[source_columns["up_down"]]),
                _text_value(raw_row.iloc[source_columns["vertical_axis"]]),
                _text_value(raw_row.iloc[source_columns["formula_or_insulation_detail"]]),
                _text_value(raw_row.iloc[source_columns["acoustic_fireproof_detail"]]),
                _text_value(raw_row.iloc[source_columns["note"]]),
                inputs["K"],
                _number(raw_row.iloc[source_columns["K"]]),
                inputs["K"],
                values["matches"]["K"],
                _number(raw_row.iloc[source_columns["insulation_quantity"]]),
                result["R"],
                values["matches"]["S"],
                _number(raw_row.iloc[source_columns["acoustic_quantity"]]),
                result["S"],
                values["matches"]["S_inner"],
                _number(raw_row.iloc[source_columns["acoustic_quantity"]]),
                result["T"],
                values["matches"]["S_inner"],
                json.dumps(formula_audits, ensure_ascii=False, separators=(",", ":")),
                json.dumps(_formula_references(source_columns, source_row), ensure_ascii=False, separators=(",", ":")),
            ])
    return preview


def export_calculated_workbook(
    input_path: str | Path,
    output_path: str | Path | None = None,
    sheet_name: str = "02",
    reference_path: str | Path | None = None,
    reference_sheet: str | None = None,
    reference_catalog: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Export a full workbook while preserving the source workbook layout."""
    source = Path(input_path).expanduser().resolve()
    destination = (
        Path(output_path).expanduser().resolve()
        if output_path is not None
        else source.with_name(f"{source.stem}_calculated.xlsx")
    )
    if source == destination:
        raise ValueError("The output workbook must be different from the input workbook.")

    if destination.suffix.lower() != ".xlsx":
        raise ValueError("File kết quả phải có đuôi .xlsx để giữ nguyên định dạng.")

    if reference_catalog is None:
        reference_catalog = (
            build_formula_catalog(reference_path, reference_sheet)
            if reference_path is not None
            else None
        )
    workbook = read_raw_workbook(source)
    if sheet_name not in workbook:
        raise ValueError(f"Worksheet not found: {sheet_name}")

    source_columns = detect_source_columns(workbook[sheet_name])
    calculated, output_values, errors = calculate_raw_sheet(
        workbook[sheet_name],
        reference_catalog=reference_catalog,
    )
    mismatch_count = sum(
        not match
        for values in output_values.values()
        for match in values["matches"].values()
    )

    # Read and edit the actual workbook instead of recreating it from a DataFrame.
    # This preserves sheet names, merged cells, row heights, fonts and cell styles.
    with tempfile.TemporaryDirectory(prefix="excel_auditor_") as temporary_name:
        xlsx_source = _xlsx_source(source, Path(temporary_name))
        workbook_out = load_workbook(xlsx_source)
        formula_workbook = load_workbook(xlsx_source, data_only=False)
        value_workbook = load_workbook(xlsx_source, data_only=True)
        worksheet = workbook_out[sheet_name]

        # Insert from left to right. These are Excel's 1-based column numbers.
        original_insertions = (
            (source_columns["K"], _K_OUTPUT, "K", 18),
            (source_columns["insulation_quantity"], _OUTER_S_OUTPUT, "S", 28),
            (source_columns["acoustic_quantity"], _INNER_S_OUTPUT, "S_inner", 34),
        )
        insertions = []
        inserted_before = 0
        for original_column, label, value_key, width in sorted(original_insertions):
            current_source = original_column + 1 + inserted_before
            insertions.append((current_source + 1, current_source, label, value_key, width))
            inserted_before += 1
        for insert_at, source_column, label, value_key, width in insertions:
            _shift_merged_ranges(worksheet, insert_at)
            worksheet.insert_cols(insert_at, 1)
            _copy_column_style(worksheet, source_column, insert_at, width=width)
            for raw_index, values in output_values.items():
                source_row = int(raw_index) + 1
                formula_references = _final_formula_references(
                    source_columns,
                    source_row,
                    original_insertions,
                )
                output_formula = _render_output_formula(
                    values,
                    value_key,
                    formula_references,
                )
                cell = worksheet.cell(
                    row=source_row,
                    column=insert_at,
                )
                cell.value = output_formula or values[value_key]
                font = copy(cell.font)
                is_match = values["matches"].get(value_key, True)
                font.color = "FF008000" if is_match else "FFC62828"
                font.bold = not is_match
                cell.font = font

        # Write headers after all insertions so merged two-row templates do
        # not move or hide labels written before a later insertion.
        header_row = _header_row(workbook[sheet_name]) + 1
        for insert_at, _source_column, label, _value_key, _width in insertions:
            _set_new_header(worksheet, insert_at, header_row, _DISPLAY_HEADERS[label])

        descriptions = {
            int(item["source_row"]) - 1: str(item.get("description", ""))
            for item in calculated
        }
        _write_original_formula_table(
            worksheet,
            formula_workbook[sheet_name],
            value_workbook[sheet_name],
            workbook[sheet_name],
            source_columns,
            output_values,
            descriptions,
        )

        destination.parent.mkdir(parents=True, exist_ok=True)
        workbook_out.calculation.fullCalcOnLoad = True
        workbook_out.calculation.forceFullCalc = True
        workbook_out.calculation.calcOnSave = True
        workbook_out.save(destination)

    preview_path = _write_preview(
        destination,
        workbook[sheet_name],
        calculated,
        output_values,
        formula_workbook[sheet_name],
        value_workbook[sheet_name],
        source_columns,
    )

    return {
        "input": str(source),
        "output": str(destination),
        "sheet": sheet_name,
        "row_count": len(workbook[sheet_name]),
        "calculated_count": len(calculated),
        "mismatch_count": mismatch_count,
        "reference_formula_count": reference_catalog["count"] if reference_catalog else 0,
        "reference_sheet": reference_catalog["sheet"] if reference_catalog else "",
        "preview": str(preview_path),
        "errors": errors,
    }


__all__ = ["calculate_raw_sheet", "export_calculated_workbook", "read_raw_workbook"]
