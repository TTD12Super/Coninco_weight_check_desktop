"""Calculate contractor rows without removing the original workbook layout."""

from __future__ import annotations

from copy import copy
import csv
import json
import math
import re
from collections import Counter
from functools import lru_cache
from fractions import Fraction
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
    _variable_name,
    _normalize_source,
    canonical,
    evaluate,
    extract_assignments,
    formula_block_result,
    formula_result,
    parse_formula,
    render_excel,
)
from .input_formats import detect_source_columns, quantity_column_for_target
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
_FIREPROOF_S_OUTPUT = "SCC (chống cháy tính bằng máy)"
_INNER_S_OUTPUT = "SCC/STA (cách nhiệt trong tính bằng máy)"

_DISPLAY_HEADERS = {
    _K_OUTPUT: "K\n(tính bằng máy)",
    _OUTER_S_OUTPUT: "S\n(cách nhiệt tính bằng máy)",
    _FIREPROOF_S_OUTPUT: "SCC\n(chống cháy tính bằng máy)",
    _INNER_S_OUTPUT: "SCC/STA\n(cách nhiệt trong tính bằng máy)",
}


# Purpose: Read every used cell, including the original title and header rows.
# Example output: {"02": DataFrame(rows=61, columns=25)}
def read_raw_workbook(path: str | Path) -> dict[str, pd.DataFrame]:
    """Read every used cell, including the original title and header rows."""
    # Compute and store: reader.
    reader = ExcelReader(path)
    # Return the value produced by this step.
    return pd.read_excel(
        reader.path,
        sheet_name=None,
        engine=reader.engine,
        header=None,
        skiprows=0,
    )


# Purpose: Read one named input from its detected zero-based source column.
# Example output: 800.0
def _cell(row: pd.Series, name: str, columns: dict[str, int]) -> Any:
    """Read one named input from its detected zero-based source column."""
    # Compute and store: index.
    index = columns.get(name)
    # Return the value produced by this step.
    return row.iloc[index] if index is not None and index < len(row) else None


# Purpose: Convert one raw worksheet row into the engine's named input fields.
# Example output: {"W": 800.0, "H": 400.0, "K": 5.325, "section": "standard"}
def _normalized_row(row: pd.Series, columns: dict[str, int]) -> dict[str, Any]:
    """Convert one raw worksheet row into the engine's named input fields."""
    # Compute and store: normalized.
    normalized = {name: _cell(row, name, columns) for name in columns}
    # Compute and store: normalized['section'].
    normalized["section"] = calculate_section(normalized)
    # Return the value produced by this step.
    return normalized


# Purpose: Read section headings such as the COMBINE template's group rows.
def _section_from_heading(row: pd.Series) -> str | None:
    """Read section headings such as the COMBINE template's group rows."""
    # Section labels live in the first description columns.  Restricting the
    # scan prevents header notes such as "chống cháy" from becoming a section.
    text = " ".join(str(value).lower() for value in row.iloc[:3].tolist() if _has_value(value))
    # Check this condition before choosing the next branch.
    if "chống cháy" in text or "chong chay" in text:
        # Return the value produced by this step.
        return "fireproof"
    # Check this condition before choosing the next branch.
    if "tiêu âm" in text or "tieu am" in text:
        # Return the value produced by this step.
        return "acoustic"
    # Check this condition before choosing the next branch.
    if "ống gió thường" in text or "ong gio thuong" in text:
        # Return the value produced by this step.
        return "standard"
    # Return the value produced by this step.
    return None


# Purpose: Return whether formula labels explicitly identify a non-standard section.
# Example output: True
def _has_section_formula_hint(row: dict[str, Any]) -> bool:
    """Return whether formula labels explicitly identify a non-standard section."""
    # Compute and store: text.
    text = " ".join(str(row.get(name, "")).lower() for name in (
        "formula_or_insulation_detail",
        "acoustic_fireproof_detail",
    ))
    # Return the value produced by this step.
    return any(marker in text for marker in ("scn", "scc", "sta"))


# Purpose: Treat blanks and pandas NaN as absent workbook values.
# Example output: True
def _has_value(value: Any) -> bool:
    """Treat blanks and pandas NaN as absent workbook values."""
    # Return the value produced by this step.
    return value is not None and not (not isinstance(value, str) and pd.isna(value)) and str(value).strip() != ""


# Purpose: Check numeric-looking cell text without raising during row filtering.
# Example output: True
def _is_numeric(value: Any) -> bool:
    """Check numeric-looking cell text without raising during row filtering."""
    # Check this condition before choosing the next branch.
    if not _has_value(value):
        # Return the value produced by this step.
        return False
    # Protect parsing/evaluation so bad workbook data becomes a controlled error.
    try:
        # Execute this expression for its result or side effect.
        float(str(value).replace(",", ".").strip())
    except (TypeError, ValueError):
        # Return the value produced by this step.
        return False
    # Return the value produced by this step.
    return True


# Purpose: Keep only product rows with a description and at least one dimension.
# Example output: True
def _is_calculable(row: dict[str, Any]) -> bool:
    """Keep only product rows with a description and at least one dimension."""
    # Return the value produced by this step.
    return _has_value(row["description"]) and any(
        _is_numeric(row[name])
        for name in ("W", "H", "W1", "H1", "W2", "H2", "L_M", "R_D")
    )


# Purpose: Convert an Excel number to a comparable value; blank cells equal zero.
# Example output: 1.0
def _number(value: Any) -> float:
    """Convert an Excel number to a comparable value; blank cells equal zero."""
    # Check this condition before choosing the next branch.
    if not _has_value(value):
        # Return the value produced by this step.
        return 0.0
    # Protect parsing/evaluation so bad workbook data becomes a controlled error.
    try:
        # Compute and store: number.
        number = float(str(value).replace(",", ".").strip())
    except (TypeError, ValueError):
        # Return the value produced by this step.
        return 0.0
    # Return the value produced by this step.
    return 0.0 if math.isnan(number) else number


# Purpose: Return a display-safe workbook value without exposing pandas NaN.
# Example output: "Square elbow 360x360"
def _text_value(value: Any) -> str:
    """Return a display-safe workbook value without exposing pandas NaN."""
    # Return the value produced by this step.
    return "" if not _has_value(value) else str(value).replace("\n", " ").replace("\r", " ")


# Purpose: Compare displayed three-decimal values, allowing normal rounding noise.
# Example output: True
def _matches_original(original: Any, calculated: float) -> bool:
    """Compare displayed three-decimal values, allowing normal rounding noise."""
    # Check this condition before choosing the next branch.
    if calculated is None:
        # Return the value produced by this step.
        return False
    # Return the value produced by this step.
    return abs(_number(original) - calculated) <= 0.0005


# Purpose: Map contractor labels to the logical output quantity columns.
# Example output: ("S_inner",)
def _formula_targets(
    label: str,
    source_columns: dict[str, int] | None = None,
) -> tuple[str, ...]:
    """Map contractor labels to the logical output quantity columns.

    ``S``/``SCN`` are the outside insulation quantity.  ``SCC`` is the
    fireproof quantity and ``STA`` is the inside/acoustic quantity.  A
    contractor's combined ``SCC/TA`` block is one shared result.  In a
    three-quantity workbook it follows the SCC column; in a legacy
    two-quantity workbook it follows the shared inside column.
    """
    # Compute and store: label.
    label = label.upper()
    # Check this condition before choosing the next branch.
    if label == "SCC":
        # Check this condition before choosing the next branch.
        if source_columns and (
            "scc_quantity" in source_columns or "fireproof_quantity" in source_columns
        ):
            # Return the value produced by this step.
            return ("S_fireproof",)
        # Return the value produced by this step.
        return ("S_inner",)
    # Check this condition before choosing the next branch.
    if label == "SCC/TA":
        # Check this condition before choosing the next branch.
        if source_columns and (
            "scc_quantity" in source_columns or "fireproof_quantity" in source_columns
        ):
            # Return the value produced by this step.
            return ("S_fireproof",)
        # Return the value produced by this step.
        return ("S_inner",)
    # Check this condition before choosing the next branch.
    if label == "STA":
        # Return the value produced by this step.
        return ("S_inner",)
    # Return the value produced by this step.
    return ("S",)


# Purpose: Resolve a logical audit target to the contractor's quantity column.
def _audit_quantity_column(source_columns: dict[str, int], target: str) -> str | None:
    """Resolve a logical audit target to the contractor's quantity column."""
    # Compute and store: target_map.
    target_map = {"S": "SCN", "S_fireproof": "SCC", "S_inner": "STA"}
    # Return the value produced by this step.
    return quantity_column_for_target(source_columns, target_map.get(target, target))


# Purpose: Read contractor labels only to choose the standard formula family.
# Example output: {"mode": "shared", "labels": {"S", "SCC/TA"}, ...}
def _formula_profile(row: dict[str, Any]) -> dict[str, Any]:
    """Read contractor labels only to choose the standard formula family.

    This detects standard, fireproof, acoustic, and shared modes.

    The expressions themselves are never used as the calculation authority.
    They only tell us whether the row contains fireproof, acoustic or shared
    inside-insulation work.
    """
    # Compute and store: assignments.
    assignments: list[tuple[str, str, str]] = []
    # Compute and store: source_assignments.
    source_assignments: dict[str, list[tuple[str, str]]] = {}
    # Process each item in this input sequence.
    for source_name in ("formula_or_insulation_detail", "acoustic_fireproof_detail"):
        # Compute and store: parsed_assignments.
        parsed_assignments = extract_assignments(row.get(source_name))
        # Compute and store: source_assignments[source_name].
        source_assignments[source_name] = parsed_assignments
        # Process each item in this input sequence.
        for label, expression in parsed_assignments:
            # Check this condition before choosing the next branch.
            if label.upper() in {"S", "SCN", "SCC", "STA", "SCC/TA"}:
                # Execute this expression for its result or side effect.
                assignments.append((label.upper(), expression, source_name))

    # Compute and store: labels.
    labels = {label for label, _expression, _source in assignments}
    # Compute and store: inner_labels.
    inner_labels = labels & {"SCC", "STA", "SCC/TA"}
    # Check this condition before choosing the next branch.
    if "SCC/TA" in inner_labels:
        # Compute and store: mode.
        mode = "shared"
    # Check this condition before choosing the next branch.
    elif "STA" in inner_labels:
        # Some acoustic rows contain both STA and SCC quantities. STA is the
        # section/family signal here; SCC is the separate fireproof output
        # column, not a reason to switch the whole row to fireproof formulas.
        mode = "acoustic"
    # Check this condition before choosing the next branch.
    elif "SCC" in inner_labels:
        # Compute and store: mode.
        mode = "fireproof"
    else:
        # Compute and store: mode.
        mode = "standard"

    # MFD/FD/MSD rows are the one deliberate exception.  When the ordinary,
    # fireproof and acoustic variants all say the inside quantity is included
    # in the valve, there is no SCC/STA/SCC-TA label to identify the family.
    # Only then may the SCN expression distinguish the reference family.
    description = str(row.get("description", "")).casefold()
    # Compute and store: detail_text.
    detail_text = " ".join(
        str(row.get(name, ""))
        for name in ("formula_or_insulation_detail", "acoustic_fireproof_detail")
    ).casefold()
    # Compute and store: scn_expression.
    scn_expression = " ".join(
        expression
        for label, expression, _source in assignments
        if label == "SCN"
    )
    # Compute and store: compact_scn.
    compact_scn = re.sub(r"\s+", "", scn_expression).casefold()
    # Compute and store: is_mfd.
    is_mfd = any(marker in description for marker in ("fire damper", "van ngăn cháy", "mfd", "msd"))
    # Compute and store: mfd_included_variant.
    mfd_included_variant = (
        is_mfd
        and "bao gồm trong van" in detail_text
        and "SCN" in labels
        and not inner_labels
    )
    # Check this condition before choosing the next branch.
    if mfd_included_variant:
        # Compute and store: mode.
        mode = "fireproof" if "a/25" in compact_scn or "4a" in compact_scn else "acoustic"
    # Return the value produced by this step.
    return {
        "assignments": assignments,
        "source_assignments": source_assignments,
        "labels": labels,
        "inner_labels": inner_labels,
        "mode": mode,
        "has_outer": bool(labels & {"S", "SCN"}),
        "mapping_exception": "mfd_included_variant" if mfd_included_variant else "",
    }


# Purpose: Return the category used by the selected standard catalog row.
# Example output: "Tiêu âm"
def _formula_category(mode: str, target: str) -> str:
    """Return the category used by the selected standard catalog row."""
    # Check this condition before choosing the next branch.
    if target.upper() in {"SCC", "STA", "SCC/TA"}:
        # Return the value produced by this step.
        return "Chống cháy" if target.upper() == "SCC" else "Tiêu âm"
    # Return the value produced by this step.
    return "Cách nhiệt"


# Purpose: Normalize only cosmetic parentheses/spacing for the explicit exception.
# Example output: "W+H"
def _strip_parentheses(value: str) -> str:
    """Normalize only cosmetic parentheses/spacing for the explicit exception."""
    # Return the value produced by this step.
    return re.sub(r"[\s()]", "", _normalize_source(value)).casefold()


# Purpose: Identify a difference caused only by spaces or cosmetic parentheses.
# Example output: True
def _same_parentheses_only(left: str, right: str) -> bool:
    """Identify a difference caused only by spaces or cosmetic parentheses."""
    # Return the value produced by this step.
    return bool(left and right) and _strip_parentheses(left) == _strip_parentheses(right)


# Purpose: Build the variable context used by the row formula parser.
# Example output: {"W": 800.0, "H": 400.0, "K": 5.325}
def _formula_variables(row: dict[str, Any]) -> dict[str, float]:
    """Build the variable context used by the row formula parser."""
    # Return the value produced by this step.
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


# Purpose: Map formula variables to their actual Excel cells for one row.
# Example output: {"W": "F21", "H": "G21", "K": "Q21"}
def _formula_references(columns: dict[str, int], source_row: int) -> dict[str, str]:
    """Map formula variables to their actual Excel cells for one row."""
    # Compute and store: references.
    references = {}
    # Process each item in this input sequence.
    for name in ("W", "H", "W1", "H1", "W2", "H2", "L_M", "R_D", "a", "t", "Q", "K"):
        # Compute and store: column.
        column = columns.get(name)
        # Check this condition before choosing the next branch.
        if column is not None:
            # Compute and store: references[name].
            references[name] = f"{get_column_letter(column + 1)}{source_row}"
    # Return the value produced by this step.
    return references


# Purpose: Repair confirmed template typos while retaining the source text.
# Example output: ("(W+H)*K", "")
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
    # Compute and store: label_key.
    label_key = label.upper()
    # Compute and store: section_key.
    section_key = section.lower()
    # Compute and store: compact.
    compact = re.sub(r"\s+", "", expression).upper()
    # Check this condition before choosing the next branch.
    if product_type == "square_elbow" and label_key == "SCC/TA":
        # Compute and store: broken.
        broken = "(K*2*(((W+R)^2-R^2)+(2*(W+R)*H)+2*R*H))/1000000"
        # Check this condition before choosing the next branch.
        if compact == broken:
            # Compute and store: corrected.
            corrected = "(K*(2*((W+R)^2-R^2)+2*(W+R)*H+2*R*H))/1000000"
            # Return the value produced by this step.
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
        ("acoustic", "tee", "S_V3"): (
            "(2*PI()*((W''+W''/2)^2-(W''/2)^2)+2*PI()*H''*((W''+W''/2)+W''/2))*90/360",
            "Restored the acoustic tee S_V3 area grouping.",
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
    # Compute and store: repair.
    repair = repairs.get((section_key, product_type, label_key))
    # Check this condition before choosing the next branch.
    if repair:
        # Only apply the table to the known malformed patterns.  Valid future
        # formulas remain fully detail-driven instead of being overwritten.
        malformed = (
            compact.count("(") != compact.count(")")
            or 'H"' in compact
            or "Q(" in compact
        )
        # Check this condition before choosing the next branch.
        if malformed:
            # Return the value produced by this step.
            return repair
    # A missing closing parenthesis at the end of one assignment is safe to
    # repair for calculation: it cannot change the order of existing terms.
    # Keep the original expression untouched in ``detail_block`` and in the
    # formula comparison, so the contractor's source remains auditable.
    missing_closings = expression.count("(") - expression.count(")")
    # Check this condition before choosing the next branch.
    if missing_closings > 0:
        # Return the value produced by this step.
        return (
            f"{expression}{')' * missing_closings}",
            f"Auto-added {missing_closings} missing closing parenthesis(es) for calculation only.",
        )
    # Return the value produced by this step.
    return expression, ""


# Purpose: Expose current-row Excel cells for formulas that explicitly use E6-like refs.
# Example output: {"E21": 800.0, "F21": 400.0, "Q21": 5.325}
def _formula_cells(raw_row: pd.Series, source_row: int) -> dict[str, float]:
    """Expose current-row Excel cells for formulas that explicitly use E6-like refs."""
    # Compute and store: cells.
    cells = {}
    # Process each item in this input sequence.
    for column, value in enumerate(raw_row.tolist(), start=1):
        # Compute and store: cells[f'{get_column_letter(column)}{source_row}'].
        cells[f"{get_column_letter(column)}{source_row}"] = _number(value)
    # Return the value produced by this step.
    return cells


# Purpose: Parse each labelled formula independently so one bad block is isolated.
# Example output: ([("S", "...")], {"S": Node(...)}, {})
def _formula_block_assignments(
    text: str,
    *,
    product_type: str = "",
    repair_section: str = "",
    repair: bool = False,
) -> tuple[list[tuple[str, str]], dict[str, Any], dict[str, str]]:
    """Parse each labelled formula independently so one bad block is isolated."""
    # Compute and store: assignments.
    assignments = extract_assignments(text)
    # Compute and store: parsed.
    parsed = {}
    # Compute and store: errors.
    errors = {}
    # Compute and store: repaired.
    repaired = []
    # Process each item in this input sequence.
    for label, expression in assignments:
        # Compute and store: effective.
        effective = expression
        # Check this condition before choosing the next branch.
        if repair:
            # Compute and store: (effective, _note).
            effective, _note = _repair_known_detail_expression(
                label,
                expression,
                product_type,
                repair_section,
            )
        # Execute this expression for its result or side effect.
        repaired.append((label, effective))
        # Protect parsing/evaluation so bad workbook data becomes a controlled error.
        try:
            # Compute and store: parsed[label.upper()].
            parsed[label.upper()] = parse_formula(effective)
        except FormulaParseError as error:
            # Compute and store: errors[label.upper()].
            errors[label.upper()] = str(error)
    # Return the value produced by this step.
    return repaired, parsed, errors


# Purpose: Compare standard and contractor assignments one block at a time.
# Example output: {"status": "different", "blocks": [...]}
def _compare_formula_blocks(
    standard_block: str,
    detail_block: str,
    *,
    product_type: str = "",
    repair_section: str = "",
    standard_value: float | None = None,
    detail_value: float | None = None,
) -> dict[str, Any]:
    """Compare standard and contractor assignments one block at a time."""
    # Compute and store: (standard_assignments, standard_nodes, standard_errors).
    standard_assignments, standard_nodes, standard_errors = _formula_block_assignments(
        standard_block,
        product_type=product_type,
        repair_section=repair_section,
        repair=True,
    )
    # Compute and store: (detail_assignments, detail_nodes, detail_errors).
    detail_assignments, detail_nodes, detail_errors = _formula_block_assignments(detail_block)

    # Compute and store: detail_by_key.
    detail_by_key = {label.upper(): (label, expression) for label, expression in detail_assignments}
    # Compute and store: used_detail_keys.
    used_detail_keys: set[str] = set()

    # Purpose: Find the matching contractor block, including SCC/TA aliases.
    # Example output: ("S_V1", "2*W"), ("S", "K*S_V1")
    def detail_for(label: str):
        """Find the matching contractor block, including SCC/TA aliases."""
        # Compute and store: key.
        key = label.upper()
        # Compute and store: candidates.
        candidates = (key, "SCC/TA" if key == "STA" else "STA" if key == "SCC/TA" else "")
        # Process each item in this input sequence.
        for candidate in candidates:
            # Check this condition before choosing the next branch.
            if candidate and candidate in detail_by_key and candidate not in used_detail_keys:
                # Execute this expression for its result or side effect.
                used_detail_keys.add(candidate)
                # Return the value produced by this step.
                return candidate, detail_by_key[candidate]
        # Return the value produced by this step.
        return None, None

    # Compute and store: blocks.
    blocks = []
    # Process each item in this input sequence.
    for standard_label, standard_expression in standard_assignments:
        # Compute and store: standard_key.
        standard_key = standard_label.upper()
        # Compute and store: (detail_key, detail_item).
        detail_key, detail_item = detail_for(standard_label)
        # Compute and store: (detail_label, detail_expression).
        detail_label, detail_expression = detail_item or (standard_label, "")
        # Compute and store: error.
        error = standard_errors.get(standard_key, "")
        # Check this condition before choosing the next branch.
        if detail_key:
            # Compute and store: error.
            error = error or detail_errors.get(detail_key, "")

        # Compute and store: status.
        status = "unmatched"
        # Check this condition before choosing the next branch.
        if not error and detail_key:
            # Compute and store: detail_node.
            detail_node = detail_nodes.get(detail_key)
            # Compute and store: standard_node.
            standard_node = standard_nodes.get(standard_key)
            # Check this condition before choosing the next branch.
            if standard_node is not None and detail_node is not None:
                # Compute and store: status.
                status = "matched" if canonical(standard_node) == canonical(detail_node) else "unmatched"
        # Check this condition before choosing the next branch.
        if not detail_key:
            # Compute and store: error.
            error = error or "Formula block is missing from the contractor description."
        # Execute this expression for its result or side effect.
        blocks.append({
            "label": detail_label if detail_key else standard_label,
            "original_label": standard_label,
            "expression": detail_expression,
            "value": None,
            "status": status,
            "original_expression": standard_expression,
            "color": "#1565C0" if status == "matched" else "#B91C1C",
            "original_spans": [],
            "matched_components": [],
            "error": error,
        })

    # Process each item in this input sequence.
    for detail_label, detail_expression in detail_assignments:
        # Compute and store: detail_key.
        detail_key = detail_label.upper()
        # Check this condition before choosing the next branch.
        if detail_key in used_detail_keys:
            # Execute this statement as part of the current workflow.
            continue
        # Execute this expression for its result or side effect.
        blocks.append({
            "label": detail_label,
            "original_label": "",
            "expression": detail_expression,
            "value": None,
            "status": "unmatched",
            "original_expression": "",
            "color": "#B91C1C",
            "original_spans": [],
            "matched_components": [],
            "error": detail_errors.get(detail_key, "Formula block is not present in the standard formula."),
        })

    # Purpose: Sort intermediate S_V blocks before the final output block.
    # Example output: (0, 1, "S_V1")
    def block_order(block: dict[str, Any]) -> tuple[int, int, str]:
        """Sort intermediate S_V blocks before the final output block."""
        # Compute and store: label.
        label = str(block.get("label", "")).upper()
        # Compute and store: match.
        match = re.fullmatch(r"S_V(\d+)", label)
        # Check this condition before choosing the next branch.
        if match:
            # Return the value produced by this step.
            return (0, int(match.group(1)), label)
        # Return the value produced by this step.
        return (1, 0, label)

    # Keep the readable calculation order: intermediate areas first, output
    # quantity last.  The source workbook is not guaranteed to store them in
    # that order, especially for shared SCC/TA blocks.
    blocks.sort(key=block_order)

    # Check this condition before choosing the next branch.
    if not blocks:
        # Compute and store: status.
        status = "unverifiable"
    # Check this condition before choosing the next branch.
    elif all(block["status"] == "matched" for block in blocks):
        # Compute and store: status.
        status = "equivalent"
    # Check this condition before choosing the next branch.
    elif any(block["status"] == "matched" for block in blocks):
        # Compute and store: status.
        status = "different"
    else:
        # Compute and store: status.
        status = "unverifiable" if any(block.get("error") for block in blocks) else "different"

    # Compute and store: comparison.
    comparison = {
        "status": status,
        "original_formula_value": standard_value,
        "detail_formula_value": detail_value,
        "difference": (
            float(standard_value) - float(detail_value)
            if standard_value is not None and detail_value is not None
            else None
        ),
        "blocks": blocks,
    }
    # Compute and store: errors.
    errors = [block["error"] for block in blocks if block.get("error")]
    # Check this condition before choosing the next branch.
    if errors:
        # Compute and store: comparison['error'].
        comparison["error"] = "; ".join(dict.fromkeys(errors))
    # Return the value produced by this step.
    return comparison


# Purpose: Audit contractor text against the selected standard formula.
# Example output: {"S": {"matches": true, "value": 15.12, ...}}
def _row_formula_audits(
    row: dict[str, Any],
    raw_row: pd.Series,
    source_columns: dict[str, int],
    source_row: int,
    product_type: str = "",
    reference_catalog: dict[str, Any] | None = None,
    profile: dict[str, Any] | None = None,
) -> dict[str, dict[str, Any]]:
    """Audit contractor text against the selected standard formula.

    It matches each output label to the standard catalog.

    This function intentionally evaluates the contractor block only for the
    comparison metadata.  The returned ``value`` and ``expanded_node`` always
    come from the reference catalog, so the contractor formula can never
    become the result written to the workbook.
    """
    # Compute and store: audits.
    audits: dict[str, dict[str, Any]] = {}
    # Compute and store: profile.
    profile = profile or _formula_profile(row)
    # Compute and store: variables.
    variables = _formula_variables(row)
    # Compute and store: cells.
    cells = _formula_cells(raw_row, source_row)
    # Compute and store: references.
    references = _formula_references(source_columns, source_row)
    # Compute and store: description.
    description = str(row.get("description", ""))
    # Compute and store: mode.
    mode = profile["mode"]
    # Compute and store: source_assignments.
    source_assignments = profile.get("source_assignments", {})

    # Purpose: Translate a contractor label into catalog target/category metadata.
    # Example output: ("S", "Cách nhiệt", "S")
    def selected(label: str) -> tuple[str, str, str]:
        """Translate a contractor label into catalog target/category metadata."""
        # Compute and store: label_key.
        label_key = label.upper()
        # Check this condition before choosing the next branch.
        if label_key in {"S", "SCN"}:
            # Compute and store: standard_target.
            standard_target = "S" if label_key == "S" and mode == "standard" else "SCN"
        # Check this condition before choosing the next branch.
        elif label_key == "SCC":
            # Compute and store: standard_target.
            standard_target = "SCC"
        # Check this condition before choosing the next branch.
        elif label_key == "STA":
            # Compute and store: standard_target.
            standard_target = "STA"
        else:
            # Compute and store: standard_target.
            standard_target = "SCC/TA"
        # Return the value produced by this step.
        return standard_target, _formula_category(mode, standard_target), label_key

    # Purpose: Find, parse, and evaluate the authoritative catalog block.
    # Example output: (catalog_entry, {"value": 15.12}, "", "S", "")
    def reference_result(standard_target: str, category: str):
        """Find, parse, and evaluate the authoritative catalog block."""
        # A few reference rows use STA while rectangular/round products use
        # the shared SCC/TA label for the same inside quantity. Treat them as
        # equivalent targets during lookup, but keep the contractor label in
        # the audit display.
        target_candidates = [standard_target]
        # Check this condition before choosing the next branch.
        if standard_target == "STA":
            # Execute this expression for its result or side effect.
            target_candidates.append("SCC/TA")
        # Check this condition before choosing the next branch.
        elif standard_target == "SCC/TA":
            # Execute this expression for its result or side effect.
            target_candidates.append("STA")
        # Check this condition before choosing the next branch.
        if standard_target == "S":
            # Compute and store: family.
            family = "standard"
        # Check this condition before choosing the next branch.
        elif standard_target == "SCN":
            # Compute and store: family.
            family = "fireproof" if mode in {"standard", "fireproof"} else "acoustic"
        # Check this condition before choosing the next branch.
        elif standard_target == "SCC":
            # Compute and store: family.
            family = "fireproof" if mode in {"standard", "fireproof"} else "acoustic"
        else:
            # Compute and store: family.
            family = "acoustic"
        # Process each item in this input sequence.
        for candidate_target in target_candidates:
            # Compute and store: reference.
            reference = match_reference_formula(
                reference_catalog,
                product_type,
                candidate_target,
                description,
                category=category,
                family=family,
            )
            # Check this condition before choosing the next branch.
            if not reference:
                # Execute this statement as part of the current workflow.
                continue
            # Protect parsing/evaluation so bad workbook data becomes a controlled error.
            try:
                # Compute and store: assignments.
                assignments = extract_assignments(reference.get("formula_block", ""))
                # Compute and store: repair_section.
                repair_section = (
                    "acoustic" if mode in {"acoustic", "shared"}
                    else "fireproof" if candidate_target in {"SCN", "SCC"} and mode == "fireproof"
                    else "standard"
                )
                # Compute and store: repaired_assignments.
                repaired_assignments = []
                # Compute and store: correction_notes.
                correction_notes = []
                # Process each item in this input sequence.
                for assignment_label, assignment_expression in assignments:
                    # Compute and store: (corrected, note).
                    corrected, note = _repair_known_detail_expression(
                        assignment_label,
                        assignment_expression,
                        product_type,
                        repair_section,
                    )
                    # Execute this expression for its result or side effect.
                    repaired_assignments.append((assignment_label, corrected))
                    # Check this condition before choosing the next branch.
                    if note:
                        # Execute this expression for its result or side effect.
                        correction_notes.append(note)
                # Compute and store: result.
                result = formula_block_result(
                    repaired_assignments,
                    variables,
                    cells=cells,
                    references=references,
                    target=candidate_target,
                )
                # Return the value produced by this step.
                return reference, result, "", candidate_target, "; ".join(dict.fromkeys(correction_notes))
            except (FormulaParseError, FormulaEvaluationError, TypeError, ValueError) as error:
                # Keep trying the equivalent target if this row is malformed;
                # if no candidate works, return the first useful diagnostic.
                last_error = str(error)
                # Execute this statement as part of the current workflow.
                continue
        # Return the value produced by this step.
        return None, None, "No matching usable standard formula in the reference catalog.", standard_target, ""

    # Purpose: Parse/evaluate contractor text for comparison evidence only.
    # Example output: ({"value": 15.12, "target": "S"}, "")
    def contractor_result(source_name: str, target: str):
        """Parse/evaluate contractor text for comparison evidence only."""
        # Compute and store: assignments.
        assignments = source_assignments.get(source_name, [])
        # Protect parsing/evaluation so bad workbook data becomes a controlled error.
        try:
            # Compute and store: result.
            result = formula_block_result(
                assignments,
                variables,
                cells=cells,
                references=references,
                target=target,
            )
            # Return the value produced by this step.
            return result, ""
        except (FormulaParseError, FormulaEvaluationError, TypeError, ValueError) as error:
            # Return the value produced by this step.
            return None, str(error)

    # Prefer the most specific label in each output column.  SCC/TA is one
    # shared inside quantity and is therefore emitted once.
    labels = profile["labels"]
    # Compute and store: selected_labels.
    selected_labels: list[str] = []
    # Check this condition before choosing the next branch.
    if "SCN" in labels:
        # Execute this expression for its result or side effect.
        selected_labels.append("SCN")
    # Check this condition before choosing the next branch.
    elif "S" in labels:
        # Execute this expression for its result or side effect.
        selected_labels.append("S")
    # Check this condition before choosing the next branch.
    if "SCC/TA" in labels:
        # Execute this expression for its result or side effect.
        selected_labels.append("SCC/TA")
    else:
        # A three-quantity workbook can contain SCC and STA independently.
        # A legacy two-quantity workbook has no separate SCC column, but the
        # same source value is still audited once through the fallback below.
        if "SCC" in labels:
            # Execute this expression for its result or side effect.
            selected_labels.append("SCC")
        # Check this condition before choosing the next branch.
        if "STA" in labels:
            # Execute this expression for its result or side effect.
            selected_labels.append("STA")

    # Process each item in this input sequence.
    for label in selected_labels:
        # Compute and store: source_name.
        source_name = ""
        # Compute and store: expression.
        expression = ""
        # Process each item in this input sequence.
        for candidate_source in ("formula_or_insulation_detail", "acoustic_fireproof_detail"):
            # Compute and store: assignments.
            assignments = source_assignments.get(candidate_source, [])
            # Compute and store: match.
            match = next(
                ((candidate_label, candidate_expression) for candidate_label, candidate_expression in assignments
                 if candidate_label.upper() == label),
                None,
            )
            # Check this condition before choosing the next branch.
            if match:
                # Compute and store: source_name.
                source_name = candidate_source
                # Compute and store: expression.
                expression = match[1]
                # Execute this statement as part of the current workflow.
                break
        # Check this condition before choosing the next branch.
        if not source_name:
            # Execute this statement as part of the current workflow.
            continue

        # Compute and store: (standard_target, category, label_key).
        standard_target, category, label_key = selected(label)
        # Compute and store: target.
        target = _formula_targets(label, source_columns)[0]
        # Compute and store: detail_assignments.
        detail_assignments = source_assignments.get(source_name, [])
        # Compute and store: detail_block.
        detail_block = "\n".join(f"{item} = {value}" for item, value in detail_assignments)
        # Compute and store: (reference, standard, standard_error, resolved_target, standard_correction).
        reference, standard, standard_error, resolved_target, standard_correction = reference_result(standard_target, category)
        # Compute and store: standard_value.
        standard_value = standard.get("value") if standard else None
        # Compute and store: source_value_column.
        source_value_column = _audit_quantity_column(source_columns, target)
        # Compute and store: source_value_index.
        source_value_index = source_columns.get(source_value_column)
        # Compute and store: numeric_match.
        numeric_match = bool(
            standard
            and standard_value is not None
            and source_value_index is not None
            and _matches_original(raw_row.iloc[source_value_index], standard_value)
        )

        # The saved quantity is the verdict. Once it matches the standard
        # result, do not parse or compare the contractor formula at all.
        contractor = None
        # Compute and store: contractor_error.
        contractor_error = ""
        # Compute and store: contractor_value.
        contractor_value = None
        # Compute and store: reference_status.
        reference_status = ""
        # Compute and store: formula_comparison.
        formula_comparison = None
        # Check this condition before choosing the next branch.
        if not numeric_match:
            # Compute and store: (contractor, contractor_error).
            contractor, contractor_error = contractor_result(source_name, label)
            # Compute and store: contractor_value.
            contractor_value = contractor.get("value") if contractor else None
            # Compute and store: reference_status.
            reference_status = "unverifiable"
            # Check this condition before choosing the next branch.
            if standard and contractor:
                # Compute and store: contractor_node.
                contractor_node = contractor.get("expanded_nodes", {}).get(label_key)
                # Compute and store: standard_node.
                standard_node = standard.get("expanded_nodes", {}).get(resolved_target)
                # Compute and store: same_structure.
                same_structure = (
                    contractor_node is not None
                    and standard_node is not None
                    and canonical(contractor_node) == canonical(standard_node)
                )
                # Compute and store: parentheses_only.
                parentheses_only = _same_parentheses_only(
                    expression,
                    reference.get("formula", "") if reference else "",
                )
                # Check this condition before choosing the next branch.
                if same_structure or (
                    parentheses_only
                    and standard_value is not None
                    and contractor_value is not None
                    and abs(float(standard_value) - float(contractor_value)) <= 0.0005
                ):
                    # Compute and store: reference_status.
                    reference_status = "equivalent"
                else:
                    # Compute and store: reference_status.
                    reference_status = "different"

            # Compute and store: repair_section.
            repair_section = (
                "acoustic" if mode in {"acoustic", "shared"}
                else "fireproof" if resolved_target in {"SCN", "SCC"} and mode == "fireproof"
                else "standard"
            )
            # Compute and store: formula_comparison.
            formula_comparison = _compare_formula_blocks(
                reference.get("formula_block", "") if reference else "",
                detail_block,
                product_type=product_type,
                repair_section=repair_section,
                standard_value=standard_value,
                detail_value=contractor_value,
            )
        # Compute and store: audit.
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
        # Check this condition before choosing the next branch.
        if standard:
            # Compute and store: audit['canonical'].
            audit["canonical"] = standard.get("canonical")
            # Compute and store: audit['expanded_node'].
            audit["expanded_node"] = standard.get("expanded_nodes", {}).get(resolved_target)
            # Compute and store: audit['standard_expanded_node'].
            audit["standard_expanded_node"] = audit["expanded_node"]
        # Compute and store: audits[target].
        audits[target] = audit
    # Return the value produced by this step.
    return audits


# Purpose: Map source fields to their columns after result columns are inserted.
# Example output: {"W": "T21", "H": "U21", "K": "R21"}
def _final_formula_references(
    source_columns: dict[str, int],
    source_row: int,
    original_insertions: tuple[tuple[int, str, str, int], ...],
) -> dict[str, str]:
    """Map source fields to their columns after result columns are inserted."""
    # Compute and store: references.
    references = {}
    # Process each item in this input sequence.
    for name in ("W", "H", "W1", "H1", "W2", "H2", "L_M", "R_D", "a", "t", "Q", "K"):
        # Compute and store: source_column.
        source_column = source_columns.get(name)
        # Check this condition before choosing the next branch.
        if source_column is None:
            # Execute this statement as part of the current workflow.
            continue
        # Compute and store: final_column.
        final_column = source_column + 1 + sum(
            1 for original_column, _label, _value_key, _width in original_insertions
            if original_column < source_column
        )
        # Compute and store: references[name].
        references[name] = f"{get_column_letter(final_column)}{source_row}"
    # Return the value produced by this step.
    return references


# Purpose: Render a row's parsed detail formula for an inserted result cell.
# Example output: "=T21+U21"
def _render_output_formula(
    values: dict[str, Any],
    value_key: str,
    references: dict[str, str],
) -> str:
    """Render a row's parsed detail formula for an inserted result cell."""
    # Compute and store: audit.
    audit = values.get("formula_audits", {}).get(value_key)
    # Check this condition before choosing the next branch.
    if not audit and values.get("shared_target"):
        # Compute and store: audit.
        audit = values.get("formula_audits", {}).get(values["shared_target"])
    # Check this condition before choosing the next branch.
    if not audit:
        # Return the value produced by this step.
        return ""
    # Protect parsing/evaluation so bad workbook data becomes a controlled error.
    try:
        # The inserted Excel formula must be the standard formula.  A bad
        # contractor expression may still be shown in the audit table, but it
        # must never replace the authoritative output.
        node = audit.get("standard_expanded_node") or audit.get("expanded_node")
        # Check this condition before choosing the next branch.
        if node is None:
            # Return the value produced by this step.
            return ""
        # Return the value produced by this step.
        return f"={render_excel(node, references)}"
    except (FormulaParseError, FormulaEvaluationError, TypeError, ValueError):
        # Return the value produced by this step.
        return ""


# Purpose: Calculate rows from the standard formula catalog.
# Example output: ([calculated_rows], {20: {"S": 15.12}}, [])
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
    # Run Audit step 1: require the selected standard catalog as authority.
    calculated: list[dict[str, Any]] = []
    # Compute and store: output_values.
    output_values: dict[int, dict[str, Any]] = {}
    # Compute and store: errors.
    errors: list[dict[str, Any]] = []
    # Compute and store: source_columns.
    source_columns = detect_source_columns(data)
    # Compute and store: active_section.
    active_section: str | None = None

    # Process each item in this input sequence.
    for index, raw_row in data.iterrows():
        # Run Audit step 2: normalize the row and skip rows without geometry.
        heading_section = _section_from_heading(raw_row)
        # Check this condition before choosing the next branch.
        if heading_section:
            # Compute and store: active_section.
            active_section = heading_section
        # Compute and store: row.
        row = _normalized_row(raw_row, source_columns)
        # Check this condition before choosing the next branch.
        if active_section and _is_calculable(row) and not _has_section_formula_hint(row):
            # Compute and store: row['section'].
            row["section"] = active_section
        # Check this condition before choosing the next branch.
        if not _is_calculable(row):
            # Execute this statement as part of the current workflow.
            continue

        # Compute and store: source_row.
        source_row = int(index) + 1
        # Protect parsing/evaluation so bad workbook data becomes a controlled error.
        try:
            # Compute and store: has_k_components.
            has_k_components = any(_has_value(row.get(column)) for column in ("plan", "up_down", "vertical_axis"))
            # Compute and store: calculated_k.
            calculated_k = calculate_K(row) if has_k_components else _number(row.get("K"))
            # K is independent only when at least one of the three component
            # columns is populated.  Otherwise the contractor K is retained
            # as the input required by the standard formula.
            row["K_By_function"] = calculated_k if has_k_components else None
            # Compute and store: row['K_for_formula'].
            row["K_for_formula"] = calculated_k
            # Compute and store: profile.
            profile = _formula_profile(row)
            # Run Audit step 3: classify the row and resolve its output labels.
            calculation_section = {
                "standard": "standard",
                "fireproof": "fireproof",
                "acoustic": "acoustic",
                "shared": "acoustic",
            }[profile["mode"]]
            # Compute and store: product.
            product = ProductFormat.from_row(
                row,
                section=section or calculation_section,
                product_type=product_type,
            )
            # Built-in calculators are the safe fallback when a reference
            # catalog is not available. They are never fed contractor formula
            # text and are replaced by catalog values below when possible.
            result = product.calculate(row)
            # Compute and store: formula_audits.
            formula_audits = _row_formula_audits(
                row,
                raw_row,
                source_columns,
                source_row,
                product.product_type,
                reference_catalog,
                profile,
            )
            # Run Audit step 4: calculate only from matched standard formulas.
            quantities = {
                "S": result.R,
                "S_fireproof": 0.0,
                "S_inner": result.T,
            }
            # Process each item in this input sequence.
            for target, audit in formula_audits.items():
                # Check this condition before choosing the next branch.
                if audit["value"] is not None:
                    # Compute and store: quantities[target].
                    quantities[target] = audit["value"]

            # SCC/TA is one shared formula.  In a three-quantity workbook its
            # result is written to both inside-quantity columns so SCC and STA
            # cannot diverge just because the audit was attached to one of
            # them.
            shared_targets = {
                target
                for target, audit in formula_audits.items()
                if str(audit.get("label", "")).upper() == "SCC/TA"
            }
            # Compute and store: shared_target.
            shared_target = next(iter(shared_targets), "")
            # Check this condition before choosing the next branch.
            if shared_target:
                # Compute and store: shared_value.
                shared_value = quantities[shared_target]
                # Compute and store: quantities['S_fireproof'].
                quantities["S_fireproof"] = shared_value
                # Compute and store: quantities['S_inner'].
                quantities["S_inner"] = shared_value
            # Compute and store: result.
            result = result.__class__(
                section=result.section,
                product_type=result.product_type,
                R=quantities["S"],
                S=quantities["S_inner"],
                T=quantities["S_inner"],
                unit=result.unit,
            )
            # Compute and store: inputs.
            inputs = product.inputs_from(row)
            # Compute and store: result_dict.
            result_dict = result.to_dict()
            # Compute and store: result_dict['quantities']['S_fireproof'].
            result_dict["quantities"]["S_fireproof"] = quantities["S_fireproof"]
            # Compute and store: calculated_row.
            calculated_row = {
                "source_row": source_row,
                "description": product.description,
                "section": product.section,
                "product_type": product.product_type,
                "formula_mode": profile["mode"],
                "inputs": inputs.__dict__,
                "result": result_dict,
            }
            # Execute this expression for its result or side effect.
            calculated.append(calculated_row)
            # Compute and store: output_values[index].
            output_values[index] = {
                "K": calculated_k,
                "S": result.R,
                "S_inner": result.T,
                "S_fireproof": quantities["S_fireproof"],
                "shared_target": shared_target,
                "formula_audits": formula_audits,
                "matches": {
                    "K": _matches_original(row["K"], calculated_k) if has_k_components else True,
                    "S": _matches_original(
                        raw_row.iloc[source_columns[_audit_quantity_column(source_columns, "S")]],
                        result.R,
                    ),
                    "S_inner": _matches_original(
                        raw_row.iloc[source_columns[_audit_quantity_column(source_columns, "S_inner")]],
                        result.T,
                    ),
                    "S_fireproof": True,
                },
            }
            # Run Audit step 5: compare standard values with saved contractor values.
            if shared_target:
                # Process each item in this input sequence.
                for shared_column_target in ("S_fireproof", "S_inner"):
                    # Compute and store: source_name.
                    source_name = _audit_quantity_column(source_columns, shared_column_target)
                    # Compute and store: source_index.
                    source_index = source_columns.get(source_name) if source_name else None
                    # Compute and store: source_value.
                    source_value = raw_row.iloc[source_index] if source_index is not None else None
                    # Compute and store: output_values[index]['matches'][shared_column_target].
                    output_values[index]["matches"][shared_column_target] = (
                        True if not _has_value(source_value)
                        else _matches_original(source_value, quantities[shared_target])
                    )
            # Process each item in this input sequence.
            for target, audit in formula_audits.items():
                # Compute and store: original_column.
                original_column = _audit_quantity_column(source_columns, target)
                # Check this condition before choosing the next branch.
                if original_column in source_columns and audit["value"] is not None:
                    # Compute and store: numeric_match.
                    numeric_match = _matches_original(
                        raw_row.iloc[source_columns[original_column]],
                        audit["value"],
                    )
                # Check this condition before choosing the next branch.
                elif audit["value"] is None:
                    # If the catalog formula is malformed/unavailable, the
                    # built-in product calculation is still the numeric
                    # fallback. A saved value that matches that result is a
                    # match; formula text must not turn it red.
                    numeric_match = output_values[index]["matches"].get(target, True)
                else:
                    # Compute and store: numeric_match.
                    numeric_match = False
                # The quantity is the primary verdict. Formula parsing is
                # explanatory evidence only after the saved result is wrong;
                # equivalent results must not be marked red because the
                # contractor's description uses different grouping or has a
                # harmless formatting typo.
                audit["matches"] = numeric_match
                # Compute and store: output_values[index]['matches'][target].
                output_values[index]["matches"][target] = audit["matches"]
        except (TypeError, ValueError) as error:
            # Execute this expression for its result or side effect.
            errors.append(
                {
                    "source_row": source_row,
                    "description": row.get("description", ""),
                    "message": str(error),
                }
            )

    # Return the value produced by this step.
    return calculated, output_values, errors


# Purpose: Locate the output header row in the original worksheet.
# Example output: 1.0
def _header_row(data: pd.DataFrame) -> int:
    """Locate the output header row in the original worksheet."""
    # Process each item in this input sequence.
    for index, row in data.iterrows():
        # Compute and store: text.
        text = " ".join(str(value).lower() for value in row if _has_value(value))
        # Check this condition before choosing the next branch.
        if "total of quantities" in text and ("khối lượng cách nhiệt" in text or "detailed quantities" in text):
            # Return the value produced by this step.
            return int(index)
    # Return the value produced by this step.
    return 0


# Purpose: Insert calculated columns beside the detected quantity columns.
# Example output: DataFrame(columns=["K (tính bằng máy)", "S (cách nhiệt tính bằng máy)"])
def _insert_result_columns(
    data: pd.DataFrame,
    output_values: dict[int, dict[str, Any]],
) -> pd.DataFrame:
    """Insert calculated columns beside the detected quantity columns."""
    # Compute and store: result.
    result = data.copy()
    # Compute and store: header.
    header = _header_row(data)
    # Compute and store: source_columns.
    source_columns = detect_source_columns(data)

    # Purpose: Build one inserted result column aligned to original row indexes.
    # Example output: [None, "S (cách nhiệt tính bằng máy)", 15.12]
    def values(key: str, label: str) -> list[Any]:
        """Build one inserted result column aligned to original row indexes."""
        # Compute and store: output.
        output = [None] * len(data)
        # Compute and store: output[header].
        output[header] = label
        # Process each item in this input sequence.
        for index, row_values in output_values.items():
            # Compute and store: output[index].
            output[index] = row_values[key]
        # Return the value produced by this step.
        return output

    # Compute and store: outer_quantity_name.
    outer_quantity_name = _audit_quantity_column(source_columns, "S")
    # Compute and store: fireproof_quantity_name.
    fireproof_quantity_name = _audit_quantity_column(source_columns, "S_fireproof")
    # Compute and store: inner_quantity_name.
    inner_quantity_name = _audit_quantity_column(source_columns, "S_inner")
    # Compute and store: insertion_specs.
    insertion_specs = [
        (source_columns["K"], _K_OUTPUT, "K"),
        (source_columns[outer_quantity_name], _OUTER_S_OUTPUT, "S"),
    ]
    # Check this condition before choosing the next branch.
    if (
        fireproof_quantity_name in {"scc_quantity", "fireproof_quantity"}
        and source_columns.get(fireproof_quantity_name) not in {
            source_columns.get(outer_quantity_name),
            source_columns.get(inner_quantity_name),
        }
    ):
        # Execute this expression for its result or side effect.
        insertion_specs.append(
            (source_columns[fireproof_quantity_name], _FIREPROOF_S_OUTPUT, "S_fireproof")
        )
    # Execute this expression for its result or side effect.
    insertion_specs.append(
        (source_columns[inner_quantity_name], _INNER_S_OUTPUT, "S_inner")
    )

    # Compute and store: inserted_before.
    inserted_before = 0
    # Process each item in this input sequence.
    for original_column, label, value_key in sorted(insertion_specs):
        # Execute this expression for its result or side effect.
        result.insert(
            original_column + 1 + inserted_before,
            label,
            values(value_key, label),
        )
        # Update the accumulated value: inserted_before.
        inserted_before += 1
    # Return the value produced by this step.
    return result


# Purpose: Find LibreOffice, used to convert legacy .xls files without losing layout.
# Example output: "/usr/bin/soffice"
def _soffice_path() -> str:
    """Find LibreOffice, used to convert legacy .xls files without losing layout."""
    # Compute and store: executable.
    executable = shutil.which("soffice")
    # Check this condition before choosing the next branch.
    if executable:
        # Return the value produced by this step.
        return executable

    # Compute and store: bundled.
    bundled = (
        Path.home()
        / ".cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice"
    )
    # Check this condition before choosing the next branch.
    if bundled.is_file():
        # Return the value produced by this step.
        return str(bundled)
    # Stop with a clear validation error for invalid input.
    raise FileNotFoundError(
        "Không tìm thấy LibreOffice (soffice) để chuyển file .xls sang .xlsx."
    )


# Purpose: Return an xlsx source while keeping the original .xls file untouched.
# Example output: Path("/tmp/source.xlsx")
def _xlsx_source(source: Path, temporary_directory: Path) -> Path:
    """Return an xlsx source while keeping the original .xls file untouched."""
    # Check this condition before choosing the next branch.
    if source.suffix.lower() == ".xlsx":
        # Return the value produced by this step.
        return source

    # Execute this expression for its result or side effect.
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
    # Compute and store: converted.
    converted = temporary_directory / f"{source.stem}.xlsx"
    # Check this condition before choosing the next branch.
    if not converted.is_file():
        # Stop with a clear validation error for invalid input.
        raise FileNotFoundError(f"Không tạo được file xlsx tạm từ {source}.")
    # Return the value produced by this step.
    return converted


# Purpose: Move merge ranges around an inserted column.
def _shift_merged_ranges(worksheet, insert_at: int) -> None:
    """Move merge ranges around an inserted column.

    openpyxl moves cell values for ``insert_cols`` but does not update merged
    ranges.  Rebuilding the ranges here keeps the original header layout.
    """
    # Compute and store: ranges.
    ranges = [str(cell_range) for cell_range in worksheet.merged_cells.ranges]
    # Process each item in this input sequence.
    for cell_range in ranges:
        # Execute this expression for its result or side effect.
        worksheet.unmerge_cells(cell_range)

    # Process each item in this input sequence.
    for cell_range in ranges:
        # Compute and store: (min_col, min_row, max_col, max_row).
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        # Check this condition before choosing the next branch.
        if min_col >= insert_at:
            # Update the accumulated value: min_col.
            min_col += 1
            # Update the accumulated value: max_col.
            max_col += 1
        # Check this condition before choosing the next branch.
        elif max_col >= insert_at:
            # Update the accumulated value: max_col.
            max_col += 1
        # Execute this expression for its result or side effect.
        worksheet.merge_cells(
            start_row=min_row,
            start_column=min_col,
            end_row=max_row,
            end_column=max_col,
        )


# Purpose: Copy width and visibility settings to an inserted column.
def _copy_column_dimension(worksheet, source_column: int, target_column: int) -> None:
    """Copy width and visibility settings to an inserted column."""
    # Compute and store: source.
    source = worksheet.column_dimensions[get_column_letter(source_column)]
    # Compute and store: target.
    target = worksheet.column_dimensions[get_column_letter(target_column)]
    # Compute and store: target.width.
    target.width = source.width
    # Compute and store: target.hidden.
    target.hidden = source.hidden
    # Compute and store: target.bestFit.
    target.bestFit = source.bestFit
    # Compute and store: target.outlineLevel.
    target.outlineLevel = source.outlineLevel
    # Compute and store: target.collapsed.
    target.collapsed = source.collapsed


# Purpose: Copy the original column's cell styles to a newly inserted column.
def _copy_column_style(
    worksheet,
    source_column: int,
    target_column: int,
    red_text: bool = True,
    width: float | None = None,
) -> None:
    """Copy the original column's cell styles to a newly inserted column."""
    # Execute this expression for its result or side effect.
    _copy_column_dimension(worksheet, source_column, target_column)
    # Check this condition before choosing the next branch.
    if width is not None:
        # Compute and store: worksheet.column_dimensions[get_column_letter(target_column)].width.
        worksheet.column_dimensions[get_column_letter(target_column)].width = width
    # Process each item in this input sequence.
    for row in range(1, worksheet.max_row + 1):
        # Compute and store: source.
        source = worksheet.cell(row=row, column=source_column)
        # Compute and store: target.
        target = worksheet.cell(row=row, column=target_column)
        # Check this condition before choosing the next branch.
        if source.has_style:
            # Compute and store: target._style.
            target._style = copy(source._style)
        # Compute and store: target.number_format.
        target.number_format = source.number_format
        # Compute and store: target.protection.
        target.protection = copy(source.protection)
        # Check this condition before choosing the next branch.
        if red_text:
            # Compute and store: font.
            font = copy(target.font)
            # Compute and store: font.color.
            font.color = "FFFF0000"
            # Compute and store: target.font.
            target.font = font


# Purpose: Write a new header without breaking an existing merged cell.
def _set_new_header(worksheet, column: int, row: int, label: str) -> None:
    """Write a new header without breaking an existing merged cell."""
    # Compute and store: cell.
    cell = worksheet.cell(row=row, column=column)
    # Compute and store: inside_merged_header.
    inside_merged_header = any(
        merged.min_row <= row <= merged.max_row
        and merged.min_col <= column <= merged.max_col
        and (merged.min_row, merged.min_col) != (row, column)
        for merged in worksheet.merged_cells.ranges
    )
    # Check this condition before choosing the next branch.
    if isinstance(cell, MergedCell) or inside_merged_header:
        # In the reference template the top title band is merged across
        # columns S:AA.  Put a new result label in the second header row
        # directly; scanning farther down can land on a data cell and the
        # label may disappear when openpyxl saves the merged layout.
        candidate = worksheet.cell(row=row + 1, column=column)
        # Check this condition before choosing the next branch.
        if not isinstance(candidate, MergedCell):
            # Compute and store: cell.
            cell = candidate
        else:
            # Process each item in this input sequence.
            for candidate_row in range(row + 2, worksheet.max_row + 1):
                # Compute and store: candidate.
                candidate = worksheet.cell(row=candidate_row, column=column)
                # Check this condition before choosing the next branch.
                if not isinstance(candidate, MergedCell):
                    # Compute and store: cell.
                    cell = candidate
                    # Execute this statement as part of the current workflow.
                    break
    # Compute and store: cell.value.
    cell.value = label
    # Compute and store: font.
    font = copy(cell.font)
    # Compute and store: font.bold.
    font.bold = True
    # Compute and store: font.color.
    font.color = "FFFF0000"
    # Compute and store: cell.font.
    cell.font = font
    # Compute and store: alignment.
    alignment = copy(cell.alignment)
    # Compute and store: alignment.horizontal.
    alignment.horizontal = "center"
    # Compute and store: alignment.vertical.
    alignment.vertical = "center"
    # Compute and store: alignment.wrap_text.
    alignment.wrap_text = True
    # Compute and store: cell.alignment.
    cell.alignment = alignment
    # Compute and store: current_height.
    current_height = worksheet.row_dimensions[row].height or 0
    # Compute and store: worksheet.row_dimensions[row].height.
    worksheet.row_dimensions[row].height = max(current_height, 42)


# Purpose: Return a compact one-line representation of an original cell formula.
# Example output: "=E6+F6"
def _formula_text(value: Any) -> str:
    """Return a compact one-line representation of an original cell formula."""
    # Check this condition before choosing the next branch.
    if not isinstance(value, str) or not value.startswith("="):
        # Return the value produced by this step.
        return ""
    # Return the value produced by this step.
    return " ".join(value.replace("\r", " ").replace("\n", " ").split())


# Purpose: Use the contractor's dimension names when comparing source ASTs.
# Example output: Node(kind="name", value="W")
def _replace_formula_cells(node, cell_variables: dict[str, str]):
    """Use the contractor's dimension names when comparing source ASTs."""
    # Check this condition before choosing the next branch.
    if node.kind == "cell":
        # Compute and store: key.
        key = str(node.value).replace("$", "").upper()
        # Compute and store: name.
        name = cell_variables.get(key)
        # Return the value produced by this step.
        return type(node)("name", name) if name else node
    # Check this condition before choosing the next branch.
    if node.kind in ("unary", "binary", "call"):
        # Return the value produced by this step.
        return type(node)(
            node.kind,
            node.value,
            tuple(_replace_formula_cells(child, cell_variables) for child in node.children),
        )
    # Return the value produced by this step.
    return node


# Purpose: Index source formula subexpressions so detail blocks can be matched.
# Example output: {canonical_node: display_node}
def _formula_subtrees(node, display_node=None) -> dict[Any, Any]:
    """Index source formula subexpressions so detail blocks can be matched."""
    # Compute and store: result.
    result = {}

    # Purpose: Walk the AST while preserving the display tree's matching shape.
    # Example output: None (the walker records results in the enclosing mapping)
    def visit(current, display):
        """Walk the AST while preserving the display tree's matching shape."""
        # Execute this expression for its result or side effect.
        result.setdefault(canonical(current), display)
        # Process each item in this input sequence.
        for index, child in enumerate(current.children):
            # Execute this expression for its result or side effect.
            visit(child, display.children[index])

    # Execute this expression for its result or side effect.
    visit(node, display_node or node)
    # Return the value produced by this step.
    return result


# Purpose: Index AST subtrees by their source spans.
# Example output: {canonical_node: [{"start": 0, "end": 8}]}
def _formula_match_span_index_uncached(
    expression: str,
    cell_variables: dict[str, str],
) -> dict[Any, list[dict[str, int]]]:
    """Index AST subtrees by their source spans.

    The previous implementation parsed every pair of token boundaries.  A
    long formula therefore created thousands of temporary expressions for one
    comparison.  The full formula already contains every useful subtree, so
    parse once, render each existing subtree, and locate that text directly.
    """
    # Compute and store: source.
    source = expression.strip()
    # Compute and store: leading.
    leading = len(source) - len(source.lstrip("= "))
    # Compute and store: formula.
    formula = source[leading:]
    # Protect parsing/evaluation so bad workbook data becomes a controlled error.
    try:
        # Compute and store: display_root.
        display_root = parse_formula(formula)
    except FormulaParseError:
        # Return the value produced by this step.
        return {}
    # Compute and store: normalized_root.
    normalized_root = _replace_formula_cells(display_root, cell_variables)

    # Normalize operators for matching while preserving character positions.
    # All supported replacements are one-for-one; fall back to the original
    # text if an exotic Unicode normalization changes its length.
    normalized_formula = _normalize_source(formula)
    # Check this condition before choosing the next branch.
    if len(normalized_formula) != len(formula):
        # Compute and store: normalized_formula.
        normalized_formula = formula

    # Purpose: Remove spaces while retaining indexes back to the original text.
    # Example output: ("W+H", [0, 2, 3])
    def compact_with_positions(text: str) -> tuple[str, list[int]]:
        """Remove spaces while retaining indexes back to the original text."""
        # Compute and store: positions.
        positions = [index for index, character in enumerate(text) if not character.isspace()]
        # Return the value produced by this step.
        return "".join(text[index] for index in positions), positions

    # Compute and store: (compact_source, source_positions).
    compact_source, source_positions = compact_with_positions(normalized_formula)
    # Compute and store: next_positions.
    next_positions: dict[Any, int] = {}
    # Compute and store: indexed.
    indexed: dict[Any, list[dict[str, int]]] = {}

    # Purpose: Render one AST subtree with its own variable names for display.
    # Example output: "W+H"
    def render_display(node) -> str:
        """Render one AST subtree with its own variable names for display."""
        # Compute and store: references.
        references = {}

        # Purpose: Collect variable aliases required by ``render_excel``.
        # Example output: None (the walker fills the enclosing reference map)
        def collect(current):
            """Collect variable aliases required by ``render_excel``."""
            # Check this condition before choosing the next branch.
            if current.kind == "name":
                # Compute and store: name.
                name = str(current.value)
                # Compute and store: references[name].
                references[name] = name
                # Execute this expression for its result or side effect.
                references.setdefault(name.upper(), name)
                # Execute this expression for its result or side effect.
                references.setdefault(_variable_name(name), name)
            # Process each item in this input sequence.
            for child in current.children:
                # Execute this expression for its result or side effect.
                collect(child)

        # Execute this expression for its result or side effect.
        collect(node)
        # Return the value produced by this step.
        return render_excel(node, references)

    # Purpose: Locate a rendered subtree in compact source text.
    def find_target(target: str, start_at: int = 0) -> tuple[int, int] | None:
        """Locate a rendered subtree in compact source text."""
        # Contractor descriptions sometimes write implicit multiplication as
        # ``4a`` or ``2PI()`` while render_excel writes ``4*a``/``2*PI()``.
        # The AST key already proves the structure, so make only multiplication
        # separators optional when locating the equivalent source text.
        pattern = re.escape(target).replace(r"\*", r"\*?")
        # Compute and store: match.
        match = re.search(pattern, compact_source[start_at:])
        # Check this condition before choosing the next branch.
        if not match:
            # Return the value produced by this step.
            return None
        # Return the value produced by this step.
        return start_at + match.start(), start_at + match.end()

    # Purpose: Record source spans for this node and recursively for its children.
    # Example output: None (the walker appends source spans to the index)
    def visit(node, display_node):
        """Record source spans for this node and recursively for its children."""
        # Compute and store: key.
        key = canonical(node)
        # Protect parsing/evaluation so bad workbook data becomes a controlled error.
        try:
            # Compute and store: rendered.
            rendered = render_display(display_node)
        except (FormulaParseError, TypeError, ValueError):
            # Compute and store: rendered.
            rendered = ""
        # Compute and store: compact_target.
        compact_target = compact_with_positions(rendered)[0]
        # Check this condition before choosing the next branch.
        if compact_target:
            # Compute and store: start_at.
            start_at = next_positions.get(key, 0)
            # Compute and store: match.
            match = find_target(compact_target, start_at)
            # Check this condition before choosing the next branch.
            if match is None:
                # Commutative canonicalization can change traversal order.
                match = find_target(compact_target)
            # Check this condition before choosing the next branch.
            if match is not None:
                # Compute and store: (compact_start, compact_end).
                compact_start, compact_end = match
                # Check this condition before choosing the next branch.
                if compact_end <= len(source_positions):
                    # Compute and store: span_start.
                    span_start = source_positions[compact_start]
                    # Compute and store: span_end.
                    span_end = source_positions[compact_end - 1] + 1
                    # Check this condition before choosing the next branch.
                    if span_start > 0 and span_end < len(formula):
                        # Check this condition before choosing the next branch.
                        if formula[span_start - 1] == "(" and formula[span_end] == ")":
                            # Update the accumulated value: span_start.
                            span_start -= 1
                            # Update the accumulated value: span_end.
                            span_end += 1
                    # Execute this expression for its result or side effect.
                    indexed.setdefault(key, []).append({
                        "start": leading + span_start,
                        "end": leading + span_end,
                    })
                    # Compute and store: next_positions[key].
                    next_positions[key] = compact_end + 1
        # Process each item in this input sequence.
        for index, child in enumerate(node.children):
            # Execute this expression for its result or side effect.
            yield from visit(child, display_node.children[index])

    # Process each item in this input sequence.
    for _item in visit(normalized_root, display_root):
        # Execute this statement as part of the current workflow.
        pass

    # Process each item in this input sequence.
    for key, spans in indexed.items():
        # Execute this expression for its result or side effect.
        spans.sort(key=lambda span: (span["start"], -(span["end"] - span["start"])))
        # Compute and store: selected.
        selected = []
        # Process each item in this input sequence.
        for span in spans:
            # Check this condition before choosing the next branch.
            if any(span["start"] >= item["start"] and span["end"] <= item["end"] for item in selected):
                # Execute this statement as part of the current workflow.
                continue
            # Execute this expression for its result or side effect.
            selected.append(span)
        # Compute and store: indexed[key].
        indexed[key] = selected
    # Return the value produced by this step.
    return indexed


@lru_cache(maxsize=256)
# Purpose: Cache expensive AST-to-source span indexing by formula and aliases.
# Example output: {canonical_node: [{"start": 0, "end": 8}]}
def _formula_match_span_index_cached(
    expression: str,
    cell_variable_items: tuple[tuple[str, str], ...],
) -> dict[Any, list[dict[str, int]]]:
    """Cache expensive AST-to-source span indexing by formula and aliases."""
    # Return the value produced by this step.
    return _formula_match_span_index_uncached(expression, dict(cell_variable_items))


# Purpose: Return cached token-aligned spans for one formula and variable map.
# Example output: {canonical_node: [{"start": 0, "end": 8}]}
def _formula_match_span_index(
    expression: str,
    cell_variables: dict[str, str],
) -> dict[Any, list[dict[str, int]]]:
    """Return cached token-aligned spans for one formula and variable map."""
    # Return the value produced by this step.
    return _formula_match_span_index_cached(
        expression,
        tuple(sorted(cell_variables.items())),
    )


# Purpose: Find source-text ranges whose parsed formula matches a canonical block.
# Example output: {canonical_node: [{"start": 0, "end": 8}]}
def _formula_match_spans(
    expression: str,
    target: Any,
    cell_variables: dict[str, str],
) -> list[dict[str, int]]:
    """Find source-text ranges whose parsed formula matches a canonical block."""
    # Return the value produced by this step.
    return _formula_match_span_index(expression, cell_variables).get(target, [])


# Purpose: Count nodes in a subtree for AST matching and ranking.
# Example output: 1.0
def _node_size(node) -> int:
    """Count nodes in a subtree for AST matching and ranking."""
    # Return the value produced by this step.
    return 1 + sum(_node_size(child) for child in node.children)


# Purpose: Count normalized variables so similar branches can be paired safely.
# Example output: Counter({"W": 1, "H": 1})
def _ast_variable_signature(node) -> Counter[str]:
    """Count normalized variables so similar branches can be paired safely."""
    # Compute and store: signature.
    signature: Counter[str] = Counter()
    # Check this condition before choosing the next branch.
    if node.kind == "name":
        # The standard catalog may use W'/W'' while contractor formulas
        # normally use W1/W2.  Compare their parser aliases, not the raw
        # spelling, otherwise S_V1 and S_V2 can be paired with each other.
        signature[_variable_name(str(node.value)).upper()] += 1
    # Process each item in this input sequence.
    for child in node.children:
        # Execute this expression for its result or side effect.
        signature.update(_ast_variable_signature(child))
    # Return the value produced by this step.
    return signature


# Purpose: Normalize safe linear arithmetic such as ``4a`` and ``2a+2a``.
def _linear_signature(node) -> dict[str, Fraction] | None:
    """Normalize safe linear arithmetic such as ``4a`` and ``2a+2a``."""
    # Check this condition before choosing the next branch.
    if node.kind == "number":
        # Return the value produced by this step.
        return {"#constant": Fraction(str(node.value))}
    # Check this condition before choosing the next branch.
    if node.kind == "name":
        # Return the value produced by this step.
        return {_variable_name(str(node.value)).upper(): Fraction(1)}
    # Check this condition before choosing the next branch.
    if node.kind == "cell":
        # Return the value produced by this step.
        return {str(node.value).upper(): Fraction(1)}
    # Check this condition before choosing the next branch.
    if node.kind == "unary":
        # Compute and store: child.
        child = _linear_signature(node.children[0])
        # Check this condition before choosing the next branch.
        if child is None:
            # Return the value produced by this step.
            return None
        # Compute and store: factor.
        factor = Fraction(1) if node.value == "+" else Fraction(-1)
        # Return the value produced by this step.
        return {key: value * factor for key, value in child.items() if value * factor}
    # Check this condition before choosing the next branch.
    if node.kind != "binary":
        # Return the value produced by this step.
        return None

    # Compute and store: left.
    left = _linear_signature(node.children[0])
    # Compute and store: right.
    right = _linear_signature(node.children[1])
    # Check this condition before choosing the next branch.
    if node.value in ("+", "-"):
        # Check this condition before choosing the next branch.
        if left is None or right is None:
            # Return the value produced by this step.
            return None
        # Compute and store: sign.
        sign = Fraction(1) if node.value == "+" else Fraction(-1)
        # Compute and store: result.
        result = dict(left)
        # Process each item in this input sequence.
        for key, value in right.items():
            # Compute and store: result[key].
            result[key] = result.get(key, Fraction(0)) + sign * value
        # Return the value produced by this step.
        return {key: value for key, value in result.items() if value}
    # Check this condition before choosing the next branch.
    if node.value == "*":
        # Check this condition before choosing the next branch.
        if left is not None and set(left) <= {"#constant"} and right is not None:
            # Compute and store: factor.
            factor = left.get("#constant", Fraction(0))
            # Return the value produced by this step.
            return {key: value * factor for key, value in right.items() if value * factor}
        # Check this condition before choosing the next branch.
        if right is not None and set(right) <= {"#constant"} and left is not None:
            # Compute and store: factor.
            factor = right.get("#constant", Fraction(0))
            # Return the value produced by this step.
            return {key: value * factor for key, value in left.items() if value * factor}
        # Return the value produced by this step.
        return None
    # Check this condition before choosing the next branch.
    if node.value == "/":
        # Check this condition before choosing the next branch.
        if left is None or right is None or set(right) != {"#constant"}:
            # Return the value produced by this step.
            return None
        # Compute and store: divisor.
        divisor = right["#constant"]
        # Check this condition before choosing the next branch.
        if divisor == 0:
            # Return the value produced by this step.
            return None
        # Return the value produced by this step.
        return {key: value / divisor for key, value in left.items() if value / divisor}
    # Return the value produced by this step.
    return None


# Purpose: Compare two AST nodes with safe linear-arithmetic normalization.
# Example output: True
def _same_ast_node(left, right) -> bool:
    """Compare two AST nodes with safe linear-arithmetic normalization."""
    # Compute and store: left_linear.
    left_linear = _linear_signature(left)
    # Compute and store: right_linear.
    right_linear = _linear_signature(right)
    # Check this condition before choosing the next branch.
    if left_linear is not None and left_linear == right_linear:
        # Return the value produced by this step.
        return True
    # Check this condition before choosing the next branch.
    if left.kind != right.kind or len(left.children) != len(right.children):
        # Return the value produced by this step.
        return False
    # Check this condition before choosing the next branch.
    if left.kind in ("number", "name", "cell"):
        # Return the value produced by this step.
        return canonical(left) == canonical(right)
    # Check this condition before choosing the next branch.
    if left.kind in ("unary", "binary") and left.value != right.value:
        # Return the value produced by this step.
        return False
    # Check this condition before choosing the next branch.
    if left.kind == "call" and str(left.value).upper() != str(right.value).upper():
        # Return the value produced by this step.
        return False
    # Return the value produced by this step.
    return all(_same_ast_node(a, b) for a, b in zip(left.children, right.children))


# Purpose: Return independent mismatched branches, without reporting nested duplicates.
# Example output: [("root[0]", "W/2", "(W/2)^2")]
def _ast_difference_frontier(
    left,
    right,
    left_path: tuple[int, ...] = (),
    right_path: tuple[int, ...] | None = None,
) -> list[tuple[tuple[int, ...], tuple[int, ...], Any, Any]]:
    """Return independent mismatched branches, without reporting nested duplicates."""
    # Compute and store: right_path.
    right_path = left_path if right_path is None else right_path
    # canonical() already flattens and sorts commutative +/* groups.  A
    # contractor may therefore write S_V2 + S_V1 while the standard writes
    # S_V1 + S_V2; that is the same formula, not an AST error to display.
    if _same_ast_node(left, right) or canonical(left) == canonical(right):
        # Return the value produced by this step.
        return []
    # Check this condition before choosing the next branch.
    if left.kind != right.kind:
        # Return the value produced by this step.
        return [(left_path, right_path, left, right)]
    # Check this condition before choosing the next branch.
    if left.kind in ("number", "name", "cell"):
        # Return the value produced by this step.
        return [(left_path, right_path, left, right)]
    # Check this condition before choosing the next branch.
    if left.kind in ("unary", "binary") and left.value != right.value:
        # Return the value produced by this step.
        return [(left_path, right_path, left, right)]
    # Check this condition before choosing the next branch.
    if left.kind == "call" and str(left.value).upper() != str(right.value).upper():
        # Return the value produced by this step.
        return [(left_path, right_path, left, right)]

    # Check this condition before choosing the next branch.
    if left.kind == "binary" and left.value in ("+", "*"):
        # Purpose: Flatten one commutative operator for order-independent pairing.
        # Example output: [(path, child_node)]
        def commutative_children(node, node_path):
            """Flatten one commutative operator for order-independent pairing."""
            # Check this condition before choosing the next branch.
            if node.kind == "binary" and node.value == left.value:
                # Process each item in this input sequence.
                for index, child in enumerate(node.children):
                    # Execute this expression for its result or side effect.
                    yield from commutative_children(child, node_path + (index,))
            else:
                # Execute this expression for its result or side effect.
                yield node_path, node

        # Compute and store: unmatched_right.
        unmatched_right = list(commutative_children(right, right_path))
        # Compute and store: pairs.
        pairs = []
        # Compute and store: unmatched_left.
        unmatched_left = []
        # Process each item in this input sequence.
        for left_child_path, left_child in commutative_children(left, left_path):
            # Compute and store: match_index.
            match_index = next(
                (
                    position
                    for position, (_right_child_path, right_child) in enumerate(unmatched_right)
                    if canonical(left_child) == canonical(right_child)
                ),
                None,
            )
            # Check this condition before choosing the next branch.
            if match_index is None:
                # Execute this expression for its result or side effect.
                unmatched_left.append((left_child_path, left_child))
                # Execute this statement as part of the current workflow.
                continue
            # Compute and store: (right_child_path, right_child).
            right_child_path, right_child = unmatched_right.pop(match_index)
            # Execute this expression for its result or side effect.
            pairs.append((left_child_path, left_child, right_child_path, right_child))

        # Pair the remaining branches by structural similarity.  This keeps a
        # changed S_V3 aligned with S_V3 even when S_V1/S_V2 were reordered.
        while unmatched_left and unmatched_right:
            # Compute and store: (left_child_path, left_child).
            left_child_path, left_child = unmatched_left.pop(0)
            # Compute and store: (match_position, (right_child_path, right_child)).
            match_position, (right_child_path, right_child) = max(
                enumerate(unmatched_right),
                key=lambda item: _ast_similarity(left_child, item[1][1]),
            )
            # Execute this expression for its result or side effect.
            unmatched_right.pop(match_position)
            # Execute this expression for its result or side effect.
            pairs.append((left_child_path, left_child, right_child_path, right_child))

        # Compute and store: differences.
        differences = []
        # Process each item in this input sequence.
        for left_child_path, left_child, right_child_path, right_child in pairs:
            # Execute this expression for its result or side effect.
            differences.extend(
                _ast_difference_frontier(
                    left_child,
                    right_child,
                    left_child_path,
                    right_child_path,
                )
            )
        # Process each item in this input sequence.
        for left_child_path, left_child in unmatched_left:
            # Execute this expression for its result or side effect.
            differences.append((left_child_path, right_path, left_child, right))
        # Process each item in this input sequence.
        for right_child_path, right_child in unmatched_right:
            # Execute this expression for its result or side effect.
            differences.append((left_path, right_child_path, left, right_child))
        # Return the value produced by this step.
        return differences

    # Check this condition before choosing the next branch.
    if len(left.children) != len(right.children):
        # Return the value produced by this step.
        return [(left_path, right_path, left, right)]

    # Compute and store: differences.
    differences = []
    # Process each item in this input sequence.
    for index, (left_child, right_child) in enumerate(zip(left.children, right.children)):
        # Execute this expression for its result or side effect.
        differences.extend(
            _ast_difference_frontier(
                left_child,
                right_child,
                left_path + (index,),
                right_path + (index,),
            )
        )
    # Return the value produced by this step.
    return differences or [(left_path, right_path, left, right)]


# Purpose: Render an AST node as readable Excel-style text for the audit report.
# Example output: "W+H"
def _ast_render(node) -> str:
    """Render an AST node as readable Excel-style text for the audit report."""
    # Compute and store: references.
    references = {}

    # Purpose: Collect names used to keep rendered variables human-readable.
    # Example output: None (the walker fills the enclosing reference map)
    def collect(current):
        """Collect names used to keep rendered variables human-readable."""
        # Check this condition before choosing the next branch.
        if current.kind == "name":
            # Compute and store: name.
            name = str(current.value)
            # Compute and store: references[name].
            references[name] = name
            # Execute this expression for its result or side effect.
            references.setdefault(name.upper(), name)
            # Execute this expression for its result or side effect.
            references.setdefault(_variable_name(name), name)
        # Process each item in this input sequence.
        for child in current.children:
            # Execute this expression for its result or side effect.
            collect(child)

    # Execute this expression for its result or side effect.
    collect(node)
    # Protect parsing/evaluation so bad workbook data becomes a controlled error.
    try:
        # Return the value produced by this step.
        return render_excel(node, references)
    except FormulaParseError:
        # Return the value produced by this step.
        return str(node.value) if node.value is not None else node.kind


# Purpose: Yield each AST subtree together with its parallel display subtree.
# Example output: [(Node(...), Node(...))]
def _ast_subtrees(node, display_node=None):
    """Yield each AST subtree together with its parallel display subtree."""
    # Compute and store: display_node.
    display_node = display_node or node
    # Execute this expression for its result or side effect.
    yield node, display_node
    # Process each item in this input sequence.
    for index, child in enumerate(node.children):
        # Execute this expression for its result or side effect.
        yield from _ast_subtrees(child, display_node.children[index])


# Purpose: Score structural similarity used only when exact subtree matching fails.
# Example output: 1.0
def _ast_similarity(left, right) -> int:
    """Score structural similarity used only when exact subtree matching fails."""
    # Check this condition before choosing the next branch.
    if left.kind != right.kind:
        # Return the value produced by this step.
        return -20
    # Check this condition before choosing the next branch.
    if left.kind in ("number", "name", "cell"):
        # Return the value produced by this step.
        return 8 if canonical(left) == canonical(right) else -4
    # Check this condition before choosing the next branch.
    if left.kind in ("unary", "binary"):
        # Compute and store: score.
        score = 7 if left.value == right.value else -7
    # Check this condition before choosing the next branch.
    elif left.kind == "call":
        # Compute and store: score.
        score = 7 if str(left.value).upper() == str(right.value).upper() else -7
    else:
        # Compute and store: score.
        score = 0
    # Return the value produced by this step.
    return score + sum(_ast_similarity(a, b) for a, b in zip(left.children, right.children))


# Purpose: Choose the closest source branch while preferring matching variables.
# Example output: (Node(...), Node(...))
def _best_ast_subtree(target, source_root, source_display_root):
    """Choose the closest source branch while preferring matching variables."""
    # Compute and store: target_signature.
    target_signature = _ast_variable_signature(target)
    # Compute and store: target_size.
    target_size = _node_size(target)
    # Compute and store: best.
    best = None
    # Process each item in this input sequence.
    for candidate, display in _ast_subtrees(source_root, source_display_root):
        # An exact canonical subtree is the only unambiguous pairing.  Return
        # it before the heuristic ranking can select a structurally similar
        # branch that happens to occur earlier in the contractor formula.
        if canonical(candidate) == canonical(target):
            # Return the value produced by this step.
            return candidate, display
        # Compute and store: candidate_signature.
        candidate_signature = _ast_variable_signature(candidate)
        # Compute and store: exact_variables.
        exact_variables = candidate_signature == target_signature
        # Compute and store: overlap.
        overlap = sum((candidate_signature & target_signature).values())
        # Compute and store: extra_variables.
        extra_variables = sum((candidate_signature - target_signature).values())
        # Compute and store: similarity.
        similarity = _ast_similarity(candidate, target)
        # Compute and store: size_penalty.
        size_penalty = abs(_node_size(candidate) - target_size)
        # A final S expression contains all S_V blocks, so raw overlap alone
        # can make the whole root look more similar than the intended block.
        # Prefer the candidate with the target variables and the fewest
        # unrelated variables before considering structural similarity.
        rank = (
            exact_variables,
            -extra_variables,
            -size_penalty,
            overlap,
            similarity,
            -_node_size(candidate),
        )
        # Check this condition before choosing the next branch.
        if best is None or rank > best[0]:
            # Compute and store: best.
            best = rank, candidate, display
    # Return the value produced by this step.
    return best[1:] if best else (None, None)


# Purpose: Convert AST mismatch paths into compact report-ready differences.
# Example output: [{"path": "root[0]", "original": "W/2^2", "standard": "(W/2)^2"}]
def _ast_difference_payloads(original, standard, original_display, standard_display) -> list[dict[str, Any]]:
    """Convert AST mismatch paths into compact report-ready differences."""
    # Compute and store: payloads.
    payloads = []
    # Process each item in this input sequence.
    for left_path, right_path, original_node, standard_node in _ast_difference_frontier(original, standard):
        # Compute and store: original_part.
        original_part = original_display
        # Compute and store: standard_part.
        standard_part = standard_display
        # Process each item in this input sequence.
        for index in left_path:
            # Compute and store: original_part.
            original_part = original_part.children[index]
        # Process each item in this input sequence.
        for index in right_path:
            # Compute and store: standard_part.
            standard_part = standard_part.children[index]
        # Execute this expression for its result or side effect.
        payloads.append({
            "path": "root" + "".join(f"[{index}]" for index in left_path),
            "original": _ast_render(original_part),
            "standard": _ast_render(standard_part),
            "original_kind": original_node.kind,
            "standard_kind": standard_node.kind,
        })
    # Return the value produced by this step.
    return payloads


# Purpose: Find a source subtree with exactly the same canonical structure.
# Example output: (Node(...), Node(...))
def _find_exact_ast_subtree(target, root, display_root):
    """Find a source subtree with exactly the same canonical structure."""
    # Process each item in this input sequence.
    for candidate, display in _ast_subtrees(root, display_root):
        # Check this condition before choosing the next branch.
        if canonical(candidate) == canonical(target):
            # Return the value produced by this step.
            return candidate, display
    # Return the value produced by this step.
    return None, None


# Purpose: Attach only the independent AST-difference branches to standard blocks.
def _attach_ast_differences(
    standard_comparison: dict[str, Any],
    source_root,
    source_display_root,
    standard_expanded_node,
) -> None:
    """Attach only the independent AST-difference branches to standard blocks."""
    # Compute and store: output_labels.
    output_labels = {"S", "SCN", "SCC", "SCC/TA", "STA"}
    # Process each item in this input sequence.
    for block in standard_comparison.get("blocks", []):
        # Compute and store: expression.
        expression = str(block.get("original_expression", ""))
        # Compute and store: label.
        label = str(block.get("original_label") or block.get("label") or "").upper()
        # Check this condition before choosing the next branch.
        if not expression and label not in output_labels:
            # Execute this statement as part of the current workflow.
            continue
        # Protect parsing/evaluation so bad workbook data becomes a controlled error.
        try:
            # Check this condition before choosing the next branch.
            if label in output_labels and standard_expanded_node is not None:
                # Compute and store: standard_display.
                standard_display = standard_expanded_node
            else:
                # Compute and store: standard_display.
                standard_display = parse_formula(expression)
            # Compute and store: standard_node.
            standard_node = standard_display
            # Compute and store: (source_node, source_display).
            source_node, source_display = (
                (source_root, source_display_root)
                if label in output_labels
                else _best_ast_subtree(standard_node, source_root, source_display_root)
            )
            # Check this condition before choosing the next branch.
            if source_node is None:
                # Execute this statement as part of the current workflow.
                continue
            # Compute and store: differences.
            differences = _ast_difference_payloads(
                source_node,
                standard_node,
                source_display,
                standard_display,
            )
            # Check this condition before choosing the next branch.
            if differences:
                # Keep the paired context for this block.  The source formula
                # may be a large final S expression, but an intermediate
                # block such as S_V2 must show only its matched contractor
                # subtree beside the standard S_V2 expression.
                block["ast_original_expression"] = _ast_render(source_display)
                # Compute and store: block['ast_standard_expression'].
                block["ast_standard_expression"] = _ast_render(standard_display)
                # Compute and store: block['ast_differences'].
                block["ast_differences"] = differences

            # For a final S/SCN/SCC output, also retain the complete
            # intermediate branch that contains the mismatch.  The UI can
            # then tint only S_Vx inside the long output formula instead of
            # coloring the entire output expression.
            if label in output_labels:
                # Compute and store: branch_originals.
                branch_originals = []
                # Compute and store: branch_standards.
                branch_standards = []
                # Process each item in this input sequence.
                for candidate_block in standard_comparison.get("blocks", []):
                    # Compute and store: candidate_label.
                    candidate_label = str(candidate_block.get("label", "")).upper()
                    # Check this condition before choosing the next branch.
                    if not re.fullmatch(r"S_V\d+", candidate_label):
                        # Execute this statement as part of the current workflow.
                        continue
                    # Check this condition before choosing the next branch.
                    if not candidate_block.get("ast_differences"):
                        # Execute this statement as part of the current workflow.
                        continue
                    # Compute and store: candidate_expression.
                    candidate_expression = str(candidate_block.get("original_expression", ""))
                    # Check this condition before choosing the next branch.
                    if not candidate_expression:
                        # Execute this statement as part of the current workflow.
                        continue
                    # Compute and store: candidate_node.
                    candidate_node = parse_formula(candidate_expression)
                    # Compute and store: (_standard_branch, standard_branch_display).
                    _standard_branch, standard_branch_display = _find_exact_ast_subtree(
                        candidate_node,
                        standard_display,
                        standard_display,
                    )
                    # Compute and store: (_source_branch, source_branch_display).
                    _source_branch, source_branch_display = _best_ast_subtree(
                        candidate_node,
                        source_root,
                        source_display_root,
                    )
                    # Check this condition before choosing the next branch.
                    if standard_branch_display is not None and source_branch_display is not None:
                        # Execute this expression for its result or side effect.
                        branch_standards.append(_ast_render(standard_branch_display))
                        # Execute this expression for its result or side effect.
                        branch_originals.append(_ast_render(source_branch_display))
                # Check this condition before choosing the next branch.
                if branch_originals:
                    # Compute and store: block['ast_block_original_expressions'].
                    block["ast_block_original_expressions"] = branch_originals
                    # Compute and store: block['ast_block_standard_expressions'].
                    block["ast_block_standard_expressions"] = branch_standards
        except (FormulaParseError, TypeError, ValueError, IndexError):
            # Execute this statement as part of the current workflow.
            continue


# Purpose: Return maximal meaningful subexpressions shared by source and detail.
# Example output: [{"color": "FF1565C0", "original_expression": "W+H"}]
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
    # Compute and store: candidates.
    candidates = []
    # Compute and store: source_index.
    source_index = _formula_match_span_index(source_formula, cell_variables)
    # Compute and store: detail_index.
    detail_index = _formula_match_span_index(detail_expression, {})

    # Purpose: Collect matching AST components before removing nested duplicates.
    # Example output: None (the walker only appends matching components)
    def visit(node):
        """Collect matching AST components before removing nested duplicates."""
        # Compute and store: key.
        key = canonical(node)
        # Compute and store: source_spans.
        source_spans = source_index.get(key, [])
        # Compute and store: detail_spans.
        detail_spans = detail_index.get(key, [])
        # Check this condition before choosing the next branch.
        if _node_size(node) >= 3 and key in source_subtrees and source_spans and detail_spans:
            # Execute this expression for its result or side effect.
            candidates.append((min(span["start"] for span in detail_spans), key, source_spans, detail_spans))
            # Keep visiting children too.  A matching parent does not mean
            # every nested matching term should be hidden from the report.
        for child in node.children:
            # Execute this expression for its result or side effect.
            visit(child)

    # Execute this expression for its result or side effect.
    visit(detail_node)
    # The same normalized subtree can occur more than once. Keep one component
    # per distinct structure; its span arrays preserve every occurrence.
    unique = {}
    # Process each item in this input sequence.
    for start, key, source_spans, detail_spans in candidates:
        # Execute this expression for its result or side effect.
        unique.setdefault(key, (start, source_spans, detail_spans))

    # Compute and store: used_colors.
    used_colors = used_colors if used_colors is not None else set()
    # Compute and store: components.
    components = []
    # Process each item in this input sequence.
    for index, (_key, (start, source_spans, detail_spans)) in enumerate(sorted(unique.items(), key=lambda item: item[1][0])):
        # Compute and store: preferred.
        preferred = preferred_color if preferred_color and len(unique) == 1 else ""
        # Compute and store: color.
        color = preferred if preferred and preferred not in used_colors else next(
            (candidate for candidate in _FORMULA_MATCH_COLORS if candidate not in used_colors),
            _FORMULA_MATCH_COLORS[index % len(_FORMULA_MATCH_COLORS)],
        )
        # Execute this expression for its result or side effect.
        used_colors.add(color)
        # Execute this expression for its result or side effect.
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
    # Return the value produced by this step.
    return components


# Purpose: Compare formula structure and row results without guessing unmatched terms.
# Example output: {"status": "equivalent", "blocks": [...]}
def _formula_comparison(
    original_formula: str,
    detail_text: str,
    inputs: dict[str, Any],
    raw_row: pd.Series,
    source_columns: dict[str, int],
    source_row: int,
    value_worksheet=None,
    standard_comparison: dict[str, Any] | None = None,
    standard_expanded_node=None,
) -> dict[str, Any]:
    """Compare formula structure and row results without guessing unmatched terms."""
    # Check this condition before choosing the next branch.
    if not original_formula or not detail_text:
        # Return the value produced by this step.
        return {"status": "unverifiable", "error": "Original formula or detail formula is missing.", "blocks": []}

    # Compute and store: cells.
    cells = _formula_cells(raw_row, source_row)
    # Check this condition before choosing the next branch.
    if value_worksheet is not None:
        # Process each item in this input sequence.
        for reference in re.findall(r"\$?[A-Za-z]{1,3}\$?\d+", original_formula):
            # Compute and store: address.
            address = reference.replace("$", "").upper()
            # Compute and store: value.
            value = value_worksheet[address].value
            # Check this condition before choosing the next branch.
            if _has_value(value):
                # Compute and store: cells[address].
                cells[address] = _number(value)
    # Compute and store: references.
    references = _formula_references(source_columns, source_row)
    # Compute and store: cell_variables.
    cell_variables = {
        address.replace("$", "").upper(): name
        for name, address in references.items()
    }
    # Compute and store: input_fields.
    input_fields = {
        "W": "w", "H": "h", "W1": "w1", "H1": "h1",
        "W2": "w2", "H2": "h2", "L_M": "L_M", "R_D": "R_D",
        "a": "a", "t": "t", "Q": "Q", "K": "K",
    }
    # Compute and store: variables.
    variables = {
        name: float(inputs.get(field, inputs.get(name)) or 0)
        for name, field in input_fields.items()
    }
    # Protect parsing/evaluation so bad workbook data becomes a controlled error.
    try:
        # Compute and store: original_node.
        original_node = parse_formula(original_formula.lstrip("= ").strip())
        # Compute and store: original_value.
        original_value = float(evaluate(original_node, {}, cells))
        # Compute and store: normalized_original.
        normalized_original = _replace_formula_cells(original_node, cell_variables)
        # Compute and store: source_subtrees.
        source_subtrees = _formula_subtrees(normalized_original, original_node)
    except (FormulaParseError, FormulaEvaluationError, TypeError, ValueError) as error:
        # Return the value produced by this step.
        return {
            "status": "unverifiable",
            "error": str(error),
            "original_formula_value": None,
            "blocks": [],
        }

    # Compute and store: assignments.
    assignments = extract_assignments(detail_text)
    # Compute and store: parsed_assignments.
    parsed_assignments = []
    # Compute and store: parse_errors.
    parse_errors = {}
    # Compute and store: parsed_nodes.
    parsed_nodes = {}
    # Process each item in this input sequence.
    for label, expression in assignments:
        # Compute and store: key.
        key = label.upper()
        # Protect parsing/evaluation so bad workbook data becomes a controlled error.
        try:
            # Compute and store: parsed_nodes[key].
            parsed_nodes[key] = parse_formula(expression)
            # Execute this expression for its result or side effect.
            parsed_assignments.append((label, expression))
        except FormulaParseError as error:
            # Compute and store: parse_errors[key].
            parse_errors[key] = str(error)

    # Compute and store: output_names.
    output_names = {"S", "SCN", "SCC", "SCC/TA", "STA"}

    # The normal case has one valid output target.  Resolve the complete
    # block once; the old per-assignment loop reparsed and reevaluated the
    # same expressions, then formula_block_result did it all again.
    output_assignments = [
        label.upper() for label, _expression in parsed_assignments
        if label.upper() in output_names
    ]
    # Compute and store: fast_block.
    fast_block = None
    # Check this condition before choosing the next branch.
    if not parse_errors and len(output_assignments) == 1:
        # Protect parsing/evaluation so bad workbook data becomes a controlled error.
        try:
            # Compute and store: fast_block.
            fast_block = formula_block_result(
                parsed_assignments,
                variables,
                cells=cells,
                target=output_assignments[0],
            )
        except (FormulaParseError, FormulaEvaluationError, TypeError, ValueError):
            # Keep the independent-resolution fallback for malformed or
            # partially unresolved contractor blocks.
            fast_block = None

    # Resolve each valid assignment independently; a malformed or unresolved
    # sibling must not hide formula blocks that can still be compared. Do not
    # use formula_block_result for intermediate labels: its target is reserved
    # for the final output labels above.
    if fast_block is not None:
        # Compute and store: resolved.
        resolved = parsed_assignments
        # Compute and store: resolved_values.
        resolved_values = fast_block["values"]
        # Compute and store: unresolved_errors.
        unresolved_errors = {}
        # Compute and store: label_results.
        label_results = {output_assignments[0]: fast_block}
    else:
        # Compute and store: resolved.
        resolved = []
        # Compute and store: resolved_values.
        resolved_values = {}
        # Compute and store: unresolved_errors.
        unresolved_errors = {}
        # Compute and store: pending.
        pending = list(parsed_assignments)
        # Repeat while this condition remains true.
        while pending:
            # Compute and store: remaining.
            remaining = []
            # Compute and store: progress.
            progress = False
            # Process each item in this input sequence.
            for assignment in pending:
                # Compute and store: (label, expression).
                label, expression = assignment
                # Compute and store: key.
                key = label.upper()
                # Protect parsing/evaluation so bad workbook data becomes a controlled error.
                try:
                    # Compute and store: single.
                    single = formula_result(
                        expression,
                        {**variables, **resolved_values},
                        cells=cells,
                    )
                except (FormulaParseError, FormulaEvaluationError, TypeError, ValueError) as error:
                    # Compute and store: unresolved_errors[key].
                    unresolved_errors[key] = str(error)
                    # Execute this expression for its result or side effect.
                    remaining.append(assignment)
                    # Execute this statement as part of the current workflow.
                    continue
                # Execute this expression for its result or side effect.
                resolved.append(assignment)
                # Compute and store: resolved_values[key].
                resolved_values[key] = single["value"]
                # Execute this expression for its result or side effect.
                unresolved_errors.pop(key, None)
                # Compute and store: progress.
                progress = True
            # Check this condition before choosing the next branch.
            if not progress:
                # Execute this statement as part of the current workflow.
                break
            # Compute and store: pending.
            pending = remaining

        # Compute and store: label_results.
        label_results = {}
        # Check this condition before choosing the next branch.
        if resolved:
            # Process each item in this input sequence.
            for label, _expression in resolved:
                # Check this condition before choosing the next branch.
                if label.upper() not in output_names:
                    # Execute this statement as part of the current workflow.
                    continue
                # Compute and store: label_results[label.upper()].
                label_results[label.upper()] = formula_block_result(
                    resolved,
                    variables,
                    cells=cells,
                    target=label,
                )

    # Compute and store: output_label.
    output_label = next(
        (label.upper() for label, _expression in assignments
         if label.upper() in output_names and label.upper() in label_results),
        None,
    )
    # Compute and store: detail_value.
    detail_value = label_results[output_label]["value"] if output_label else None
    # Check this condition before choosing the next branch.
    if detail_value is None:
        # Compute and store: status.
        status = "unverifiable"
    # Check this condition before choosing the next branch.
    elif canonical(normalized_original) == canonical(label_results[output_label]["expanded_nodes"][output_label]):
        # Compute and store: status.
        status = "equivalent"
    # Check this condition before choosing the next branch.
    elif abs(original_value - float(detail_value)) <= 0.0005:
        # Compute and store: status.
        status = "same_value_structure_diff"
    else:
        # Compute and store: status.
        status = "different"

    # Compute and store: blocks.
    blocks = []
    # Compute and store: output_block_labels.
    output_block_labels = {"S", "SCN", "SCC", "SCC/TA", "STA"}
    # Process each item in this input sequence.
    for label, expression in assignments:
        # Compute and store: key.
        key = label.upper()
        # Compute and store: is_output_block.
        is_output_block = key in output_block_labels
        # Compute and store: node.
        node = parsed_nodes.get(key)
        # Compute and store: result.
        result = label_results.get(key)
        # Compute and store: expanded.
        expanded = result["expanded_nodes"].get(key, node) if result else node
        # Compute and store: normalized.
        normalized = _replace_formula_cells(expanded, cell_variables) if expanded else None
        # Compute and store: block_key.
        block_key = canonical(normalized) if normalized else None
        # Compute and store: matched.
        matched = source_subtrees.get(block_key) if block_key is not None else None
        # Compute and store: color.
        color = "#1565C0" if matched is not None else "#B91C1C"
        # Compute and store: components.
        components = _matched_formula_components(
            original_formula,
            expression,
            source_subtrees,
            expanded,
            cell_variables,
            preferred_color="#1565C0",
            used_colors=set(),
        ) if expanded is not None and is_output_block else []
        # Process each item in this input sequence.
        for component in components:
            # Compute and store: component['color'].
            component["color"] = "#1565C0" if matched is not None else "#B91C1C"
        # Execute this expression for its result or side effect.
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

    # Compute and store: errors.
    errors = [*parse_errors.values(), *unresolved_errors.values()]
    # Compute and store: comparison.
    comparison = {
        "status": status,
        "original_formula": original_formula,
        "original_formula_value": original_value,
        "detail_formula_value": detail_value,
        "difference": original_value - float(detail_value) if detail_value is not None else None,
        "blocks": blocks,
    }
    # Check this condition before choosing the next branch.
    if standard_comparison:
        # Execute this expression for its result or side effect.
        _attach_ast_differences(
            standard_comparison,
            normalized_original,
            original_node,
            standard_expanded_node,
        )
    # Check this condition before choosing the next branch.
    if errors:
        # Compute and store: comparison['error'].
        comparison["error"] = "; ".join(dict.fromkeys(errors))
    # Return the value produced by this step.
    return comparison


_CELL_REFERENCE_RE = re.compile(r"\$?[A-Za-z]{1,3}\$?\d+")
_REFERENCE_COLORS = (
    "FF1565C0",
    "FF2E7D32",
    "FFEF6C00",
    "FF6A1B9A",
    "FF00838F",
)


# Purpose: Make cell references visually distinct without changing the source text.
# Example output: CellRichText("Original: =E6+F6\nDetail: W+H")
def _rich_formula_text(original: str, detail: str) -> CellRichText:
    """Make cell references visually distinct without changing the source text."""
    # Compute and store: text.
    text = f"Original: {original or '(not available)'}\nDetail:   {detail or '(not available)'}"
    # Compute and store: rich.
    rich = CellRichText()
    # Compute and store: cursor.
    cursor = 0
    # Compute and store: colors.
    colors: dict[str, str] = {}
    # Process each item in this input sequence.
    for match in _CELL_REFERENCE_RE.finditer(text):
        # Check this condition before choosing the next branch.
        if match.start() > cursor:
            # Execute this expression for its result or side effect.
            rich.append(text[cursor:match.start()])
        # Compute and store: reference.
        reference = match.group(0)
        # Compute and store: key.
        key = reference.replace("$", "").upper()
        # Compute and store: color.
        color = colors.setdefault(key, _REFERENCE_COLORS[len(colors) % len(_REFERENCE_COLORS)])
        # Execute this expression for its result or side effect.
        rich.append(TextBlock(InlineFont(color=color, b=True), reference))
        # Compute and store: cursor.
        cursor = match.end()
    # Check this condition before choosing the next branch.
    if cursor < len(text):
        # Execute this expression for its result or side effect.
        rich.append(text[cursor:])
    # Return the value produced by this step.
    return rich


# Purpose: Append only mismatched formula rows with original/detail text.
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
    # Compute and store: formula_rows.
    formula_rows = []
    # Process each item in this input sequence.
    for raw_index, values in output_values.items():
        # Compute and store: audits.
        audits = values.get("formula_audits", {})
        # Compute and store: mismatched.
        mismatched = [audit for audit in audits.values() if not audit.get("matches", False)]
        # Check this condition before choosing the next branch.
        if not mismatched:
            # Execute this statement as part of the current workflow.
            continue
        # Compute and store: source_row.
        source_row = int(raw_index) + 1
        # Compute and store: original_parts.
        original_parts = []
        # Compute and store: detail_parts.
        detail_parts = []
        # Compute and store: original_values.
        original_values = []
        # Compute and store: calculated_values.
        calculated_values = []
        # Compute and store: differences.
        differences = []
        # Process each item in this input sequence.
        for target, audit in audits.items():
            # Check this condition before choosing the next branch.
            if audit.get("matches", False) and audit.get("reference_status") != "different":
                # Execute this statement as part of the current workflow.
                continue
            # Compute and store: source_name.
            source_name = _audit_quantity_column(source_columns, target)
            # Compute and store: original_formula.
            original_formula = ""
            # Compute and store: original_value.
            original_value: Any = None
            # Check this condition before choosing the next branch.
            if source_name in source_columns:
                # Compute and store: source_column.
                source_column = source_columns[source_name]
                # Compute and store: original_formula.
                original_formula = _formula_text(
                    formula_worksheet.cell(source_row, source_column + 1).value
                )
                # Compute and store: original_value.
                original_value = value_worksheet.cell(source_row, source_column + 1).value
            # The comparison is standard-vs-contractor.  Keep the source
            # Excel formula/value below as an additional trace when present.
            original_parts.append(f"{target}: {audit.get('reference_formula') or original_formula}")
            # Execute this expression for its result or side effect.
            detail_parts.append(f"{target}: {audit.get('expression', '')}")
            # Execute this expression for its result or side effect.
            original_values.append(f"{target}={_text_value(original_value)}")
            # Compute and store: calculated.
            calculated = audit.get("value")
            # Execute this expression for its result or side effect.
            calculated_values.append(
                f"{target}={audit.get('error') or _text_value(calculated)}"
            )
            # Execute this expression for its result or side effect.
            differences.append(
                f"{target}={_number(original_value) - _number(calculated)}"
                if calculated is not None and not audit.get("error")
                else f"{target}=error"
            )
        # Compute and store: item_no.
        item_no = _text_value(data.iloc[raw_index].iloc[source_columns["item_no"]])
        # Execute this expression for its result or side effect.
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

    # Check this condition before choosing the next branch.
    if not formula_rows:
        # Return the value produced by this step.
        return

    # Compute and store: start_row.
    start_row = worksheet.max_row + 3
    # Compute and store: headers.
    headers = ["STT", "Name", "Formula audit", "Original value", "Calculated value", "Difference"]
    # Compute and store: thin_gray.
    thin_gray = Side(style="thin", color="FFD9D9D9")
    # Compute and store: header_fill.
    header_fill = PatternFill("solid", fgColor="FFFFE5E5")
    # Process each item in this input sequence.
    for column, header in enumerate(headers, start=1):
        # Compute and store: cell.
        cell = worksheet.cell(start_row, column, header)
        # Compute and store: cell.font.
        cell.font = Font(bold=True, color="FF000000")
        # Compute and store: cell.fill.
        cell.fill = header_fill
        # Compute and store: cell.alignment.
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        # Compute and store: cell.border.
        cell.border = Border(top=thin_gray, bottom=thin_gray, left=thin_gray, right=thin_gray)

    # Process each item in this input sequence.
    for row_offset, values in enumerate(formula_rows, start=1):
        # Compute and store: row_number.
        row_number = start_row + row_offset
        # Process each item in this input sequence.
        for column, value in enumerate(values, start=1):
            # Compute and store: cell.
            cell = worksheet.cell(row_number, column, value)
            # Compute and store: cell.border.
            cell.border = Border(top=thin_gray, bottom=thin_gray, left=thin_gray, right=thin_gray)
            # Compute and store: cell.alignment.
            cell.alignment = Alignment(
                horizontal="left" if column in (2, 3) else "center",
                vertical="center",
                wrap_text=False,
            )
        # Compute and store: worksheet.row_dimensions[row_number].height.
        worksheet.row_dimensions[row_number].height = 38

    # Compute and store: widths.
    widths = {1: 8, 2: 42, 3: 160, 4: 22, 5: 22, 6: 22}
    # Process each item in this input sequence.
    for column, width in widths.items():
        # Compute and store: worksheet.column_dimensions[get_column_letter(column)].width.
        worksheet.column_dimensions[get_column_letter(column)].width = width
    # Compute and store: worksheet.auto_filter.ref.
    worksheet.auto_filter.ref = f"A{start_row}:F{start_row + len(formula_rows)}"


# Purpose: Write a compact sidecar for the desktop table preview.
# Example output: Path("/tmp/calculated_preview.tsv")
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
    # Compute and store: preview.
    preview = destination.with_name(f"{destination.stem}_preview.tsv")
    # Compute and store: source_columns.
    source_columns = source_columns or detect_source_columns(data)
    # Compute and store: headers.
    headers = [
        "source_row", "item_no", "description", "unit", "unit_duplicate",
        "section", "product_type",
        "w", "h", "w1", "h1", "w2", "h2", "l_m", "r_d", "a", "t", "q",
        "plan", "up_down", "vertical_axis", "formula_detail", "acoustic_fireproof_detail", "note",
        "k_by_function", "contractor_k", "calculated_k", "k_match",
        "contractor_s", "calculated_s", "s_match", "contractor_fireproof_s", "calculated_fireproof_s", "fireproof_match", "contractor_inner_s",
        "calculated_inner_s", "inner_match",
        "contractor_scc", "calculated_scc", "scc_match", "contractor_sta", "calculated_sta", "sta_match", "quantity_layout", "shared_target",
        "formula_audits", "input_references",
    ]
    # Open the required resource for this operation.
    with preview.open("w", encoding="utf-8", newline="") as stream:
        # Compute and store: writer.
        writer = csv.writer(stream, delimiter="\t", lineterminator="\n")
        # Execute this expression for its result or side effect.
        writer.writerow(headers)
        # Process each item in this input sequence.
        for item in calculated:
            # Compute and store: source_row.
            source_row = int(item["source_row"])
            # Compute and store: raw_row.
            raw_row = data.iloc[source_row - 1]
            # Compute and store: values.
            values = output_values[source_row - 1]
            # Compute and store: result.
            result = item["result"]["quantities"]
            # Compute and store: inputs.
            inputs = item["inputs"]
            # Compute and store: formula_audits.
            formula_audits = []
            # Process each item in this input sequence.
            for target, audit in values.get("formula_audits", {}).items():
                # Compute and store: source_name.
                source_name = _audit_quantity_column(source_columns, target)
                # Compute and store: source_column.
                source_column = source_columns.get(source_name) if source_name else None
                # Compute and store: original_formula.
                original_formula = ""
                # Compute and store: original_value.
                original_value = None
                # Check this condition before choosing the next branch.
                if source_column is not None and formula_worksheet is not None:
                    # Compute and store: original_formula.
                    original_formula = _formula_text(
                        formula_worksheet.cell(source_row, source_column + 1).value
                    )
                # Check this condition before choosing the next branch.
                if source_column is not None and value_worksheet is not None:
                    # Compute and store: original_value.
                    original_value = value_worksheet.cell(source_row, source_column + 1).value
                # Compute and store: calculated_value.
                calculated_value = audit.get("value")
                # Compute and store: comparison.
                comparison = audit.get("formula_comparison")
                # Compute and store: source_comparison.
                source_comparison = None
                # Check this condition before choosing the next branch.
                if not audit.get("matches"):
                    # Compute and store: source_comparison.
                    source_comparison = _formula_comparison(
                        original_formula,
                        audit.get("detail_block", ""),
                        inputs,
                        raw_row,
                        source_columns,
                        source_row,
                        value_worksheet,
                        standard_comparison=comparison,
                        standard_expanded_node=audit.get("standard_expanded_node"),
                    )
                # Compute and store: comparison_values.
                comparison_values = source_comparison or {}
                # Execute this expression for its result or side effect.
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
                    "source_formula_comparison": source_comparison,
                })

            # ``formula_audits`` normally comes from a parsed detail formula.
            # K is different: it is calculated from Plan + Up/Down + Vertical
            # Axis, so it has no S/SCN/SCC/STA assignment to add to that list.
            # Also keep a visible fallback for an output mismatch whose detail
            # text could not be parsed.  Otherwise the comparison table can
            # show a red cell while Formula error review appears empty.
            audit_by_target = {str(audit["target"]): audit for audit in formula_audits}

            # Purpose: Read an original Excel formula from the selected source row.
            # Example output: "=E21+F21"
            def source_formula(column_name: str) -> str:
                """Read an original Excel formula from the selected source row."""
                # Compute and store: column.
                column = source_columns.get(column_name)
                # Check this condition before choosing the next branch.
                if column is None or formula_worksheet is None:
                    # Return the value produced by this step.
                    return ""
                # Return the value produced by this step.
                return _formula_text(formula_worksheet.cell(source_row, column + 1).value)

            # Purpose: Combine detail cells used as the contractor formula source.
            # Example output: "S=K*(W+H)"
            def detail_text(*column_names: str) -> str:
                """Combine detail cells used as the contractor formula source."""
                # Return the value produced by this step.
                return "\n".join(
                    text
                    for text in (_text_value(raw_row.iloc[source_columns[name]]) for name in column_names if name in source_columns)
                    if text
                )

            # Compute and store: outer_quantity_name.
            outer_quantity_name = _audit_quantity_column(source_columns, "S")
            # Compute and store: fireproof_quantity_name.
            fireproof_quantity_name = _audit_quantity_column(source_columns, "S_fireproof")
            # Compute and store: inner_quantity_name.
            inner_quantity_name = _audit_quantity_column(source_columns, "S_inner")
            # Compute and store: separate_fireproof.
            separate_fireproof = (
                fireproof_quantity_name in {"scc_quantity", "fireproof_quantity"}
                and source_columns.get(fireproof_quantity_name) not in {
                    source_columns.get(outer_quantity_name),
                    source_columns.get(inner_quantity_name),
                }
            )
            # Compute and store: mismatch_specs.
            mismatch_specs = (
                ("K", "K", "K", values.get("K"), source_columns.get("K"), "K = PLAN + UP-DOWN + VERTICAL AXIS"),
                ("S", "S", outer_quantity_name, result["R"], source_columns.get(outer_quantity_name), detail_text("formula_or_insulation_detail")),
                ("S_fireproof", "SCC", fireproof_quantity_name, result.get("S_fireproof", 0.0), source_columns.get(fireproof_quantity_name), detail_text("acoustic_fireproof_detail", "formula_or_insulation_detail")),
                ("S_inner", "STA / SCC-TA", inner_quantity_name, result["T"], source_columns.get(inner_quantity_name), detail_text("acoustic_fireproof_detail", "formula_or_insulation_detail")),
            )
            # Process each item in this input sequence.
            for target, label, source_name, calculated_value, source_column, detail in mismatch_specs:
                # SCC/TA is one shared contractor formula.  In a workbook
                # with separate SCC and STA columns, the audit already belongs
                # to the shared target (normally S_fireproof); do not create a
                # second metadata-less fallback for the other output column.
                shared_target = values.get("shared_target")
                # Check this condition before choosing the next branch.
                if (
                    shared_target
                    and target in {"S_fireproof", "S_inner"}
                    and target != shared_target
                ):
                    # Execute this statement as part of the current workflow.
                    continue
                # Check this condition before choosing the next branch.
                if values["matches"].get(target, True) or target in audit_by_target:
                    # Execute this statement as part of the current workflow.
                    continue
                # Compute and store: original_value.
                original_value = (
                    _number(raw_row.iloc[source_column])
                    if source_column is not None
                    else 0.0
                )
                # Compute and store: original.
                original = source_formula(source_name)
                # Check this condition before choosing the next branch.
                if not original:
                    # Compute and store: original.
                    original = f"{label} source value = {_text_value(raw_row.iloc[source_column])}" if source_column is not None else "Original value not available"
                # Compute and store: detail_block.
                detail_block = detail or f"{label} calculated value = {_number(calculated_value)}"
                # Compute and store: audit_by_target[target].
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

            # Compute and store: formula_audits.
            formula_audits = [audit_by_target[target] for target in ("K", "S", "S_fireproof", "S_inner") if target in audit_by_target]
            # Execute this expression for its result or side effect.
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
                _number(raw_row.iloc[source_columns[outer_quantity_name]]),
                result["R"],
                values["matches"]["S"],
                _number(raw_row.iloc[source_columns[fireproof_quantity_name]]) if separate_fireproof else _number(raw_row.iloc[source_columns[inner_quantity_name]]),
                result.get("S_fireproof", 0.0) if separate_fireproof else result["T"],
                values["matches"].get("S_fireproof", True) if separate_fireproof else values["matches"]["S_inner"],
                _number(raw_row.iloc[source_columns[inner_quantity_name]]),
                result["T"],
                values["matches"]["S_inner"],
                _number(raw_row.iloc[source_columns[fireproof_quantity_name]]) if separate_fireproof else _number(raw_row.iloc[source_columns[inner_quantity_name]]),
                result.get("S_fireproof", 0.0) if separate_fireproof else result["T"],
                values["matches"].get("S_fireproof", True) if separate_fireproof else values["matches"]["S_inner"],
                _number(raw_row.iloc[source_columns[inner_quantity_name]]),
                result["T"],
                values["matches"]["S_inner"],
                "three" if separate_fireproof else "two",
                values.get("shared_target", ""),
                json.dumps(formula_audits, ensure_ascii=False, separators=(",", ":")),
                json.dumps(_formula_references(source_columns, source_row), ensure_ascii=False, separators=(",", ":")),
            ])
    # Return the value produced by this step.
    return preview


# Purpose: Export a full workbook while preserving the source workbook layout.
# Example output: {"output": "/tmp/calculated.xlsx", "calculated_count": 15, "errors": []}
def export_calculated_workbook(
    input_path: str | Path,
    output_path: str | Path | None = None,
    sheet_name: str = "02",
    reference_path: str | Path | None = None,
    reference_sheet: str | None = None,
    reference_catalog: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Export a full workbook while preserving the source workbook layout."""
    # Run Audit step 6: validate paths and load the standard catalog.
    source = Path(input_path).expanduser().resolve()
    # Compute and store: destination.
    destination = (
        Path(output_path).expanduser().resolve()
        if output_path is not None
        else source.with_name(f"{source.stem}_calculated.xlsx")
    )
    # Check this condition before choosing the next branch.
    if source == destination:
        # Stop with a clear validation error for invalid input.
        raise ValueError("The output workbook must be different from the input workbook.")

    # Check this condition before choosing the next branch.
    if destination.suffix.lower() != ".xlsx":
        # Stop with a clear validation error for invalid input.
        raise ValueError("File kết quả phải có đuôi .xlsx để giữ nguyên định dạng.")

    # Check this condition before choosing the next branch.
    if reference_catalog is None:
        # Compute and store: reference_catalog.
        reference_catalog = (
            build_formula_catalog(reference_path, reference_sheet)
            if reference_path is not None
            else None
        )
    # Compute and store: workbook.
    workbook = read_raw_workbook(source)
    # Check this condition before choosing the next branch.
    if sheet_name not in workbook:
        # Stop with a clear validation error for invalid input.
        raise ValueError(f"Worksheet not found: {sheet_name}")

    # Compute and store: source_columns.
    source_columns = detect_source_columns(workbook[sheet_name])
    # Compute and store: (calculated, output_values, errors).
    calculated, output_values, errors = calculate_raw_sheet(
        workbook[sheet_name],
        reference_catalog=reference_catalog,
    )
    # Run Audit step 7: insert calculated columns and formula comparison data.
    mismatch_count = sum(
        not match
        for values in output_values.values()
        for match in values["matches"].values()
    )

    # Read and edit the actual workbook instead of recreating it from a DataFrame.
    # This preserves sheet names, merged cells, row heights, fonts and cell styles.
    with tempfile.TemporaryDirectory(prefix="excel_auditor_") as temporary_name:
        # Compute and store: xlsx_source.
        xlsx_source = _xlsx_source(source, Path(temporary_name))
        # Compute and store: workbook_out.
        workbook_out = load_workbook(xlsx_source)
        # Compute and store: formula_workbook.
        formula_workbook = load_workbook(xlsx_source, data_only=False)
        # Compute and store: value_workbook.
        value_workbook = load_workbook(xlsx_source, data_only=True)
        # Compute and store: worksheet.
        worksheet = workbook_out[sheet_name]

        # Insert from left to right. These are Excel's 1-based column numbers.
        outer_quantity_name = _audit_quantity_column(source_columns, "S")
        # Compute and store: fireproof_quantity_name.
        fireproof_quantity_name = _audit_quantity_column(source_columns, "S_fireproof")
        # Compute and store: inner_quantity_name.
        inner_quantity_name = _audit_quantity_column(source_columns, "S_inner")
        # Compute and store: insertion_specs.
        insertion_specs = [
            (source_columns["K"], _K_OUTPUT, "K", 18),
            (source_columns[outer_quantity_name], _OUTER_S_OUTPUT, "S", 28),
        ]
        # Only add a separate SCC result when the source workbook really has
        # a separate fireproof quantity column.  In the legacy two-quantity
        # layout SCC is stored together with the inside/acoustic quantity and
        # must not be duplicated.
        if (
            fireproof_quantity_name in {"scc_quantity", "fireproof_quantity"}
            and source_columns.get(fireproof_quantity_name) not in {
                source_columns.get(outer_quantity_name),
                source_columns.get(inner_quantity_name),
            }
        ):
            # Execute this expression for its result or side effect.
            insertion_specs.append(
                (source_columns[fireproof_quantity_name], _FIREPROOF_S_OUTPUT, "S_fireproof", 28)
            )
        # Execute this expression for its result or side effect.
        insertion_specs.append(
            (source_columns[inner_quantity_name], _INNER_S_OUTPUT, "S_inner", 34)
        )
        # Aliases may point to the same source column in unusual templates.
        # Insert one result beside that column, not two overlapping columns.
        seen_columns: set[int] = set()
        # Compute and store: unique_insertions.
        unique_insertions = []
        # Process each item in this input sequence.
        for spec in sorted(insertion_specs):
            # Check this condition before choosing the next branch.
            if spec[0] in seen_columns:
                # Execute this statement as part of the current workflow.
                continue
            # Execute this expression for its result or side effect.
            seen_columns.add(spec[0])
            # Execute this expression for its result or side effect.
            unique_insertions.append(spec)
        # Compute and store: original_insertions.
        original_insertions = tuple(unique_insertions)
        # Compute and store: insertions.
        insertions = []
        # Compute and store: inserted_before.
        inserted_before = 0
        # Process each item in this input sequence.
        for original_column, label, value_key, width in sorted(original_insertions):
            # Compute and store: current_source.
            current_source = original_column + 1 + inserted_before
            # Execute this expression for its result or side effect.
            insertions.append((current_source + 1, current_source, label, value_key, width))
            # Update the accumulated value: inserted_before.
            inserted_before += 1
        # Process each item in this input sequence.
        for insert_at, source_column, label, value_key, width in insertions:
            # Execute this expression for its result or side effect.
            _shift_merged_ranges(worksheet, insert_at)
            # Execute this expression for its result or side effect.
            worksheet.insert_cols(insert_at, 1)
            # Execute this expression for its result or side effect.
            _copy_column_style(worksheet, source_column, insert_at, width=width)
            # Process each item in this input sequence.
            for raw_index, values in output_values.items():
                # Compute and store: source_row.
                source_row = int(raw_index) + 1
                # Compute and store: formula_references.
                formula_references = _final_formula_references(
                    source_columns,
                    source_row,
                    original_insertions,
                )
                # Compute and store: output_formula.
                output_formula = _render_output_formula(
                    values,
                    value_key,
                    formula_references,
                )
                # Compute and store: cell.
                cell = worksheet.cell(
                    row=source_row,
                    column=insert_at,
                )
                # Compute and store: cell.value.
                cell.value = output_formula or values[value_key]
                # Compute and store: font.
                font = copy(cell.font)
                # Compute and store: is_match.
                is_match = values["matches"].get(value_key, True)
                # Compute and store: font.color.
                font.color = "FF008000" if is_match else "FFC62828"
                # Compute and store: font.bold.
                font.bold = not is_match
                # Compute and store: cell.font.
                cell.font = font

        # Write headers after all insertions so merged two-row templates do
        # not move or hide labels written before a later insertion.
        header_row = _header_row(workbook[sheet_name]) + 1
        # Process each item in this input sequence.
        for insert_at, _source_column, label, _value_key, _width in insertions:
            # Execute this expression for its result or side effect.
            _set_new_header(worksheet, insert_at, header_row, _DISPLAY_HEADERS[label])

        # Compute and store: descriptions.
        descriptions = {
            int(item["source_row"]) - 1: str(item.get("description", ""))
            for item in calculated
        }
        # Execute this expression for its result or side effect.
        _write_original_formula_table(
            worksheet,
            formula_workbook[sheet_name],
            value_workbook[sheet_name],
            workbook[sheet_name],
            source_columns,
            output_values,
            descriptions,
        )

        # Execute this expression for its result or side effect.
        destination.parent.mkdir(parents=True, exist_ok=True)
        # Compute and store: workbook_out.calculation.fullCalcOnLoad.
        workbook_out.calculation.fullCalcOnLoad = True
        # Compute and store: workbook_out.calculation.forceFullCalc.
        workbook_out.calculation.forceFullCalc = True
        # Compute and store: workbook_out.calculation.calcOnSave.
        workbook_out.calculation.calcOnSave = True
        # Execute this expression for its result or side effect.
        workbook_out.save(destination)

    # Compute and store: preview_path.
    preview_path = _write_preview(
        destination,
        workbook[sheet_name],
        calculated,
        output_values,
        formula_workbook[sheet_name],
        value_workbook[sheet_name],
        source_columns,
    )
    # Run Audit step 8: return output paths and summary counts to the desktop app.

    # Return the value produced by this step.
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
