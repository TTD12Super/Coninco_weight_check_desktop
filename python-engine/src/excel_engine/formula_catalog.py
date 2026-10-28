"""Build and match a formula catalog from a contractor reference workbook."""

from __future__ import annotations

from pathlib import Path
import re
import subprocess
import tempfile
from typing import Any, Iterator

import pandas as pd

from .excel_reader import ExcelReader
from .formula_parser import (
    FormulaParseError,
    _expand_assignments,
    canonical,
    extract_assignments,
    parse_formula,
)
from .input_formats import detect_source_columns
from .product_formatter import _product_type


OUTPUT_LABELS = frozenset({"S", "SCN", "SCC", "STA", "SCC/TA"})

_NUMBERS_EXPORT_SCRIPT = (
    "on run argv\n"
    "    set sourcePath to item 1 of argv\n"
    "    set targetPath to item 2 of argv\n"
    "    tell application \"Numbers\"\n"
    "        set referenceDocument to open POSIX file sourcePath\n"
    "        export referenceDocument to POSIX file targetPath as Microsoft Excel\n"
    "        close referenceDocument saving no\n"
    "    end tell\n"
    "end run\n"
)


# Purpose: Convert one workbook cell to clean display text.
# Example output: "Square elbow 360x360"
def _text(value: Any) -> str:
    """Convert one workbook cell to clean display text."""
    # Check this condition before choosing the next branch.
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        # Return the value produced by this step.
        return ""
    # Return the value produced by this step.
    return str(value).replace("\r", " ").replace("\n", " ").strip()


# Purpose: Load every worksheet from an Excel or Numbers-converted workbook.
# Example output: {"02": DataFrame(rows=61, columns=25)}
def _read_workbook(path: Path) -> dict[str, pd.DataFrame]:
    """Load every worksheet from an Excel or Numbers-converted workbook."""
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


# Purpose: Read worksheet names without loading every cell in the workbook.
# Example output: ["02", "COMBINE"]
def _sheet_names(path: Path) -> list[str]:
    """Read worksheet names without loading every cell in the workbook."""
    # Compute and store: reader.
    reader = ExcelReader(path)
    # Check this condition before choosing the next branch.
    if reader.engine == "openpyxl":
        # Execute this statement as part of the current workflow.
        from openpyxl import load_workbook

        # Compute and store: workbook.
        workbook = load_workbook(reader.path, read_only=True, data_only=False)
        # Protect parsing/evaluation so bad workbook data becomes a controlled error.
        try:
            # Return the value produced by this step.
            return list(workbook.sheetnames)
        finally:
            # Execute this expression for its result or side effect.
            workbook.close()

    # Open the required resource for this operation.
    with pd.ExcelFile(reader.path, engine=reader.engine) as workbook:
        # Return the value produced by this step.
        return list(workbook.sheet_names)


# Purpose: Return an Excel path, exporting Apple Numbers files when necessary.
# Example output: Path("/tmp/reference.xlsx")
def _workbook_path(path: str | Path, temporary_directory: Path) -> Path:
    """Return an Excel path, exporting Apple Numbers files when necessary."""
    # Compute and store: source.
    source = Path(path).expanduser().resolve()
    # Check this condition before choosing the next branch.
    if source.suffix.lower() != ".numbers":
        # Return the value produced by this step.
        return source

    # Compute and store: target.
    target = temporary_directory / "reference.xlsx"
    # Protect parsing/evaluation so bad workbook data becomes a controlled error.
    try:
        # Execute this expression for its result or side effect.
        subprocess.run(
            ["osascript", "-", str(source), str(target)],
            input=_NUMBERS_EXPORT_SCRIPT,
            text=True,
            capture_output=True,
            check=True,
        )
    except FileNotFoundError as error:
        # Stop with a clear validation error for invalid input.
        raise ValueError(
            "Không thể đọc file .numbers vì máy chưa có Apple Numbers/osascript. "
            "Hãy xuất file chuẩn thành .xlsx."
        ) from error
    except subprocess.CalledProcessError as error:
        # Compute and store: detail.
        detail = error.stderr.strip() or "Numbers không thể xuất file thành Excel."
        # Stop with a clear validation error for invalid input.
        raise ValueError(detail) from error
    # Check this condition before choosing the next branch.
    if not target.is_file():
        # Stop with a clear validation error for invalid input.
        raise ValueError("Không tạo được file Excel tạm từ file .numbers.")
    # Return the value produced by this step.
    return target


# Purpose: Parse and canonicalize one target formula for catalog matching.
def _formula_signature(block: str, target: str) -> Any | None:
    """Parse and canonicalize one target formula for catalog matching."""
    # Compute and store: assignments.
    assignments = extract_assignments(block)
    # Compute and store: parsed.
    parsed: dict[str, Any] = {}
    # Process each item in this input sequence.
    for label, expression in assignments:
        # Protect parsing/evaluation so bad workbook data becomes a controlled error.
        try:
            # Compute and store: parsed[label.upper()].
            parsed[label.upper()] = parse_formula(expression)
        except FormulaParseError:
            # Return the value produced by this step.
            return None
    # Compute and store: node.
    node = parsed.get(target.upper())
    # Check this condition before choosing the next branch.
    if node is None:
        # Return the value produced by this step.
        return None
    # Protect parsing/evaluation so bad workbook data becomes a controlled error.
    try:
        # Return the value produced by this step.
        return canonical(_expand_assignments(node, parsed))
    except Exception:  # noqa: BLE001 - keep other reference rows usable.
        # Return the value produced by this step.
        return None


# Purpose: Map a formula source column and label to the work category.
# Example output: "Tiêu âm"
def _category(source_name: str, label: str) -> str:
    """Map a formula source column and label to the work category."""
    # Compute and store: target.
    target = label.upper()
    # Check this condition before choosing the next branch.
    if target in {"STA", "SCC/TA"}:
        # Return the value produced by this step.
        return "Tiêu âm"
    # Check this condition before choosing the next branch.
    if target in {"SCC", "SCN"} and source_name == "acoustic_fireproof_detail":
        # Return the value produced by this step.
        return "Chống cháy"
    # Return the value produced by this step.
    return "Cách nhiệt"


# Purpose: Yield one catalog entry for every output formula in every row.
# Example output: {"name": "Co vuông", "target": "S", "formula": "..."}
def _iter_reference_rows(data: pd.DataFrame, sheet_name: str) -> Iterator[dict[str, Any]]:
    """Yield one catalog entry for every output formula in every row."""
    # Compute and store: columns.
    columns = detect_source_columns(data)
    # Compute and store: family.
    family = "standard"
    # Process each item in this input sequence.
    for index, raw_row in data.iterrows():
        # Compute and store: heading.
        heading = " ".join(_text(value).casefold() for value in raw_row.iloc[:3].tolist())
        # Check this condition before choosing the next branch.
        if "ống gió chống cháy" in heading or "ong gio chong chay" in heading:
            # Compute and store: family.
            family = "fireproof"
        # Check this condition before choosing the next branch.
        elif "ống gió tiêu âm" in heading or "ong gio tieu am" in heading:
            # Compute and store: family.
            family = "acoustic"
        # Check this condition before choosing the next branch.
        elif "ống gió thường" in heading or "ong gio thuong" in heading:
            # Compute and store: family.
            family = "standard"
        # Compute and store: name.
        name = _text(raw_row.iloc[columns["description"]])
        # Check this condition before choosing the next branch.
        if not name or name.casefold() in {"mô tả", "mo ta"}:
            # Execute this statement as part of the current workflow.
            continue
        # Protect parsing/evaluation so bad workbook data becomes a controlled error.
        try:
            # Compute and store: product_type.
            product_type = _product_type(name)
        except (TypeError, ValueError):
            # Compute and store: product_type.
            product_type = ""

        # Process each item in this input sequence.
        for source_name in ("formula_or_insulation_detail", "acoustic_fireproof_detail"):
            # Compute and store: source_column.
            source_column = columns.get(source_name)
            # Check this condition before choosing the next branch.
            if source_column is None or source_column >= len(raw_row):
                # Execute this statement as part of the current workflow.
                continue
            # Compute and store: detail.
            detail = _text(raw_row.iloc[source_column])
            # Check this condition before choosing the next branch.
            if not detail:
                # Execute this statement as part of the current workflow.
                continue
            # Compute and store: assignments.
            assignments = extract_assignments(detail)
            # Compute and store: outputs.
            outputs = [
                (label, expression)
                for label, expression in assignments
                if label.upper() in OUTPUT_LABELS
            ]
            # Process each item in this input sequence.
            for label, expression in outputs:
                # Compute and store: target.
                target = label.upper()
                # Execute this expression for its result or side effect.
                yield {
                    "name": name,
                    "category": _category(source_name, target),
                    "family": family,
                    "target": target,
                    "product_type": product_type,
                    "formula": expression,
                    "formula_block": "\n".join(
                        f"{assignment_label} = {assignment_expression}"
                        for assignment_label, assignment_expression in assignments
                    ),
                    "signature": _formula_signature(detail, target),
                    "source_sheet": sheet_name,
                    "source_row": int(index) + 1,
                }


# Purpose: Extract formula descriptions without calculating the reference workbook.
# Example output: {"sheet": "COMBINE", "entries": [...], "count": 12}
def build_formula_catalog(path: str | Path, sheet_name: str | None = None) -> dict[str, Any]:
    """Extract formula descriptions without calculating the reference workbook."""
    # Open the required resource for this operation.
    with tempfile.TemporaryDirectory(prefix="coninco_reference_") as temporary_name:
        # Compute and store: workbook_path.
        workbook_path = _workbook_path(path, Path(temporary_name))
        # Compute and store: workbook.
        workbook = _read_workbook(workbook_path)
        # Compute and store: selected_sheet.
        selected_sheet = sheet_name or next(iter(workbook), None)
        # Check this condition before choosing the next branch.
        if selected_sheet is None or selected_sheet not in workbook:
            # Stop with a clear validation error for invalid input.
            raise ValueError("Reference workbook has no readable worksheet.")
        # Compute and store: entries.
        entries = list(_iter_reference_rows(workbook[selected_sheet], selected_sheet))

    # Return the value produced by this step.
    return {
        "source": str(Path(path).expanduser().resolve()),
        "sheet": selected_sheet,
        "entries": entries,
        "count": len(entries),
    }


# Purpose: List worksheets in an Excel or Numbers reference workbook.
# Example output: ["02", "COMBINE"]
def list_reference_sheets(path: str | Path) -> list[str]:
    """List worksheets in an Excel or Numbers reference workbook."""
    # Open the required resource for this operation.
    with tempfile.TemporaryDirectory(prefix="coninco_reference_sheets_") as temporary_name:
        # Compute and store: workbook_path.
        workbook_path = _workbook_path(path, Path(temporary_name))
        # Return the value produced by this step.
        return _sheet_names(workbook_path)


# Purpose: Find the authoritative formula for one product/quantity combination.
# Example output: {"product_type": "square_elbow", "target": "S", "formula": "..."}
def match_reference_formula(
    catalog: dict[str, Any] | None,
    product_type: str,
    target: str,
    description: str = "",
    category: str | None = None,
    family: str | None = None,
) -> dict[str, Any] | None:
    """Find the authoritative formula for one product/quantity combination.

    Product type, category, family, and target select the standard formula.

    The contractor description is often more detailed than the reference
    product name (for example ``Square elbow 360x360`` vs ``Co Vuong``), so
    product type is the primary key.  Category is the important second key:
    the same product has different formulas for insulation, fireproofing and
    acoustic work.
    """
    # Check this condition before choosing the next branch.
    if not catalog:
        # Return the value produced by this step.
        return None
    # Compute and store: target_key.
    target_key = target.upper()
    # Compute and store: candidates.
    candidates = [
        entry
        for entry in catalog.get("entries", [])
        if str(entry.get("target", "")).upper() == target_key
        and (not product_type or entry.get("product_type") == product_type)
        and (not category or str(entry.get("category", "")).casefold() == category.casefold())
        and (not family or str(entry.get("family", "")).casefold() == family.casefold())
    ]
    # Check this condition before choosing the next branch.
    if not candidates:
        # Return the value produced by this step.
        return None
    # Compute and store: normalized_description.
    normalized_description = re.sub(r"\s+", " ", description).strip().casefold()
    # Compute and store: exact.
    exact = [
        entry for entry in candidates
        if re.sub(r"\s+", " ", str(entry.get("name", ""))).strip().casefold()
        == normalized_description
    ]
    # Check this condition before choosing the next branch.
    if exact:
        # Return the value produced by this step.
        return exact[0]

    # Prefer a formula with a valid parsed signature.  Malformed reference
    # rows remain visible in the catalog, but should not win over a usable
    # formula for the same product/category/target.
    usable = [entry for entry in candidates if entry.get("signature") is not None]
    # Return the value produced by this step.
    return usable[0] if usable else candidates[0]


__all__ = ["OUTPUT_LABELS", "build_formula_catalog", "list_reference_sheets", "match_reference_formula"]
