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


def _text(value: Any) -> str:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    return str(value).replace("\r", " ").replace("\n", " ").strip()


def _read_workbook(path: Path) -> dict[str, pd.DataFrame]:
    reader = ExcelReader(path)
    return pd.read_excel(
        reader.path,
        sheet_name=None,
        engine=reader.engine,
        header=None,
        skiprows=0,
    )


def _sheet_names(path: Path) -> list[str]:
    """Read worksheet names without loading every cell in the workbook."""
    reader = ExcelReader(path)
    if reader.engine == "openpyxl":
        from openpyxl import load_workbook

        workbook = load_workbook(reader.path, read_only=True, data_only=False)
        try:
            return list(workbook.sheetnames)
        finally:
            workbook.close()

    with pd.ExcelFile(reader.path, engine=reader.engine) as workbook:
        return list(workbook.sheet_names)


def _workbook_path(path: str | Path, temporary_directory: Path) -> Path:
    source = Path(path).expanduser().resolve()
    if source.suffix.lower() != ".numbers":
        return source

    target = temporary_directory / "reference.xlsx"
    try:
        subprocess.run(
            ["osascript", "-", str(source), str(target)],
            input=_NUMBERS_EXPORT_SCRIPT,
            text=True,
            capture_output=True,
            check=True,
        )
    except FileNotFoundError as error:
        raise ValueError(
            "Không thể đọc file .numbers vì máy chưa có Apple Numbers/osascript. "
            "Hãy xuất file chuẩn thành .xlsx."
        ) from error
    except subprocess.CalledProcessError as error:
        detail = error.stderr.strip() or "Numbers không thể xuất file thành Excel."
        raise ValueError(detail) from error
    if not target.is_file():
        raise ValueError("Không tạo được file Excel tạm từ file .numbers.")
    return target


def _formula_signature(block: str, target: str) -> Any | None:
    assignments = extract_assignments(block)
    parsed: dict[str, Any] = {}
    for label, expression in assignments:
        try:
            parsed[label.upper()] = parse_formula(expression)
        except FormulaParseError:
            return None
    node = parsed.get(target.upper())
    if node is None:
        return None
    try:
        return canonical(_expand_assignments(node, parsed))
    except Exception:  # noqa: BLE001 - keep other reference rows usable.
        return None


def _category(source_name: str, label: str) -> str:
    target = label.upper()
    if target in {"STA", "SCC/TA"}:
        return "Tiêu âm"
    if target in {"SCC", "SCN"} and source_name == "acoustic_fireproof_detail":
        return "Chống cháy"
    return "Cách nhiệt"


def _iter_reference_rows(data: pd.DataFrame, sheet_name: str) -> Iterator[dict[str, Any]]:
    columns = detect_source_columns(data)
    family = "standard"
    for index, raw_row in data.iterrows():
        heading = " ".join(_text(value).casefold() for value in raw_row.iloc[:3].tolist())
        if "ống gió chống cháy" in heading or "ong gio chong chay" in heading:
            family = "fireproof"
        elif "ống gió tiêu âm" in heading or "ong gio tieu am" in heading:
            family = "acoustic"
        elif "ống gió thường" in heading or "ong gio thuong" in heading:
            family = "standard"
        name = _text(raw_row.iloc[columns["description"]])
        if not name or name.casefold() in {"mô tả", "mo ta"}:
            continue
        try:
            product_type = _product_type(name)
        except (TypeError, ValueError):
            product_type = ""

        for source_name in ("formula_or_insulation_detail", "acoustic_fireproof_detail"):
            source_column = columns.get(source_name)
            if source_column is None or source_column >= len(raw_row):
                continue
            detail = _text(raw_row.iloc[source_column])
            if not detail:
                continue
            assignments = extract_assignments(detail)
            outputs = [
                (label, expression)
                for label, expression in assignments
                if label.upper() in OUTPUT_LABELS
            ]
            for label, expression in outputs:
                target = label.upper()
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


def build_formula_catalog(path: str | Path, sheet_name: str | None = None) -> dict[str, Any]:
    """Extract formula descriptions without calculating the reference workbook."""
    with tempfile.TemporaryDirectory(prefix="coninco_reference_") as temporary_name:
        workbook_path = _workbook_path(path, Path(temporary_name))
        workbook = _read_workbook(workbook_path)
        selected_sheet = sheet_name or next(iter(workbook), None)
        if selected_sheet is None or selected_sheet not in workbook:
            raise ValueError("Reference workbook has no readable worksheet.")
        entries = list(_iter_reference_rows(workbook[selected_sheet], selected_sheet))

    return {
        "source": str(Path(path).expanduser().resolve()),
        "sheet": selected_sheet,
        "entries": entries,
        "count": len(entries),
    }


def list_reference_sheets(path: str | Path) -> list[str]:
    """List worksheets in an Excel or Numbers reference workbook."""
    with tempfile.TemporaryDirectory(prefix="coninco_reference_sheets_") as temporary_name:
        workbook_path = _workbook_path(path, Path(temporary_name))
        return _sheet_names(workbook_path)


def match_reference_formula(
    catalog: dict[str, Any] | None,
    product_type: str,
    target: str,
    description: str = "",
    category: str | None = None,
    family: str | None = None,
) -> dict[str, Any] | None:
    """Find the authoritative formula for one product/quantity combination.

    The contractor description is often more detailed than the reference
    product name (for example ``Square elbow 360x360`` vs ``Co Vuong``), so
    product type is the primary key.  Category is the important second key:
    the same product has different formulas for insulation, fireproofing and
    acoustic work.
    """
    if not catalog:
        return None
    target_key = target.upper()
    candidates = [
        entry
        for entry in catalog.get("entries", [])
        if str(entry.get("target", "")).upper() == target_key
        and (not product_type or entry.get("product_type") == product_type)
        and (not category or str(entry.get("category", "")).casefold() == category.casefold())
        and (not family or str(entry.get("family", "")).casefold() == family.casefold())
    ]
    if not candidates:
        return None
    normalized_description = re.sub(r"\s+", " ", description).strip().casefold()
    exact = [
        entry for entry in candidates
        if re.sub(r"\s+", " ", str(entry.get("name", ""))).strip().casefold()
        == normalized_description
    ]
    if exact:
        return exact[0]

    # Prefer a formula with a valid parsed signature.  Malformed reference
    # rows remain visible in the catalog, but should not win over a usable
    # formula for the same product/category/target.
    usable = [entry for entry in candidates if entry.get("signature") is not None]
    return usable[0] if usable else candidates[0]


__all__ = ["OUTPUT_LABELS", "build_formula_catalog", "list_reference_sheets", "match_reference_formula"]
