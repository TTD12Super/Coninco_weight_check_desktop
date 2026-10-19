"""Prepare contractor sheet rows and run the selected formula."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

from .excel_converter import ExcelConverter
from .excel_reader import ExcelReader
from .product_formatter import ProductFormat, calculate_section


_COLUMN_NAMES = {
    1: "item_no",
    2: "description",
    4: "unit",
    5: "W",
    6: "H",
    7: "W1",
    8: "H1",
    9: "W2",
    10: "H2",
    11: "L_M",
    12: "R_D",
    13: "a",
    14: "t",
    15: "Q",
    16: "K",
    17: "plan",
    18: "up_down",
    19: "vertical_axis",
    20: "insulation_quantity",
    21: "acoustic_quantity",
    22: "formula_or_insulation_detail",
    23: "acoustic_fireproof_detail",
    24: "note",
}


def read_contractor_sheet(path: str | Path, sheet_name: str = "02") -> pd.DataFrame:
    """Read the contractor layout and return normalized rows."""
    raw = ExcelReader(path).read_sheet(sheet_name, skiprows=16, header=None)
    converter = ExcelConverter(raw)
    converter.remove_columns([0, 3])
    converter.rename_columns(_COLUMN_NAMES)
    converter.add_k_by_function()
    data = converter.get_data()
    data["section"] = data.apply(calculate_section, axis=1)
    geometry_columns = ["W", "H", "W1", "H1", "W2", "H2", "L_M", "R_D"]
    return data.dropna(subset=geometry_columns, how="all").copy()


def process_workbook(
    path: str | Path,
    sheet_name: str | None = None,
    section: str | None = None,
    product_type: str | None = None,
    header_map: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Format each row and calculate R/S/T where the mapping is supported."""
    del header_map  # The contractor layout is normalized by read_contractor_sheet.
    selected_sheet = sheet_name or "02"
    data = read_contractor_sheet(path, selected_sheet)
    calculated: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []

    for index, row in data.iterrows():
        row_data = row.to_dict()
        row_data["source_row"] = int(index) + 17
        try:
            product = ProductFormat.from_row(
                row_data,
                section=section,
                product_type=product_type,
            )
            result = product.calculate(row_data)
            calculated.append(
                {
                    "source_row": row_data["source_row"],
                    "description": product.description,
                    "section": product.section,
                    "product_type": product.product_type,
                    "inputs": product.inputs_from(row_data).__dict__,
                    "result": result.to_dict(),
                }
            )
        except (TypeError, ValueError) as error:
            errors.append(
                {
                    "source_row": row_data["source_row"],
                    "description": row_data.get("description", ""),
                    "message": str(error),
                }
            )

    return {
        "path": str(Path(path).expanduser().resolve()),
        "sheet": selected_sheet,
        "row_count": len(data),
        "calculated_count": len(calculated),
        "calculated": calculated,
        "errors": errors,
    }


__all__ = ["process_workbook", "read_contractor_sheet"]
