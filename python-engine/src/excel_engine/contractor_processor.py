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


# Purpose: Read the contractor layout and return normalized rows.
# Example output: DataFrame(normalized_rows=15)
def read_contractor_sheet(path: str | Path, sheet_name: str = "02") -> pd.DataFrame:
    """Read the contractor layout and return normalized rows."""
    # Compute and store: raw.
    raw = ExcelReader(path).read_sheet(sheet_name, skiprows=16, header=None)
    # Compute and store: converter.
    converter = ExcelConverter(raw)
    # Execute this expression for its result or side effect.
    converter.remove_columns([0, 3])
    # Execute this expression for its result or side effect.
    converter.rename_columns(_COLUMN_NAMES)
    # Execute this expression for its result or side effect.
    converter.add_k_by_function()
    # Compute and store: data.
    data = converter.get_data()
    # Compute and store: data['section'].
    data["section"] = data.apply(calculate_section, axis=1)
    # Compute and store: geometry_columns.
    geometry_columns = ["W", "H", "W1", "H1", "W2", "H2", "L_M", "R_D"]
    # Return the value produced by this step.
    return data.dropna(subset=geometry_columns, how="all").copy()


# Purpose: Run the legacy row-by-row contractor calculation workflow.
# Example output: {"calculated_count": 15, "calculated": [...], "errors": []}
def process_workbook(
    path: str | Path,
    sheet_name: str | None = None,
    section: str | None = None,
    product_type: str | None = None,
    header_map: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Run the legacy row-by-row contractor calculation workflow."""
    # Execute this statement as part of the current workflow.
    del header_map  # The contractor layout is normalized by read_contractor_sheet.
    # Compute and store: selected_sheet.
    selected_sheet = sheet_name or "02"
    # Compute and store: data.
    data = read_contractor_sheet(path, selected_sheet)
    # Compute and store: calculated.
    calculated: list[dict[str, Any]] = []
    # Compute and store: errors.
    errors: list[dict[str, Any]] = []

    # Process each item in this input sequence.
    for index, row in data.iterrows():
        # Compute and store: row_data.
        row_data = row.to_dict()
        # Compute and store: row_data['source_row'].
        row_data["source_row"] = int(index) + 17
        # Protect parsing/evaluation so bad workbook data becomes a controlled error.
        try:
            # Compute and store: product.
            product = ProductFormat.from_row(
                row_data,
                section=section,
                product_type=product_type,
            )
            # Compute and store: result.
            result = product.calculate(row_data)
            # Execute this expression for its result or side effect.
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
            # Execute this expression for its result or side effect.
            errors.append(
                {
                    "source_row": row_data["source_row"],
                    "description": row_data.get("description", ""),
                    "message": str(error),
                }
            )

    # Return the value produced by this step.
    return {
        "path": str(Path(path).expanduser().resolve()),
        "sheet": selected_sheet,
        "row_count": len(data),
        "calculated_count": len(calculated),
        "calculated": calculated,
        "errors": errors,
    }


__all__ = ["process_workbook", "read_contractor_sheet"]
