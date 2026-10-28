"""Minimal pandas transformations for the legacy contractor layout."""

import pandas as pd

from .quantity_calculator import calculate_K
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

class ExcelConverter:
    """Apply small in-memory transformations to contractor DataFrames."""

    # Purpose: Store the DataFrame that will be transformed in place.
    def __init__(self, dataFrame: pd.DataFrame):
        """Store the DataFrame that will be transformed in place."""
        # Compute and store: self.dataFrame.
        self.dataFrame = dataFrame.copy()

    # Purpose: Remove the requested columns when they exist.
    def remove_columns(self, columns: list[str]) -> None:
        """Remove the requested columns when they exist."""
        # Compute and store: self.dataFrame.
        self.dataFrame = self.dataFrame.drop(
            columns=columns,
            errors="ignore"
        )

    # Purpose: Drop rows empty across the requested columns.
    def remove_empty_rows(self, columns: list[str]) -> None:
        """Drop rows empty across the requested columns."""
        # Compute and store: self.dataFrame.
        self.dataFrame = self.dataFrame.dropna(
            subset=columns,
            how="all"
        )

    # Purpose: Drop rows containing null values in required columns.
    def remove_rows_with_null(self, columns: list[str]) -> None:
        """Drop rows containing null values in required columns."""
        # Compute and store: self.dataFrame.
        self.dataFrame = self.dataFrame.dropna(
            subset=columns,
            how="any"
        )

    # Purpose: Remove duplicate rows while keeping the first occurrence.
    def remove_duplicate_rows(self) -> None:
        """Remove duplicate rows while keeping the first occurrence."""
        # Compute and store: self.dataFrame.
        self.dataFrame = self.dataFrame.drop_duplicates()

    # Purpose: Rename DataFrame columns using the supplied mapping.
    def rename_columns(self, mapping: dict[str, str]) -> None:
        """Rename DataFrame columns using the supplied mapping."""
        # Compute and store: self.dataFrame.
        self.dataFrame = self.dataFrame.rename(columns=mapping)

    # Purpose: Set or replace one DataFrame column.
    def set_column(self, column: str, values) -> None:
        """Set or replace one DataFrame column."""
        # Compute and store: self.dataFrame[column].
        self.dataFrame[column] = values

    # Purpose: Add K calculated from plan, up/down, and vertical-axis cells.
    def add_k_by_function(self) -> None:
        """Add K calculated from plan, up/down, and vertical-axis cells."""
        # Compute and store: self.dataFrame['K_By_function'].
        self.dataFrame["K_By_function"] = self.dataFrame.apply(
            calculate_K,
            axis=1,
        )

    # Purpose: Return the transformed DataFrame.
    # Example output: DataFrame(rows=15, columns=21)
    def get_data(self) -> pd.DataFrame:
        """Return the transformed DataFrame."""
        # Return the value produced by this step.
        return self.dataFrame

    # Purpose: Return the transformed DataFrame.
    # Example output: DataFrame(rows=15, columns=21)
    def get_data(self) -> pd.DataFrame:
        """Return the transformed DataFrame."""
        # Return the value produced by this step.
        return self.dataFrame

    # Purpose: Return the first rows for debugging or UI inspection.
    # Example output: DataFrame(rows=10, columns=21)
    def preview(self, rows: int = 10) -> pd.DataFrame:
        """Return the first rows for debugging or UI inspection."""
        # Return the value produced by this step.
        return self.dataFrame.head(rows)
    
    # Purpose: Write the transformed DataFrame to an Excel file.
    def export_excel(self, output_path: str) -> None:
        """Write the transformed DataFrame to an Excel file."""
        # Execute this expression for its result or side effect.
        self.dataFrame.to_excel(
            output_path,
            index=False,
            engine="openpyxl"
        )

        # Execute this statement as part of the current workflow.
        from openpyxl import load_workbook

        # Compute and store: workbook.
        workbook = load_workbook(output_path)
        # Compute and store: worksheet.
        worksheet = workbook.active

        # Compute and store: thin_border.
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        # Process each item in this input sequence.
        for cell in worksheet[1]:
            # Compute and store: cell.font.
            cell.font = Font(bold=True)
            # Compute and store: cell.alignment.
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        # Process each item in this input sequence.
        for row in worksheet.iter_rows():
            # Process each item in this input sequence.
            for cell in row:
                # Compute and store: cell.border.
                cell.border = thin_border
                # Compute and store: cell.alignment.
                cell.alignment = Alignment(
                    vertical="center",
                    wrap_text=True
                )

        # Process each item in this input sequence.
        for column_cells in worksheet.columns:
            # Compute and store: max_length.
            max_length = 0
            # Compute and store: column_letter.
            column_letter = get_column_letter(column_cells[0].column)

            # Process each item in this input sequence.
            for cell in column_cells:
                # Check this condition before choosing the next branch.
                if cell.value is not None:
                    # Compute and store: max_length.
                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )

            # Compute and store: worksheet.column_dimensions[column_letter].width.
            worksheet.column_dimensions[column_letter].width = min(
                max_length + 2,
                40
            )

        # Compute and store: worksheet.freeze_panes.
        worksheet.freeze_panes = "A2"
        # Compute and store: worksheet.auto_filter.ref.
        worksheet.auto_filter.ref = worksheet.dimensions

        # Execute this expression for its result or side effect.
        workbook.save(output_path)
