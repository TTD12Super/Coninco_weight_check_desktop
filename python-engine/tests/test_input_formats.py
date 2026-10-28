"""Unit tests for workbook header detection and quantity-column mapping."""

import unittest

import pandas as pd

from excel_engine.input_formats import (
    PRODUCT_INPUT_FORMATS,
    SHEET_02_COLUMN_MAP,
    detect_source_columns,
    quantity_column_for_target,
)


class InputFormatsTest(unittest.TestCase):
    """Verify flexible contractor headers map to stable internal names."""
    # Confirm the known Sheet 02 headers map to internal field names.
    def test_sheet_02_maps_formula_columns(self) -> None:
        self.assertEqual("W", SHEET_02_COLUMN_MAP["Unnamed: 5"])
        self.assertEqual("K", SHEET_02_COLUMN_MAP["Unnamed: 16"])

    # Confirm every supported Sheet 02 product has an input format.
    def test_product_formats_cover_sheet_02_products(self) -> None:
        for product_type in (
            "rectangular_duct",
            "round_duct",
            "takeoff",
            "elbow",
            "reducer",
            "mfd_valve",
            "nrd_valve",
        ):
            self.assertIn(product_type, PRODUCT_INPUT_FORMATS)

    # Confirm header detection survives a missing duplicate Unit column.
    def test_header_detection_handles_missing_duplicate_unit_column(self) -> None:
        columns = [None] * 24
        columns[1] = "No."
        columns[2] = "Name of Material/ Equipment"
        columns[3] = "Unit"
        columns[4:16] = ["W", "H", "W'", "H'", "W''", "H''", "L M", "R D", "a", "t", "Q", "Total of Quantities (K)"]
        columns[16:19] = ["Plan", "Up-Down", "Vertical axis"]
        columns[20] = "S Khối lượng cách nhiệt"
        columns[21] = "Formula"
        columns[22] = "Diễn giải chi tiết khối lượng cách nhiệt trong / chống cháy"
        columns[23] = "Note"
        detected = detect_source_columns(pd.DataFrame([columns]))
        self.assertEqual(4, detected["W"])
        self.assertEqual(5, detected["H"])
        self.assertEqual(15, detected["K"])

    # Confirm separate SCN, SCC, and STA columns map independently.
    def test_quantity_mapping_supports_three_named_columns(self) -> None:
        columns = [None] * 24
        columns[1] = "No."
        columns[2] = "Name of Material/ Equipment"
        columns[3] = "Unit"
        columns[4:16] = ["W", "H", "W'", "H'", "W''", "H''", "L M", "R D", "a", "t", "Q", "K"]
        columns[16:19] = ["Plan", "Up-Down", "Vertical axis"]
        columns[19:22] = ["SCN (m2)", "SCC (m2)", "STA (m2)"]
        columns[22] = "Công thức"
        columns[23] = "Diễn giải"

        detected = detect_source_columns(pd.DataFrame([columns]))

        self.assertEqual(19, detected["scn_quantity"])
        self.assertEqual(20, detected["scc_quantity"])
        self.assertEqual(21, detected["sta_quantity"])
        self.assertEqual("scn_quantity", quantity_column_for_target(detected, "S"))
        self.assertEqual("scc_quantity", quantity_column_for_target(detected, "SCC"))
        self.assertEqual("sta_quantity", quantity_column_for_target(detected, "STA"))

    # Confirm legacy two-column layouts use the shared SCC/TA fallback.
    def test_quantity_mapping_keeps_two_column_scc_fallback(self) -> None:
        columns = [None] * 24
        columns[20] = "Khối lượng cách nhiệt (m2)"
        columns[21] = "Khối lượng cách nhiệt trong (tiêu âm) (m2)"

        detected = detect_source_columns(pd.DataFrame([columns]))

        self.assertNotIn("scc_quantity", detected)
        self.assertEqual("acoustic_quantity", quantity_column_for_target(detected, "SCC"))
        self.assertEqual("sta_quantity", quantity_column_for_target(detected, "SCC/TA"))

if __name__ == "__main__":
    unittest.main()
