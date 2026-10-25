import unittest

import pandas as pd

from excel_engine.input_formats import (
    PRODUCT_INPUT_FORMATS,
    SHEET_02_COLUMN_MAP,
    detect_source_columns,
)


class InputFormatsTest(unittest.TestCase):
    def test_sheet_02_maps_formula_columns(self) -> None:
        self.assertEqual("W", SHEET_02_COLUMN_MAP["Unnamed: 5"])
        self.assertEqual("K", SHEET_02_COLUMN_MAP["Unnamed: 16"])

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

if __name__ == "__main__":
    unittest.main()
