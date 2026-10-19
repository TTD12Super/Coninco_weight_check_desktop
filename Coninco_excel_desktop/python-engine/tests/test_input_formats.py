import unittest

from excel_engine.input_formats import (
    PRODUCT_INPUT_FORMATS,
    SHEET_02_COLUMN_MAP,
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


if __name__ == "__main__":
    unittest.main()
