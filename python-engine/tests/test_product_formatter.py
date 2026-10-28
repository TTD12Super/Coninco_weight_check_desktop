"""Unit tests for product classification and row-to-input conversion."""

import unittest

from excel_engine.product_formatter import ProductFormat, calculate_section


class ProductFormatTest(unittest.TestCase):
    """Verify descriptions and sections select the expected calculator."""
    # Confirm a normal rectangular duct selects the standard calculator.
    def test_formats_standard_rectangular_duct(self) -> None:
        row = {
            "description": "Fresh air duct 800x400",
            "W": 800,
            "H": 400,
            "Q": 9,
            "K": 5.325,
        }

        product = ProductFormat.from_row(row)
        result = product.calculate(row)

        self.assertEqual("standard", product.section)
        self.assertEqual("rectangular_duct", product.product_type)
        self.assertAlmostEqual(15.12, result.R)

    # Confirm an explicit fireproof section overrides default classification.
    def test_section_column_selects_fireproof_section(self) -> None:
        row = {
            "description": "Reducer 360x360-350x350",
            "section": "fireproof",
            "W": 360,
            "H": 360,
            "W1": 350,
            "H1": 350,
            "L_M": 175,
            "Q": 1,
            "K": 1,
        }

        product = ProductFormat.from_row(row)

        self.assertEqual("fireproof", product.section)
        self.assertEqual("reducer", product.product_type)

    # Confirm an STA label identifies the acoustic family.
    def test_calculate_section_from_sta_formula(self) -> None:
        row = {
            "description": "Reducer 360x360-350x350",
            "formula_or_insulation_detail": "Sta=...",
        }

        self.assertEqual("acoustic", calculate_section(row))

    # Confirm a shared SCC/TA label identifies the acoustic family.
    def test_calculate_section_from_scc_ta_formula(self) -> None:
        row = {
            "description": "Reducer 360x360-350x350",
            "acoustic_fireproof_detail": "Scc/ta=...",
        }

        self.assertEqual("acoustic", calculate_section(row))


if __name__ == "__main__":
    unittest.main()
