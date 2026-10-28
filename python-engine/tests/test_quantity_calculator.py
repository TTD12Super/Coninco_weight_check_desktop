"""Unit tests for Plan, Up/Down, and Vertical Axis quantity parsing."""

import unittest

from excel_engine.quantity_calculator import calculate_K, evaluate_expression


class QuantityCalculatorTest(unittest.TestCase):
    """Verify safe K calculation from contractor quantity components."""
    # Confirm quantity expressions accept Vietnamese decimal commas.
    def test_evaluates_decimal_comma_expression(self) -> None:
        self.assertAlmostEqual(0.92, evaluate_expression("0,120 + (2*0,400)"))

    # Confirm K sums Plan, Up/Down, and Vertical Axis when present.
    def test_sums_the_three_k_columns(self) -> None:
        row = {
            "plan": "0,120 + (2*0,400)",
            "up_down": "1,065",
            "vertical_axis": None,
        }
        self.assertAlmostEqual(1.985, calculate_K(row))


if __name__ == "__main__":
    unittest.main()
