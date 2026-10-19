import unittest

from excel_engine.quantity_calculator import calculate_K, evaluate_expression


class QuantityCalculatorTest(unittest.TestCase):
    def test_evaluates_decimal_comma_expression(self) -> None:
        self.assertAlmostEqual(0.92, evaluate_expression("0,120 + (2*0,400)"))

    def test_sums_the_three_k_columns(self) -> None:
        row = {
            "plan": "0,120 + (2*0,400)",
            "up_down": "1,065",
            "vertical_axis": None,
        }
        self.assertAlmostEqual(1.985, calculate_K(row))


if __name__ == "__main__":
    unittest.main()
