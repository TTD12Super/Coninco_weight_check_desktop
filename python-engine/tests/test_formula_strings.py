import unittest

from excel_engine import formula_engine as engine
from excel_engine.formula_strings import _FORMULAS, excel_input_references, formula_strings


class FormulaStringsTest(unittest.TestCase):
    def test_every_registered_calculator_has_three_excel_formulas(self) -> None:
        for section, calculators in engine._CALCULATORS.items():
            for product_type in calculators:
                formulas = formula_strings(section, product_type)
                self.assertEqual(3, len(formulas))
                for formula in formulas:
                    self.assertIsInstance(formula, str)
                    self.assertTrue(formula.startswith("="), formula)
                    self.assertEqual(formula.count("("), formula.count(")"), formula)
                    self.assertNotIn("math.", formula)
                    self.assertNotIn("**", formula)
                    self.assertNotRegex(formula, r"\b(?:x|sqrt)\.")

    def test_formula_functions_return_strings(self) -> None:
        for name in _FORMULAS:
            primary = getattr(__import__("excel_engine.formula_strings", fromlist=[name]), f"{name}_formula")
            secondary = getattr(__import__("excel_engine.formula_strings", fromlist=[name]), f"{name}_s_formula")
            tertiary = getattr(__import__("excel_engine.formula_strings", fromlist=[name]), f"{name}_t_formula")
            self.assertTrue(primary().startswith("="))
            self.assertTrue(secondary().startswith("="))
            self.assertTrue(tertiary().startswith("="))

    def test_excel_functions_and_cell_references_are_supported(self) -> None:
        formulas = formula_strings("standard", "square_round_reducer")
        self.assertIn("PI()", formulas[0])
        self.assertIn("SQRT(", formulas[0])

        references = excel_input_references(21)
        formula = formula_strings("standard", "rectangular_duct", references)[0]
        self.assertIn("F21", formula)
        self.assertIn("G21", formula)
        self.assertIn("P21", formula)
        self.assertIn("Q21", formula)
        self.assertNotRegex(formula, r"(?<![A-Za-z0-9_])(?:W|H|Q|K)(?![A-Za-z0-9_])")


if __name__ == "__main__":
    unittest.main()
