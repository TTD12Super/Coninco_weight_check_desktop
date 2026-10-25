import unittest

from excel_engine.formula_parser import (
    FormulaEvaluationError,
    FormulaParseError,
    canonical,
    extract_assignments,
    formula_block_result,
    formula_result,
    parse_formula,
    render_excel,
)


class FormulaParserTest(unittest.TestCase):
    def test_extracts_single_label_without_s_prefix_in_expression(self) -> None:
        self.assertEqual(
            [("SCN", "(W+H)*Q")],
            extract_assignments("Scn = (W+H)*Q"),
        )

    def test_extracts_multiple_labels_from_one_detail_cell(self) -> None:
        self.assertEqual(
            [("STA", "W*H"), ("SCC", "(W+2*t)*(H+2*t)")],
            extract_assignments("Sta=W*H\nScc=(W+2*t)*(H+2*t)"),
        )

    def test_maps_variables_to_actual_excel_cells(self) -> None:
        result = formula_result(
            "W+H",
            {"W": 800, "H": 500},
            references={"W": "E6", "H": "F6"},
        )
        self.assertEqual(1300, result["value"])
        self.assertEqual("=E6+F6", result["excel_formula"])

    def test_renders_formula_with_only_required_parentheses(self) -> None:
        node = parse_formula(
            "K*2*(((W+R)^2-R^2)+(2*(W+R)*H)+2*R*H)/1000000"
        )
        self.assertEqual(
            "Q37*2*((F37+M37)^2-M37^2+2*(F37+M37)*G37+2*M37*G37)/1000000",
            render_excel(
                node,
                {"K": "Q37", "W": "F37", "R": "M37", "H": "G37"},
            ),
        )

    def test_render_preserves_non_associative_grouping(self) -> None:
        references = {"A": "A1", "B": "B1", "C": "C1"}
        self.assertEqual(
            "A1-(B1-C1)", render_excel(parse_formula("A-(B-C)"), references)
        )
        self.assertEqual(
            "A1/(B1*C1)", render_excel(parse_formula("A/(B*C)"), references)
        )

    def test_reads_an_explicit_cell_reference(self) -> None:
        result = formula_result("W+E6+Q", {"W": 2, "Q": 3}, {"E6": 5})
        self.assertEqual(10, result["value"])

    def test_supports_implicit_multiplication(self) -> None:
        result = formula_result("4t+2PI()*Q", {"t": 5, "Q": 2})
        self.assertAlmostEqual(12.566370614359172 + 20, result["value"])

    def test_if_and_true_branch(self) -> None:
        expression = "IF(AND(t=0;W<=550);(W+H)*Q;(W+H+100)*Q)"
        result = formula_result(expression, {"W": 500, "H": 300, "t": 0, "Q": 2})
        self.assertEqual(1600, result["value"])

    def test_if_and_false_branch(self) -> None:
        expression = "IF(AND(t=0;W<=550);(W+H)*Q;(W+H+100)*Q)"
        result = formula_result(expression, {"W": 600, "H": 300, "t": 0, "Q": 2})
        self.assertEqual(2000, result["value"])

    def test_decimal_comma_in_detail_formula(self) -> None:
        result = formula_result("W-1,5*H", {"W": 10, "H": 2})
        self.assertEqual(7, result["value"])

    def test_if_comma_arguments_are_not_collapsed_into_decimal(self) -> None:
        result = formula_result("IF(L=0,0,W+H)", {"L_M": 1, "W": 3, "H": 4})
        self.assertEqual(7, result["value"])

    def test_dimension_names_that_look_like_cells(self) -> None:
        result = formula_result("W1+H1", {"W1": 2, "H1": 3})
        self.assertEqual(5, result["value"])

    def test_excel_unary_and_power_precedence(self) -> None:
        self.assertEqual(-4, formula_result("-W^2", {"W": 2})["value"])
        self.assertEqual(0.25, formula_result("2^-2", {})["value"])
        self.assertEqual(512, formula_result("2^3^2", {})["value"])

    def test_pi_and_sqrt(self) -> None:
        result = formula_result("PI()*SQRT(W^2+H^2)", {"W": 3, "H": 4})
        self.assertAlmostEqual(5 * 3.141592653589793, result["value"])

    def test_repairs_extra_closing_parentheses(self) -> None:
        result = formula_result(
            "(((W+2a)*(H+2a))*K)/1000000)*a/25",
            {"W": 1000, "H": 1000, "a": 25, "K": 1},
        )
        self.assertAlmostEqual(1.1025, result["value"])

    def test_repairs_multiple_extra_closing_parentheses(self) -> None:
        result = formula_result(
            "((W+2t)*(H+2t)))*K)/1000000",
            {"W": 1000, "H": 1000, "t": 25, "K": 1},
        )
        self.assertAlmostEqual(1.1025, result["value"])

    def test_repairs_missing_opening_without_changing_division_precedence(self) -> None:
        result = formula_result(
            "(W+H)*2*K)/1000000",
            {"W": 1000, "H": 500, "K": 1},
        )
        self.assertAlmostEqual(0.003, result["value"])

    def test_commutative_blocks_have_same_canonical_form(self) -> None:
        first = canonical(parse_formula("(A+B)+(B*2)"))
        second = canonical(parse_formula("(B*2)+(A+B)"))
        self.assertEqual(first, second)

    def test_unknown_variable_and_bad_parentheses_are_visible(self) -> None:
        with self.assertRaises(FormulaEvaluationError):
            formula_result("W+X", {"W": 1})
        with self.assertRaises(FormulaParseError):
            parse_formula("(W+H")

    def test_formula_block_resolves_intermediate_variables_in_any_order(self) -> None:
        assignments = extract_assignments(
            "S=K*(S_V1+S_V2)\nS_V2=2*H\nS_V1=2*W"
        )
        result = formula_block_result(
            assignments,
            {"W": 3, "H": 4, "K": 2},
            references={"W": "E6", "H": "F6", "K": "Q6"},
            target="S",
        )
        self.assertEqual(28, result["value"])
        self.assertIn("E6", result["excel_formula"])
        self.assertIn("F6", result["excel_formula"])

    def test_formula_block_supports_pi_and_prime_dimension_aliases(self) -> None:
        assignments = extract_assignments(
            "S=K*(S_V1-S_V4)\n"
            "S_V1=2*PI()*W\n"
            "S_V4=H'*W''/2"
        )
        result = formula_block_result(
            assignments,
            {"W": 10, "H1": 4, "W2": 6, "K": 1},
            target="S",
        )
        self.assertAlmostEqual(2 * 3.141592653589793 * 10 - 12, result["value"])

    def test_formula_block_auto_selects_one_output_target(self) -> None:
        assignments = extract_assignments(
            "S_V1=2*W\nS_V2=2*H\nSCC=S_V1+S_V2"
        )
        result = formula_block_result(assignments, {"W": 3, "H": 4})
        self.assertEqual(14, result["value"])

    def test_formula_block_rejects_intermediate_target(self) -> None:
        assignments = extract_assignments("S_V1=2*W\nS=S_V1")
        with self.assertRaises(FormulaEvaluationError):
            formula_block_result(assignments, {"W": 3}, target="S_V1")

    def test_formula_block_requires_explicit_target_for_multiple_outputs(self) -> None:
        assignments = extract_assignments("S=2*W\nSCC=3*W")
        with self.assertRaises(FormulaEvaluationError):
            formula_block_result(assignments, {"W": 3})


if __name__ == "__main__":
    unittest.main()
