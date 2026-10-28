"""Integration tests for row calculation, export, preview, and AST reporting."""

import csv
import json
import os
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from openpyxl import Workbook, load_workbook

from excel_engine.workbook_report import (
    _ast_difference_frontier,
    _formula_comparison,
    _repair_known_detail_expression,
    calculate_raw_sheet,
    export_calculated_workbook,
    read_raw_workbook,
)
from excel_engine.formula_parser import parse_formula


_EXTERNAL_REFERENCE_WORKBOOK = Path(
    "/Users/trantrungdung/Downloads/4. Form tính tiêu âm, cách nhiệt, chống cháy combine (2) (1).xlsx"
)


def _value_only_workbook(test_case: unittest.TestCase) -> Path:
    """Create the fixture workbook containing cached values only."""
    candidates = (
        Path(__file__).parents[1] / "src" / "excel_engine" / "L01GF0087_value_only.xls",
        Path("/Users/trantrungdung/Documents/L01GF0087_value_only (1).xls"),
    )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    test_case.skipTest("Value-only workbook fixture is not available")


def _external_reference_workbook(test_case: unittest.TestCase) -> Path:
    """Create the fixture workbook containing original formula cells."""
    if os.environ.get("CONINCO_RUN_EXTERNAL_WORKBOOK_TESTS") != "1":
        test_case.skipTest("External reference workbook tests are opt-in")
    if not _EXTERNAL_REFERENCE_WORKBOOK.is_file():
        test_case.skipTest("Reference workbook is not available on this machine")
    return _EXTERNAL_REFERENCE_WORKBOOK


class WorkbookReportTest(unittest.TestCase):
    """Verify the complete workbook audit and comparison-report workflow."""
    # Confirm a valid PI formula is preserved instead of repaired.
    def test_valid_pi_formula_is_not_replaced_by_template_repair(self) -> None:
        expression = "K*PI()*D/1000000"
        self.assertEqual(
            (expression, ""),
            _repair_known_detail_expression(
                "S", expression, "square_round_reducer", "standard"
            ),
        )

    # Confirm comparison accepts lowercase dataclass-style input names.
    def test_formula_comparison_accepts_dataclass_lowercase_input_names(self) -> None:
        comparison = _formula_comparison(
            "=E6+F6",
            "S=W+H",
            {"w": 10, "h": 20},
            pd.Series([None, None, None, None, 10, 20]),
            {"W": 4, "H": 5},
            6,
        )

        self.assertEqual("equivalent", comparison["status"])
        self.assertEqual(30, comparison["original_formula_value"])
        self.assertEqual(30, comparison["detail_formula_value"])

    # Confirm original Excel values remain available when detail parsing fails.
    def test_formula_comparison_keeps_original_value_if_detail_fails(self) -> None:
        comparison = _formula_comparison(
            "=E6+F6",
            "S=W+UNKNOWN",
            {"w": 10, "h": 20},
            pd.Series([None, None, None, None, 10, 20]),
            {"W": 4, "H": 5},
            6,
        )

        self.assertEqual("unverifiable", comparison["status"])
        self.assertEqual(30, comparison["original_formula_value"])

    # Confirm original formulas can read referenced cells from another row.
    def test_formula_comparison_reads_referenced_cells_from_other_rows(self) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet["F5"] = 7
        worksheet["E6"] = 10

        comparison = _formula_comparison(
            "=E6+F5",
            "S=E6+F5",
            {},
            pd.Series([None, None, None, None, 10, 20]),
            {},
            6,
            worksheet,
        )

        self.assertEqual("equivalent", comparison["status"])
        self.assertEqual(17, comparison["original_formula_value"])

    # Confirm an S_V intermediate is compared as one complete block.
    def test_matching_intermediate_block_is_checked_as_one_whole_block(self) -> None:
        original = "=(N12*((2*(E12+F12)*(I12-1.5*G12))+0))"
        detail = "S_V2=2*(W'+H')*(L-1.5*W'')\nS=K*S_V2"
        row = pd.Series([None] * 14)
        columns = {"W1": 4, "H1": 5, "W2": 6, "L_M": 8, "K": 13}

        comparison = _formula_comparison(
            original,
            detail,
            {"w1": 100, "h1": 200, "w2": 50, "L_M": 400, "K": 2},
            row,
            columns,
            12,
        )

        block = next(item for item in comparison["blocks"] if item["label"] == "S_V2")
        self.assertEqual("matched", block["status"])
        self.assertTrue(block["color"])
        self.assertEqual([], block["matched_components"])
        self.assertTrue(block["original_spans"])
        span = block["original_spans"][0]
        self.assertIn(original[span["start"]:span["end"]], original)
        self.assertIn("I12-1.5*G12", original[span["start"]:span["end"]])

    # Confirm AST comparison maps G12 and W'' to the same logical dimension.
    def test_ast_difference_maps_excel_cells_to_formula_dimensions(self) -> None:
        source = "((2*PI()*((G12+G12/2)^2-G12/2^2)+2*PI()*H12*((G12+G12/2)+G12/2))*90/360)"
        detail = "(2*PI()*(((W''+W''/2))^2-(W''/2)^2)+2*PI()*H''*((W''+W''/2)+W''/2))*90/360"
        standard = {"blocks": [{"label": "S_V3", "original_label": "S_V3", "original_expression": detail}]}

        _formula_comparison(
            source,
            f"S_V3={detail}",
            {"w2": 100, "h2": 50},
            pd.Series([None] * 14),
            {"W2": 6, "H2": 7},
            12,
            standard_comparison=standard,
        )

        differences = standard["blocks"][0]["ast_differences"]
        self.assertEqual(1, len(differences))
        self.assertEqual("G12/2^2", differences[0]["original"])
        self.assertEqual("(W''/2)^2", differences[0]["standard"])

    # Confirm equivalent linear grouping is not reported as an AST error.
    def test_ast_difference_ignores_equivalent_linear_grouping(self) -> None:
        left = parse_formula("2*(W1+H1+4*a)*(L+2*a-1.5*(W2+2*a))")
        right = parse_formula("2*((W1+2*a)+(H1+2*a))*(L+2*a-1.5*(W2+2*a))")

        self.assertEqual([], _ast_difference_frontier(left, right))

    # Confirm S_V1 and S_V2 cannot be paired with each other.
    def test_ast_block_matching_does_not_cross_pair_v1_and_v2(self) -> None:
        standard_v1 = "(2*PI()*((W+W/2)^2-(W/2)^2)+2*PI()*H*((W+W/2)+W/2))*90/360"
        standard_v2 = "2*(W'+H')*(L-1.5*W'')"
        source = (
            "(2*PI()*((C12+C12/2)^2-(C12/2)^2)+"
            "2*PI()*D12*((C12+C12/2)+C12/2))*90/360+"
            "2*(E12+F12)*(I12-1.5*G12)"
        )
        standard = {
            "blocks": [
                {"label": "S_V1", "original_label": "S_V1", "original_expression": standard_v1},
                {"label": "S_V2", "original_label": "S_V2", "original_expression": standard_v2},
            ]
        }

        _formula_comparison(
            source,
            f"S_V1={standard_v1}\nS_V2={standard_v2}",
            {"w": 1, "h": 2, "w1": 3, "h1": 4, "w2": 5, "L_M": 6},
            pd.Series([None] * 14),
            {"W": 2, "H": 3, "W1": 4, "H1": 5, "W2": 6, "L_M": 8},
            12,
            standard_comparison=standard,
        )

        self.assertNotIn("ast_differences", standard["blocks"][0])
        self.assertNotIn("ast_differences", standard["blocks"][1])

    # Confirm reordered commutative terms remain equivalent.
    def test_ast_difference_ignores_commutative_block_order(self) -> None:
        original = parse_formula(
            "2*(W1+H1)*(L-1.5*W2)+PI()+(W2/2^2)"
        )
        standard = parse_formula(
            "PI()+(W2/2)^2+2*(W1+H1)*(L-1.5*W2)"
        )

        differences = _ast_difference_frontier(original, standard)

        self.assertEqual(1, len(differences))

    # Confirm an S_V2 error highlights its own source subtree only.
    def test_ast_intermediate_block_uses_its_source_subtree(self) -> None:
        source = (
            "(2*(E12+F12+5*K12)*(I12+2*K12-1.5*(G12+2*K12))"
            "+(2*PI()*((C12+C12/2)^2-(C12/2)^2)+2*PI()*D12*"
            "((C12+C12/2)+C12/2))*90/360)"
        )
        standard_v2 = "2*((W'+2a)+(H'+2a))*(L+2a-1.5*(W''+2a))"
        standard = {"blocks": [{
            "label": "S_V2",
            "original_label": "S_V2",
            "original_expression": standard_v2,
        }]}

        _formula_comparison(
            source,
            f"S_V2={standard_v2}",
            {"w1": 100, "h1": 200, "w2": 50, "L_M": 400, "a": 25},
            pd.Series([None] * 14),
            {"W": 2, "H": 3, "W1": 4, "H1": 5, "W2": 6, "L_M": 8, "a": 10},
            12,
            standard_comparison=standard,
        )

        block = standard["blocks"][0]
        self.assertIn("E12+F12+5*K12", block["ast_original_expression"])
        self.assertNotIn("PI()", block["ast_original_expression"])

    # Confirm unmatched SCC text still exposes useful matched subcomponents.
    def test_unmatched_scc_formula_still_matches_and_pairs_its_subblocks(self) -> None:
        original = "=(N24*((C24+E24+4*K24)*SQRT(I24^2+((D24-F24)/2)^2)+(D24+F24+4*K24)*SQRT(I24^2+(C24-E24/2)^2)))/1000000"
        detail = "SCC=(K*((W+W'+4a)*SQRT(L^2+((H-H')/2)^2)+(H+H'+4a)*SQRT(L^2+((W-W')/2)^2)))/1000000"
        row = pd.Series([None] * 14)
        columns = {"W": 2, "H": 3, "W1": 4, "H1": 5, "W2": 6, "L_M": 8, "a": 10, "K": 13}

        comparison = _formula_comparison(
            original,
            detail,
            {"w": 500, "h": 400, "w1": 300, "h1": 200, "w2": 100, "L_M": 600, "a": 25, "K": 2},
            row,
            columns,
            24,
        )

        block = comparison["blocks"][0]
        self.assertEqual("unmatched", block["status"])
        pairs = {(item["original_expression"], item["detail_expression"]) for item in block["matched_components"]}
        self.assertIn(("(C24+E24+4*K24)", "(W+W'+4a)"), pairs)
        self.assertIn(("(D24+F24+4*K24)", "(H+H'+4a)"), pairs)
        self.assertGreaterEqual(len(block["matched_components"]), 3)
        colors = [item["color"] for item in block["matched_components"]]
        self.assertEqual({"#B91C1C"}, set(colors))

    # Confirm one malformed block does not hide valid sibling blocks.
    def test_malformed_detail_assignment_does_not_hide_valid_sibling_blocks(self) -> None:
        original = "=(N39*((2*(E39+F39)*(I39-1.5*G39))+((2*PI()*((C39+C39/2)^2-(C39/2)^2)+2*PI()*D39*((C39+C39/2)+C39/2))*90/360)+((2*PI()*((G39+G39/2)^2-G39/2^2)+2*PI()*H39*((G39+G39/2)+G39/2))*90/360)-((D39*C39/2)+(H39*G39/2))))/1000000"
        detail = (
            "S_V1=(2*PI()*((W+W/2)^2-(W/2)^2)+2*PI()*H*((W+W/2)+W/2))*90/360\n"
            "S_V2=2(W'+H')*(L-1.5*W'')\n"
            "S_V3=(2*PI()*((W''+W''/2)^2-((W''/2)^2)+2*PI()*H''*((W''+W''/2)+W''/2))*90/360\n"
            "S_V4=(H*W/2)+(H''*W''/2)\n"
            "SCC/TA=(K*(S_V1+S_V2+S_V3-S_V4)/1000000"
        )
        row = pd.Series([None] * 14)
        columns = {"W": 2, "H": 3, "W1": 4, "H1": 5, "W2": 6, "H2": 7, "L_M": 8, "K": 13}

        comparison = _formula_comparison(
            original,
            detail,
            {"w": 800, "h": 400, "w1": 200, "h1": 100, "w2": 150, "h2": 75, "L_M": 300, "K": 2},
            row,
            columns,
            39,
        )

        blocks = {block["label"]: block for block in comparison["blocks"]}
        self.assertEqual("unverifiable", comparison["status"])
        self.assertTrue(blocks["S_V3"]["error"])
        self.assertTrue(blocks["SCC/TA"]["error"])
        self.assertEqual("matched", blocks["S_V1"]["status"])
        self.assertEqual("matched", blocks["S_V2"]["status"])
        self.assertEqual("matched", blocks["S_V4"]["status"])

    # Confirm source layout is retained while calculable rows are returned.
    def test_original_sheet_is_not_trimmed_and_calculates_rows(self) -> None:
        path = _value_only_workbook(self)
        workbook = read_raw_workbook(path)
        sheet = workbook["02"]

        self.assertEqual((61, 25), sheet.shape)
        calculated, values, errors = calculate_raw_sheet(sheet)

        self.assertEqual(15, len(calculated))
        self.assertEqual(15, len(values))
        self.assertEqual([], errors)
        self.assertEqual(5.325, values[20]["K"])
        self.assertEqual(15.12, values[20]["S"])

    # Confirm export preserves layout and inserts calculated result columns.
    def test_export_keeps_layout_and_adds_red_result_columns(self) -> None:
        source = _value_only_workbook(self)

        with TemporaryDirectory() as temporary_name:
            output = Path(temporary_name) / "calculated.xlsx"
            result = export_calculated_workbook(source, output, "02")
            workbook = load_workbook(output)
            sheet = workbook["02"]

            self.assertEqual(15, result["calculated_count"])
            self.assertEqual("K\n(tính bằng máy)", sheet["R12"].value)
            self.assertEqual("S\n(cách nhiệt tính bằng máy)", sheet["W12"].value)
            self.assertEqual("SCC/STA\n(cách nhiệt trong tính bằng máy)", sheet["Y12"].value)
            self.assertEqual(5.325, sheet["R21"].value)
            self.assertEqual(15.12, sheet["W21"].value)
            self.assertEqual(0, sheet["Y21"].value)
            self.assertEqual("FF008000", sheet["R21"].font.color.rgb)
            self.assertEqual("FF008000", sheet["W42"].font.color.rgb)
            self.assertEqual(28, sheet.column_dimensions["W"].width)
            self.assertEqual(34, sheet.column_dimensions["Y"].width)
            self.assertGreaterEqual(sheet.row_dimensions[12].height, 42)
            self.assertIn("S12:U12", {str(value) for value in sheet.merged_cells.ranges})

    # Confirm IF/AND detail formulas are evaluated and audited correctly.
    def test_if_and_formula_is_evaluated_from_detail_text(self) -> None:
        source = _value_only_workbook(self)
        workbook = read_raw_workbook(source)
        _calculated, values, _errors = calculate_raw_sheet(workbook["02"])

        audits = [
            audit
            for row_values in values.values()
            for audit in row_values["formula_audits"].values()
            if audit["label"] == "SCN" and audit["expression"].startswith("IF(")
        ]
        self.assertTrue(audits)
        self.assertTrue(all(audit["contractor_value"] is not None for audit in audits))
        self.assertTrue(all(audit["matches"] for audit in audits))
        self.assertAlmostEqual(0.84, audits[0]["contractor_value"])

    # Confirm formula errors remain visible without stopping other rows.
    def test_formula_error_is_visible_as_mismatch_without_stopping_other_rows(self) -> None:
        source = _value_only_workbook(self)
        workbook = read_raw_workbook(source)
        calculated, values, errors = calculate_raw_sheet(workbook["02"])

        self.assertEqual(15, len(calculated))
        self.assertEqual([], errors)
        broken_audits = [
            audit
            for row_values in values.values()
            for audit in row_values["formula_audits"].values()
            if audit["error"]
        ]
        self.assertTrue(broken_audits)
        # A missing reference formula is diagnostic evidence only.  The
        # saved quantity remains authoritative, so a matching result stays
        # green instead of becoming red just because the detail is unavailable.
        self.assertTrue(all(audit["matches"] for audit in broken_audits))

    # Confirm inserted output columns do not corrupt original K references.
    def test_export_formula_keeps_original_k_reference_after_insertions(self) -> None:
        source = _external_reference_workbook(self)

        with TemporaryDirectory() as temporary_name:
            output = Path(temporary_name) / "reference.xlsx"
            export_calculated_workbook(source, output, "COMBINE")
            sheet = load_workbook(output, data_only=False)["COMBINE"]

            # T is the inserted outer-S result; N remains the original K cell.
            self.assertIn("N6", str(sheet["T6"].value))
            self.assertNotIn("O6", str(sheet["T6"].value))

    # Confirm the exported workbook contains the formula audit table.
    def test_mismatch_table_lists_detail_vs_original(self) -> None:
        source = _external_reference_workbook(self)

        with TemporaryDirectory() as temporary_name:
            output = Path(temporary_name) / "reference.xlsx"
            export_calculated_workbook(source, output, "COMBINE")
            sheet = load_workbook(output, data_only=False)["COMBINE"]
            self.assertTrue(
                any(
                    sheet.cell(row, 3).value == "Formula audit"
                    for row in range(1, sheet.max_row + 1)
                )
            )

    # Confirm written formula precedence is respected and differences are flagged.
    def test_detail_formulas_keep_written_precedence_and_flag_differences(self) -> None:
        source = _external_reference_workbook(self)

        workbook = read_raw_workbook(source)
        _calculated, values, errors = calculate_raw_sheet(workbook["COMBINE"])

        self.assertEqual([], errors)
        mismatches = [
            source_row + 1
            for source_row, row_values in values.items()
            if not all(row_values["matches"].values())
        ]
        self.assertEqual([12, 24, 26, 38, 39], mismatches)

        # Row 12 is the important precedence case: the detail writes
        # (W''/2)^2, while the original Excel cell uses W''/2^2.
        row_12 = values[11]
        self.assertAlmostEqual(8.168981633974482, row_12["S"])
        self.assertFalse(row_12["matches"]["S"])
        self.assertNotIn("source Excel precedence", row_12["formula_audits"]["S"].get("correction", ""))

    # Confirm preview formula blocks appear in S_V1 through S_V4 order.
    def test_preview_contains_ordered_formula_blocks_for_mismatches(self) -> None:
        source = _external_reference_workbook(self)

        with TemporaryDirectory() as temporary_name:
            result = export_calculated_workbook(source, Path(temporary_name) / "reference.xlsx", "COMBINE")
            with Path(result["preview"]).open(encoding="utf-8", newline="") as stream:
                rows = list(csv.DictReader(stream, delimiter="\t"))

            row_12 = next(row for row in rows if row["source_row"] == "12")
            audits = json.loads(row_12["formula_audits"])
            audit = next(item for item in audits if item["target"] == "S")
            detail = audit["detail_block"]
            self.assertLess(detail.index("S_V1"), detail.index("S_V2"))
            self.assertLess(detail.index("S_V2"), detail.index("S_V3"))
            self.assertLess(detail.index("S_V3"), detail.index("S_V4"))
            self.assertLess(detail.index("S_V4"), detail.index("S ="))

    # Confirm SCN maps to the outer insulation quantity.
    def test_scn_maps_to_outer_insulation_not_fireproof(self) -> None:
        source = _external_reference_workbook(self)

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[20]  # Excel row 21: SCN is the outer-insulation formula.
        self.assertEqual("SCN", row["formula_audits"]["S"]["label"])
        self.assertAlmostEqual(95.976, row["S"])

    # Confirm SCC maps to the fireproof quantity.
    def test_scc_maps_to_fireproof_quantity(self) -> None:
        source = _external_reference_workbook(self)

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[22]  # Excel row 23: SCC is the fireproof formula.
        self.assertEqual("SCC", row["formula_audits"]["S_fireproof"]["label"])
        self.assertAlmostEqual(1.1881, row["S_fireproof"])

    # Confirm the end-cap SCN repair produces a usable result.
    def test_end_cap_scn_uses_repaired_formula_instead_of_zero(self) -> None:
        source = _external_reference_workbook(self)

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[22]  # Excel row 23: Bít đầu.
        self.assertAlmostEqual(1.1025, row["S"])
        self.assertFalse(row["formula_audits"]["S"]["error"])

    # Confirm SCC/TA writes one shared result to both inside outputs.
    def test_scc_ta_uses_one_shared_result_for_both_inside_columns(self) -> None:
        source = _external_reference_workbook(self)

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[43]  # Excel row 44: SCC/TA has only the source S value.
        self.assertEqual("SCC/TA", row["formula_audits"]["S_fireproof"]["label"])
        self.assertNotIn("S_inner", row["formula_audits"])
        self.assertEqual(row["S_fireproof"], row["S_inner"])

    # Confirm shared tee output is written once and S_V3 is audited separately.
    def test_shared_tee_formula_is_exported_once_and_s_v3_is_grouped(self) -> None:
        source = _external_reference_workbook(self)

        from excel_engine.formula_catalog import build_formula_catalog

        catalog = build_formula_catalog(source, "COMBINE")
        _calculated, values, errors = calculate_raw_sheet(
            read_raw_workbook(source)["COMBINE"],
            reference_catalog=catalog,
        )

        self.assertEqual([], errors)
        row = values[38]  # Excel row 39: shared SCC/TA tee.
        self.assertEqual("SCC/TA", row["formula_audits"]["S_fireproof"]["label"])
        self.assertAlmostEqual(7.728981633974482, row["S_fireproof"])
        self.assertAlmostEqual(row["S_fireproof"], row["S_inner"])
        self.assertEqual("S_fireproof", row["shared_target"])
        blocks = {
            block["label"]: block
            for block in row["formula_audits"]["S_fireproof"]["formula_comparison"]["blocks"]
        }
        self.assertEqual("matched", blocks["S_V1"]["status"])
        self.assertEqual("matched", blocks["S_V2"]["status"])
        self.assertEqual("unmatched", blocks["S_V3"]["status"])
        self.assertEqual("matched", blocks["S_V4"]["status"])

    # Confirm STA maps only to the acoustic inner quantity.
    def test_sta_maps_only_to_acoustic_inner_quantity(self) -> None:
        source = _external_reference_workbook(self)

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[35]  # Excel row 36: STA is the inner/acoustic formula.
        self.assertEqual("STA", row["formula_audits"]["S_inner"]["label"])
        self.assertAlmostEqual(1.0, row["S_inner"])
        self.assertEqual("SCN", row["formula_audits"]["S"]["label"])
        self.assertAlmostEqual(0.0, row["S"])

    # Confirm a damper note does not override the SCN mapping.
    def test_damper_note_does_not_overwrite_scn_formula(self) -> None:
        source = _external_reference_workbook(self)

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[29]  # Excel row 30 has SCN plus the note "Bao gồm trong van".
        self.assertEqual("SCN", row["formula_audits"]["S"]["label"])
        self.assertAlmostEqual(1.7, row["S"])

    # Confirm square-elbow SCC/TA uses the source area grouping.
    def test_square_elbow_scc_ta_uses_source_area_grouping(self) -> None:
        source = _external_reference_workbook(self)

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[41]  # Excel row 42: square elbow.
        self.assertAlmostEqual(0.652, row["S"])
        self.assertAlmostEqual(3.6, row["S_fireproof"])
        self.assertAlmostEqual(3.6, row["S_inner"])
        self.assertIn("Moved the leading 2", row["formula_audits"]["S_fireproof"]["correction"])


if __name__ == "__main__":
    unittest.main()
