import csv
import json
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from openpyxl import Workbook, load_workbook

from excel_engine.workbook_report import (
    _formula_comparison,
    _repair_known_detail_expression,
    calculate_raw_sheet,
    export_calculated_workbook,
    read_raw_workbook,
)


class WorkbookReportTest(unittest.TestCase):
    def test_valid_pi_formula_is_not_replaced_by_template_repair(self) -> None:
        expression = "K*PI()*D/1000000"
        self.assertEqual(
            (expression, ""),
            _repair_known_detail_expression(
                "S", expression, "square_round_reducer", "standard"
            ),
        )

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
        self.assertEqual(len(colors), len(set(colors)))

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

    def test_original_sheet_is_not_trimmed_and_calculates_rows(self) -> None:
        path = Path(__file__).parents[1] / "src" / "excel_engine" / "L01GF0087_value_only.xls"
        workbook = read_raw_workbook(path)
        sheet = workbook["02"]

        self.assertEqual((61, 25), sheet.shape)
        calculated, values, errors = calculate_raw_sheet(sheet)

        self.assertEqual(15, len(calculated))
        self.assertEqual(15, len(values))
        self.assertEqual([], errors)
        self.assertEqual(5.325, values[20]["K"])
        self.assertEqual(15.12, values[20]["S"])

    def test_export_keeps_layout_and_adds_red_result_columns(self) -> None:
        source = Path(__file__).parents[1] / "src" / "excel_engine" / "L01GF0087_value_only.xls"

        with TemporaryDirectory() as temporary_name:
            output = Path(temporary_name) / "calculated.xlsx"
            result = export_calculated_workbook(source, output, "02")
            workbook = load_workbook(output)
            sheet = workbook["02"]

            self.assertEqual(15, result["calculated_count"])
            self.assertEqual("K\n(tính bằng máy)", sheet["R12"].value)
            self.assertEqual("S\n(cách nhiệt tính bằng máy)", sheet["W12"].value)
            self.assertEqual("S\n(cách nhiệt trong tính bằng máy)", sheet["Y12"].value)
            self.assertEqual(5.325, sheet["R21"].value)
            self.assertTrue(str(sheet["W21"].value).startswith("="))
            self.assertIn("F21", sheet["W21"].value)
            self.assertIn("G21", sheet["W21"].value)
            self.assertEqual(0, sheet["Y21"].value)
            self.assertEqual("FF008000", sheet["R21"].font.color.rgb)
            self.assertEqual("FF008000", sheet["W42"].font.color.rgb)
            self.assertEqual(28, sheet.column_dimensions["W"].width)
            self.assertEqual(34, sheet.column_dimensions["Y"].width)
            self.assertGreaterEqual(sheet.row_dimensions[12].height, 42)
            self.assertIn("S12:U12", {str(value) for value in sheet.merged_cells.ranges})

    def test_if_and_formula_is_evaluated_from_detail_text(self) -> None:
        source = Path(__file__).parents[1] / "src" / "excel_engine" / "L01GF0087_value_only.xls"
        workbook = read_raw_workbook(source)
        _calculated, values, _errors = calculate_raw_sheet(workbook["02"])

        audits = [
            audit
            for row_values in values.values()
            for audit in row_values["formula_audits"].values()
            if audit["label"] == "SCN" and audit["expression"].startswith("IF(")
        ]
        self.assertTrue(audits)
        self.assertTrue(all(not audit["error"] for audit in audits))
        self.assertAlmostEqual(0.84, audits[0]["value"])

    def test_formula_error_is_visible_as_mismatch_without_stopping_other_rows(self) -> None:
        source = Path(__file__).parents[1] / "src" / "excel_engine" / "L01GF0087_value_only.xls"
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
        self.assertTrue(all(not audit["matches"] for audit in broken_audits))
        for row_values in values.values():
            for target, audit in row_values["formula_audits"].items():
                if audit["error"]:
                    output_key = {"S": "S", "S_fireproof": "S_fireproof", "S_inner": "S_inner"}[target]
                    self.assertEqual(0.0, row_values[output_key])

    def test_export_formula_keeps_original_k_reference_after_insertions(self) -> None:
        source = Path("/Users/trantrungdung/Downloads/4. Form tính tiêu âm, cách nhiệt, chống cháy combine (2) (1).xlsx")
        if not source.is_file():
            self.skipTest("Reference workbook is not available on this machine")

        with TemporaryDirectory() as temporary_name:
            output = Path(temporary_name) / "reference.xlsx"
            export_calculated_workbook(source, output, "COMBINE")
            sheet = load_workbook(output, data_only=False)["COMBINE"]

            # T is the inserted outer-S result; N remains the original K cell.
            self.assertIn("N6", str(sheet["T6"].value))
            self.assertNotIn("O6", str(sheet["T6"].value))

    def test_mismatch_table_lists_detail_vs_original(self) -> None:
        source = Path("/Users/trantrungdung/Downloads/4. Form tính tiêu âm, cách nhiệt, chống cháy combine (2) (1).xlsx")
        if not source.is_file():
            self.skipTest("Reference workbook is not available on this machine")

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

    def test_detail_formulas_keep_written_precedence_and_flag_differences(self) -> None:
        source = Path("/Users/trantrungdung/Downloads/4. Form tính tiêu âm, cách nhiệt, chống cháy combine (2) (1).xlsx")
        if not source.is_file():
            self.skipTest("Reference workbook is not available on this machine")

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

    def test_preview_contains_ordered_formula_blocks_for_mismatches(self) -> None:
        source = Path("/Users/trantrungdung/Downloads/4. Form tính tiêu âm, cách nhiệt, chống cháy combine (2) (1).xlsx")
        if not source.is_file():
            self.skipTest("Reference workbook is not available on this machine")

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

    def test_scn_maps_to_outer_insulation_not_fireproof(self) -> None:
        source = Path("/Users/trantrungdung/Downloads/4. Form tính tiêu âm, cách nhiệt, chống cháy combine (2) (1).xlsx")
        if not source.is_file():
            self.skipTest("Reference workbook is not available on this machine")

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[20]  # Excel row 21: SCN is the outer-insulation formula.
        self.assertEqual("SCN", row["formula_audits"]["S"]["label"])
        self.assertAlmostEqual(95.976, row["S"])

    def test_scc_maps_to_fireproof_quantity(self) -> None:
        source = Path("/Users/trantrungdung/Downloads/4. Form tính tiêu âm, cách nhiệt, chống cháy combine (2) (1).xlsx")
        if not source.is_file():
            self.skipTest("Reference workbook is not available on this machine")

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[22]  # Excel row 23: SCC is the fireproof formula.
        self.assertEqual("SCC", row["formula_audits"]["S_fireproof"]["label"])
        self.assertAlmostEqual(1.1881, row["S_fireproof"])

    def test_end_cap_scn_uses_repaired_formula_instead_of_zero(self) -> None:
        source = Path("/Users/trantrungdung/Downloads/4. Form tính tiêu âm, cách nhiệt, chống cháy combine (2) (1).xlsx")
        if not source.is_file():
            self.skipTest("Reference workbook is not available on this machine")

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[22]  # Excel row 23: Bít đầu.
        self.assertAlmostEqual(1.1025, row["S"])
        self.assertFalse(row["formula_audits"]["S"]["error"])

    def test_scc_ta_does_not_fill_blank_companion_column(self) -> None:
        source = Path("/Users/trantrungdung/Downloads/4. Form tính tiêu âm, cách nhiệt, chống cháy combine (2) (1).xlsx")
        if not source.is_file():
            self.skipTest("Reference workbook is not available on this machine")

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[43]  # Excel row 44: SCC/TA has only the source S value.
        self.assertEqual("SCC/TA", row["formula_audits"]["S_fireproof"]["label"])
        self.assertNotIn("S_inner", row["formula_audits"])
        self.assertEqual(0.0, row["S_inner"])

    def test_sta_maps_only_to_acoustic_inner_quantity(self) -> None:
        source = Path("/Users/trantrungdung/Downloads/4. Form tính tiêu âm, cách nhiệt, chống cháy combine (2) (1).xlsx")
        if not source.is_file():
            self.skipTest("Reference workbook is not available on this machine")

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[35]  # Excel row 36: STA is the inner/acoustic formula.
        self.assertEqual("STA", row["formula_audits"]["S_inner"]["label"])
        self.assertAlmostEqual(1.0, row["S_inner"])
        self.assertEqual("SCN", row["formula_audits"]["S"]["label"])
        self.assertAlmostEqual(0.0, row["S"])

    def test_damper_note_does_not_overwrite_scn_formula(self) -> None:
        source = Path("/Users/trantrungdung/Downloads/4. Form tính tiêu âm, cách nhiệt, chống cháy combine (2) (1).xlsx")
        if not source.is_file():
            self.skipTest("Reference workbook is not available on this machine")

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[29]  # Excel row 30 has SCN plus the note "Bao gồm trong van".
        self.assertEqual("SCN", row["formula_audits"]["S"]["label"])
        self.assertAlmostEqual(1.7, row["S"])

    def test_square_elbow_scc_ta_uses_source_area_grouping(self) -> None:
        source = Path("/Users/trantrungdung/Downloads/4. Form tính tiêu âm, cách nhiệt, chống cháy combine (2) (1).xlsx")
        if not source.is_file():
            self.skipTest("Reference workbook is not available on this machine")

        _calculated, values, _errors = calculate_raw_sheet(read_raw_workbook(source)["COMBINE"])
        row = values[41]  # Excel row 42: square elbow.
        self.assertAlmostEqual(0.652, row["S"])
        self.assertAlmostEqual(3.6, row["S_fireproof"])
        self.assertAlmostEqual(3.6, row["S_inner"])
        self.assertIn("Moved the leading 2", row["formula_audits"]["S_fireproof"]["correction"])


if __name__ == "__main__":
    unittest.main()
