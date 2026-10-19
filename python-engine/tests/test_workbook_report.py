import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook

from excel_engine.workbook_report import calculate_raw_sheet, read_raw_workbook
from excel_engine.workbook_report import export_calculated_workbook


class WorkbookReportTest(unittest.TestCase):
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
            self.assertEqual(15.12, sheet["W21"].value)
            self.assertEqual(0, sheet["Y21"].value)
            self.assertEqual("FF008000", sheet["R21"].font.color.rgb)
            self.assertEqual("FFFF0000", sheet["W58"].font.color.rgb)
            self.assertEqual(28, sheet.column_dimensions["W"].width)
            self.assertEqual(34, sheet.column_dimensions["Y"].width)
            self.assertGreaterEqual(sheet.row_dimensions[12].height, 42)
            self.assertIn("S12:U12", {str(value) for value in sheet.merged_cells.ranges})


if __name__ == "__main__":
    unittest.main()
