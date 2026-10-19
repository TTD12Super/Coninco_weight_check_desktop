# excel_engine package

The package reads `.xlsx` and `.xls` workbooks, calculates the contractor rows,
and exports a new `.xlsx` report without removing original rows or sheets.

The export workflow automatically:

- converts `.xls` to a temporary `.xlsx` with LibreOffice;
- keeps the original workbook's layout, merged cells, fonts, sizes, and styles;
- adds calculated `K`, outer `S`, and inner `S` columns beside the original
  columns;
- writes the calculated cells in red.

The JSON entry point accepts:

```json
{
  "action": "export_calculated_workbook",
  "path": "/path/to/contractor.xls",
  "output": "/path/to/contractor_calculated.xlsx",
  "sheet": "02"
}
```
