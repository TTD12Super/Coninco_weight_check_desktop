# Python engine

This directory contains the offline Excel processing engine used by the
React/Tauri desktop application.

Implemented module boundaries:

- `excel_reader.py`: load `.xlsx` and `.xls` workbooks and worksheets. **Implemented.**
- `formula_engine.py`: calculate the workbook's R/S/T areas. **Implemented.**
- `product_formatter.py`: normalize section/product type, build `FormulaInputs`, and call the formula engine. **Implemented.**
- `contractor_processor.py`: read contractor rows, map headers, classify products, and apply the formula engine. **Implemented.**
- `workbook_report.py`: read the original workbook without skipping rows or dropping columns, then export calculated columns. **Implemented.**
- `quantity_calculator.py`: calculate `K_By_function` from `plan`, `up_down`, and `vertical_axis`. **Implemented.**
- `normalizer.py`: normalize headers, values, dates, and numbers.
- `reconciler.py`: compare records across imported files.
- `validators.py`: apply configurable business rules.
- `report_writer.py`: export comparison results.
- `main.py`: future command-line processing entry point.

## Test the reader

Create a virtual environment and install the Python dependencies:

```bash
python -m venv python-engine/.venv
python-engine/.venv/bin/python -m pip install -r python-engine/requirements.txt
```

Run the automated reader tests:

```bash
PYTHONPATH=python-engine/src python-engine/.venv/bin/python \
  -m unittest discover -s python-engine/tests -p 'test_*.py'
```

Read and summarize a workbook manually:

```bash
PYTHONPATH=python-engine/src python-engine/.venv/bin/python \
  -m excel_engine.excel_reader /absolute/path/to/workbook.xlsx
```

The Tauri desktop application uses the JSON reader entry point when a file is
added. To test that process boundary directly:

```bash
printf '%s' '{"files":["/absolute/path/to/workbook.xlsx"]}' | \
  PYTHONPATH=python-engine/src python-engine/.venv/bin/python \
  -m excel_engine.main
```

The reader returns a dictionary of pandas DataFrames, one DataFrame per
worksheet. The contractor processor uses the formula text on each row first
to choose `standard`, `fireproof`, or `acoustic`, then uses the product
description to choose the product formula. Unsupported or unclassified rows
are returned in the `errors` list.

## Process contractor Excel rows

The section and product type can still be supplied explicitly when a project
needs to override the row classification:

```bash
printf '%s' '{"action":"process_contractor","path":"/absolute/path/contractor.xlsx","sheet":"Input","section":"standard","product_type":"rectangular_duct","header_map":{"Q":"Joint","K":"Qty"}}' | \
  PYTHONPATH=python-engine/src python-engine/.venv/bin/python \
  -m excel_engine.main
```

Rows that cannot be classified or calculated are returned in the `errors`
list; they are not silently calculated with a guessed product type.

## Export a full workbook with calculated columns

This keeps all original rows and columns, then adds bold calculated columns
beside the original K, S, and inner-S columns:

```bash
printf '%s' '{"action":"export_calculated_workbook","path":"/absolute/path/L01GF0087.xls","sheet":"02"}' | \
  PYTHONPATH=python-engine/src python-engine/.venv/bin/python \
  -m excel_engine.main
```

The default output is `<input-name>_calculated.xlsx`. Rows without numeric
geometry remain in the workbook and receive blank calculated cells.

## Test one calculation

```bash
printf '%s' '{"action":"calculate","section":"standard","product_type":"rectangular_duct","inputs":{"W":800,"H":500,"Q":11,"K":10.88}}' | \
  PYTHONPATH=python-engine/src python-engine/.venv/bin/python \
  -m excel_engine.main
```
