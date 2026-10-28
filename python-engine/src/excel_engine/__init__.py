"""Python components for the CONINCO Excel Auditor.

The package is split into small workflow layers:

* ``excel_reader`` opens supported workbook formats.
* ``input_formats`` detects contractor columns and quantity layouts.
* ``product_formatter`` maps each row to a product type and section.
* ``formula_catalog`` extracts the authoritative formulas from the reference
  workbook.
* ``formula_parser`` parses, evaluates, and compares contractor formula text.
* ``workbook_report`` coordinates Run Audit and writes the result workbook.
* ``main`` is the JSON request boundary used by the desktop application.
"""
