"""JSON entry point for workbook inspection and formula calculation."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from .contractor_processor import process_workbook
from .excel_reader import read_workbook
from .formula_engine import calculate
from .formula_catalog import build_formula_catalog, list_reference_sheets
from .workbook_report import export_calculated_workbook


# Purpose: Inspect workbook sheets without calculating any formula.
# Example output: {"path": "/tmp/input.xlsx", "sheets": [{"name": "02", "rows": 61, "columns": 25}]}
def inspect_file(path: str) -> dict[str, Any]:
    """Inspect workbook sheets without calculating any formula."""
    # Compute and store: workbook_path.
    workbook_path = Path(path).expanduser().resolve()
    # Compute and store: workbook.
    workbook = read_workbook(workbook_path)
    # Return the value produced by this step.
    return {
        "path": str(workbook_path),
        "sheets": [
            {
                "name": sheet_name,
                "rows": len(data),
                "columns": len(data.columns),
                "column_names": [str(column) for column in data.columns],
            }
            for sheet_name, data in workbook.items()
        ],
    }


# Purpose: Dispatch one JSON request to the matching engine workflow.
# Example output: {"ok": true, "result": {...}, "files": [], "errors": []}
def process_request(request: dict[str, Any]) -> dict[str, Any]:
    """Dispatch one JSON request to the matching engine workflow."""
    # Check this condition before choosing the next branch.
    if request.get("action") == "list_sheets":
        # Step 1: return the submitted workbook's worksheet names and shape.
        try:
            # Compute and store: result.
            result = inspect_file(request["path"])
            # Return the value produced by this step.
            return {"ok": True, "result": result, "files": [], "errors": []}
        except (KeyError, TypeError, ValueError, OSError) as error:
            # Return the value produced by this step.
            return {"ok": False, "files": [], "errors": [{"message": str(error)}]}

    # Check this condition before choosing the next branch.
    if request.get("action") == "list_reference_sheets":
        # Step 2: list worksheets so the user can choose the reference sheet.
        try:
            # Compute and store: sheets.
            sheets = list_reference_sheets(request["path"])
            # Return the value produced by this step.
            return {
                "ok": True,
                "result": {"sheets": [{"name": name} for name in sheets]},
                "files": [],
                "errors": [],
            }
        except (KeyError, TypeError, ValueError, OSError) as error:
            # Return the value produced by this step.
            return {"ok": False, "files": [], "errors": [{"message": str(error)}]}

    # Check this condition before choosing the next branch.
    if request.get("action") == "load_reference_catalog":
        # Step 3: extract standard product formulas from the selected sheet.
        try:
            # Compute and store: catalog.
            catalog = build_formula_catalog(request["path"], request.get("sheet"))
            # Return the value produced by this step.
            return {
                "ok": True,
                "result": {
                    "sheet": catalog["sheet"],
                    "entries": catalog["entries"],
                    "count": catalog["count"],
                },
                "files": [],
                "errors": [],
            }
        except (KeyError, TypeError, ValueError, OSError) as error:
            # Return the value produced by this step.
            return {"ok": False, "files": [], "errors": [{"message": str(error)}]}

    # Check this condition before choosing the next branch.
    if request.get("action") == "export_calculated_workbook":
        # Step 4: run the audit and write the calculated workbook.
        try:
            # Compute and store: result.
            result = export_calculated_workbook(
                request["path"],
                request.get("output"),
                request.get("sheet", "02"),
                request.get("reference_path"),
                request.get("reference_sheet"),
                request.get("reference_catalog"),
            )
            # Return the value produced by this step.
            return {
                "ok": True,
                "result": result,
                "files": [result["output"]],
                "errors": result["errors"],
            }
        except (KeyError, TypeError, ValueError) as error:
            # Return the value produced by this step.
            return {"ok": False, "files": [], "errors": [{"message": str(error)}]}

    # Check this condition before choosing the next branch.
    if request.get("action") == "process_contractor":
        # Legacy workflow: process a contractor workbook directly.
        try:
            # Compute and store: result.
            result = process_workbook(
                request["path"],
                request.get("sheet"),
                request.get("section"),
                request.get("product_type"),
                request.get("header_map"),
            )
            # Return the value produced by this step.
            return {"ok": True, "result": result, "files": [], "errors": []}
        except (KeyError, TypeError, ValueError) as error:
            # Return the value produced by this step.
            return {"ok": False, "files": [], "errors": [{"message": str(error)}]}

    # Check this condition before choosing the next branch.
    if request.get("action") == "calculate":
        # Legacy workflow: calculate one product from explicit inputs.
        try:
            # Compute and store: result.
            result = calculate(
                request["section"],
                request["product_type"],
                request["inputs"],
            )
            # Return the value produced by this step.
            return {"ok": True, "result": result.to_dict(), "files": [], "errors": []}
        except (KeyError, TypeError, ValueError) as error:
            # Return the value produced by this step.
            return {"ok": False, "files": [], "errors": [{"message": str(error)}]}

    # Compute and store: paths.
    paths = request.get("files")
    # Check this condition before choosing the next branch.
    if not isinstance(paths, list) or not paths:
        # Return the value produced by this step.
        return {"ok": False, "files": [], "errors": [{"message": "files must be a non-empty list."}]}

    # Compute and store: files.
    files = []
    # Compute and store: errors.
    errors = []
    # Process each item in this input sequence.
    for path in paths:
        # Protect parsing/evaluation so bad workbook data becomes a controlled error.
        try:
            # Execute this expression for its result or side effect.
            files.append(inspect_file(path))
        except Exception as error:  # noqa: BLE001 - keep the process boundary JSON-only.
            # Execute this expression for its result or side effect.
            errors.append({"path": str(path), "message": str(error)})

    # Return the value produced by this step.
    return {"ok": not errors, "files": files, "errors": errors}


# Purpose: Read one JSON request, execute it, and print one JSON response.
# Example output: 0
def main() -> int:
    """Read one JSON request, execute it, and print one JSON response."""
    # Protect parsing/evaluation so bad workbook data becomes a controlled error.
    try:
        # Compute and store: request.
        request = json.load(sys.stdin)
        # Compute and store: response.
        response = process_request(request)
    except json.JSONDecodeError as error:
        # Compute and store: response.
        response = {"ok": False, "files": [], "errors": [{"message": f"Invalid JSON: {error}"}]}
    except Exception as error:  # noqa: BLE001 - keep the process boundary JSON-only.
        # Compute and store: response.
        response = {"ok": False, "files": [], "errors": [{"message": str(error)}]}

    # Execute this expression for its result or side effect.
    print(json.dumps(response, ensure_ascii=False))
    # Return the value produced by this step.
    return 0 if response["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
