"""JSON entry point for workbook inspection and formula calculation."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from .contractor_processor import process_workbook
from .excel_reader import read_workbook
from .formula_engine import calculate
from .workbook_report import export_calculated_workbook


def inspect_file(path: str) -> dict[str, Any]:
    workbook_path = Path(path).expanduser().resolve()
    workbook = read_workbook(workbook_path)
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


def process_request(request: dict[str, Any]) -> dict[str, Any]:
    if request.get("action") == "list_sheets":
        try:
            result = inspect_file(request["path"])
            return {"ok": True, "result": result, "files": [], "errors": []}
        except (KeyError, TypeError, ValueError, OSError) as error:
            return {"ok": False, "files": [], "errors": [{"message": str(error)}]}

    if request.get("action") == "export_calculated_workbook":
        try:
            result = export_calculated_workbook(
                request["path"],
                request.get("output"),
                request.get("sheet", "02"),
            )
            return {
                "ok": True,
                "result": result,
                "files": [result["output"]],
                "errors": result["errors"],
            }
        except (KeyError, TypeError, ValueError) as error:
            return {"ok": False, "files": [], "errors": [{"message": str(error)}]}

    if request.get("action") == "process_contractor":
        try:
            result = process_workbook(
                request["path"],
                request.get("sheet"),
                request.get("section"),
                request.get("product_type"),
                request.get("header_map"),
            )
            return {"ok": True, "result": result, "files": [], "errors": []}
        except (KeyError, TypeError, ValueError) as error:
            return {"ok": False, "files": [], "errors": [{"message": str(error)}]}

    if request.get("action") == "calculate":
        try:
            result = calculate(
                request["section"],
                request["product_type"],
                request["inputs"],
            )
            return {"ok": True, "result": result.to_dict(), "files": [], "errors": []}
        except (KeyError, TypeError, ValueError) as error:
            return {"ok": False, "files": [], "errors": [{"message": str(error)}]}

    paths = request.get("files")
    if not isinstance(paths, list) or not paths:
        return {"ok": False, "files": [], "errors": [{"message": "files must be a non-empty list."}]}

    files = []
    errors = []
    for path in paths:
        try:
            files.append(inspect_file(path))
        except Exception as error:  # noqa: BLE001 - keep the process boundary JSON-only.
            errors.append({"path": str(path), "message": str(error)})

    return {"ok": not errors, "files": files, "errors": errors}


def main() -> int:
    try:
        request = json.load(sys.stdin)
        response = process_request(request)
    except json.JSONDecodeError as error:
        response = {"ok": False, "files": [], "errors": [{"message": f"Invalid JSON: {error}"}]}
    except Exception as error:  # noqa: BLE001 - keep the process boundary JSON-only.
        response = {"ok": False, "files": [], "errors": [{"message": str(error)}]}

    print(json.dumps(response, ensure_ascii=False))
    return 0 if response["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
