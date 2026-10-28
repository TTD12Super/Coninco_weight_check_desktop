"""Names and header detection used by contractor workbook layouts."""

from __future__ import annotations

import re
import unicodedata
from typing import Any

import pandas as pd

FORMULA_INPUT_FIELDS = (
    "W", "H", "W1", "H1", "W2", "H2", "L_M", "R_D", "a", "t", "Q", "K"
)

SHEET_02_COLUMN_MAP = {
    "Unnamed: 0": "section_marker",
    "Unnamed: 1": "item_no",
    "Unnamed: 2": "description",
    "Unnamed: 3": "unit",
    "Unnamed: 5": "W",
    "Unnamed: 6": "H",
    "Unnamed: 7": "W1",
    "Unnamed: 8": "H1",
    "Unnamed: 9": "W2",
    "Unnamed: 10": "H2",
    "Unnamed: 11": "L_M",
    "Unnamed: 12": "R_D",
    "Unnamed: 13": "a",
    "Unnamed: 14": "t",
    "Unnamed: 15": "Q",
    "Unnamed: 16": "K",
}

DEFAULT_SOURCE_COLUMNS = {
    "section_marker": 0, "item_no": 1, "description": 2, "unit": 3,
    "unit_duplicate": 4, "W": 5, "H": 6, "W1": 7, "H1": 8,
    "W2": 9, "H2": 10, "L_M": 11, "R_D": 12, "a": 13, "t": 14,
    "Q": 15, "K": 16, "plan": 17, "up_down": 18, "vertical_axis": 19,
    "insulation_quantity": 20, "acoustic_quantity": 21,
    "formula_or_insulation_detail": 22, "acoustic_fireproof_detail": 23,
    "note": 24,
}


# Purpose: Normalize one header cell for tolerant label matching.
# Example output: "total of quantities (k)"
def _header_text(value: Any) -> str:
    """Normalize one header cell for tolerant label matching."""
    # Check this condition before choosing the next branch.
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        # Return the value produced by this step.
        return ""
    # Compute and store: text.
    text = unicodedata.normalize("NFD", str(value)).lower()
    # Compute and store: text.
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    # Return the value produced by this step.
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


# Purpose: Detect source fields from multi-row Excel headers.
# Example output: {"description": 2, "W": 5, "H": 6, "K": 16}
def detect_source_columns(data: pd.DataFrame) -> dict[str, int]:
    """Detect source fields from multi-row Excel headers.

    The old workbook layout remains the fallback.  When a contractor removes
    optional columns such as the duplicate Unit column, W/H/etc. are found by
    their own header labels and therefore do not shift into the wrong fields.
    """
    # Compute and store: columns.
    columns = dict(DEFAULT_SOURCE_COLUMNS)
    # Compute and store: scan_rows.
    scan_rows = min(len(data), 40)
    # Compute and store: header_row.
    header_row = next(
        (
            row
            for row in range(scan_rows)
            if "total of quantities" in " ".join(_header_text(value) for value in data.iloc[row].tolist())
        ),
        None,
    )
    # Compute and store: rows_to_scan.
    rows_to_scan = (
        range(header_row, min(header_row + 2, scan_rows))
        if header_row is not None
        else range(scan_rows)
    )
    # Compute and store: header_values.
    header_values = {
        column: " ".join(_header_text(data.iat[row, column]) for row in rows_to_scan).strip()
        for column in range(data.shape[1])
    }

    # Purpose: Assign the first column whose normalized header matches a rule.
    def set_first(name: str, predicate) -> None:
        """Assign the first column whose normalized header matches a rule."""
        # Process each item in this input sequence.
        for column, text in header_values.items():
            # Check this condition before choosing the next branch.
            if predicate(text):
                # Compute and store: columns[name].
                columns[name] = column
                # Return the value produced by this step.
                return

    # Execute this expression for its result or side effect.
    set_first("item_no", lambda text: "no stt" in text or text in ("no", "stt") or " stt " in f" {text} ")
    # Execute this expression for its result or side effect.
    set_first("description", lambda text: "name of material" in text or "ten vat tu" in text or text == "mo ta")
    # Execute this expression for its result or side effect.
    set_first("unit", lambda text: "unit don vi tinh" in text or text == "unit")
    # Execute this expression for its result or side effect.
    set_first("Q", lambda text: "so moi ghep" in text or "joint quantity" in text or text == "q")
    # Execute this expression for its result or side effect.
    set_first("K", lambda text: "total of quantities" in text or "so luong tong" in text)
    # Execute this expression for its result or side effect.
    set_first("plan", lambda text: "plan mat bang" in text)
    # Execute this expression for its result or side effect.
    set_first("up_down", lambda text: "up down len xuong" in text)
    # Execute this expression for its result or side effect.
    set_first("vertical_axis", lambda text: "vertical exis" in text or "truc dung" in text)
    # Execute this expression for its result or side effect.
    set_first("formula_or_insulation_detail", lambda text: "dien giai chi tiet" in text and "cach nhiet" in text and "trong" not in text)
    # Execute this expression for its result or side effect.
    set_first("acoustic_fireproof_detail", lambda text: ("dien giai" in text and "trong" in text) or "dien giai" in text and "chong chay" in text)
    # Execute this expression for its result or side effect.
    set_first("note", lambda text: "note ghi chu" in text or "chu thich" in text or text == "note")

    # Dimension labels are usually on the second header row.  Exact matching
    # avoids confusing W' and W'' with W.
    dimension_names = {"W": "w", "H": "h", "W1": "w", "H1": "h", "W2": "w", "H2": "h"}
    # Process each item in this input sequence.
    for name, label in dimension_names.items():
        # Compute and store: ordinal.
        ordinal = {"W": 0, "H": 0, "W1": 1, "H1": 1, "W2": 2, "H2": 2}[name]
        # Compute and store: candidates.
        candidates = []
        # Process each item in this input sequence.
        for column, text in header_values.items():
            # Check this condition before choosing the next branch.
            if header_row is not None:
                # Compute and store: header_rows.
                header_rows = range(header_row + 1, min(header_row + 2, scan_rows))
                # Check this condition before choosing the next branch.
                if header_row + 1 >= scan_rows:
                    # Compute and store: header_rows.
                    header_rows = range(header_row, header_row + 1)
            else:
                # Compute and store: header_rows.
                header_rows = range(min(scan_rows, 18))
            # Compute and store: header_cells.
            header_cells = [
                str(data.iat[row, column]).replace("\n", " ").strip()
                for row in header_rows
                if pd.notna(data.iat[row, column])
            ]
            # Check this condition before choosing the next branch.
            if any(re.fullmatch(rf"{label}\s*['\"]*", value, flags=re.IGNORECASE) for value in header_cells):
                # Execute this expression for its result or side effect.
                candidates.append(column)
        # Check this condition before choosing the next branch.
        if len(candidates) > ordinal:
            # Compute and store: columns[name].
            columns[name] = candidates[ordinal]

    # Execute this expression for its result or side effect.
    set_first("L_M", lambda text: re.search(r"(^| )l m($| )", text) is not None)
    # Execute this expression for its result or side effect.
    set_first("R_D", lambda text: re.search(r"(^| )r d($| )", text) is not None)
    # Execute this expression for its result or side effect.
    set_first("a", lambda text: text.strip() == "a")
    # Execute this expression for its result or side effect.
    set_first("t", lambda text: text.strip() == "t")
    # Execute this expression for its result or side effect.
    set_first("insulation_quantity", lambda text: "khoi luong cach nhiet m2" in text and "trong" not in text)
    # Execute this expression for its result or side effect.
    set_first("fireproof_quantity", lambda text: "khoi luong chong chay" in text)
    # Execute this expression for its result or side effect.
    set_first("acoustic_quantity", lambda text: "khoi luong cach nhiet trong" in text)

    # Some workbooks use the target labels directly instead of descriptive
    # headers. Keep SCN, SCC and STA separate when those columns exist.
    set_first(
        "scn_quantity",
        lambda text: text in {"scn", "s c n"}
        or (
            re.search(r"(^| )scn($| )", text) is not None
            and ("m2" in text or "quantity" in text)
        ),
    )
    # Execute this expression for its result or side effect.
    set_first(
        "scc_quantity",
        lambda text: (
            (
                text in {"scc", "s c c"}
                or (
                    re.search(r"(^| )scc($| )", text) is not None
                    and ("m2" in text or "quantity" in text)
                )
            )
            and "ta" not in text
        ),
    )
    # Execute this expression for its result or side effect.
    set_first(
        "sta_quantity",
        lambda text: text in {"sta", "s t a", "scc ta", "scc t a"}
        or (
            re.search(r"(^| )sta($| )", text) is not None
            and ("m2" in text or "quantity" in text)
        )
        or (
            "scc" in text
            and "ta" in text
            and ("m2" in text or "quantity" in text)
        ),
    )

    # The reference template uses descriptive headers. Expose the same
    # semantic targets without changing the legacy two-quantity layout.
    if "scn_quantity" not in columns and "insulation_quantity" in columns:
        # Compute and store: columns['scn_quantity'].
        columns["scn_quantity"] = columns["insulation_quantity"]
    # Check this condition before choosing the next branch.
    if "scc_quantity" not in columns and "fireproof_quantity" in columns:
        # Compute and store: columns['scc_quantity'].
        columns["scc_quantity"] = columns["fireproof_quantity"]
    # Check this condition before choosing the next branch.
    if "sta_quantity" not in columns and "acoustic_quantity" in columns:
        # Compute and store: columns['sta_quantity'].
        columns["sta_quantity"] = columns["acoustic_quantity"]
    # Return the value produced by this step.
    return columns


# Purpose: Return the source quantity field for S/SCN/SCC/STA/SCC-TA.
# Example output: "scn_quantity"
def quantity_column_for_target(columns: dict[str, int], target: str) -> str | None:
    """Return the source quantity field for S/SCN/SCC/STA/SCC-TA."""
    # Compute and store: target_key.
    target_key = target.upper().replace(" ", "")
    # Check this condition before choosing the next branch.
    if target_key in {"S", "SCN"}:
        # Return the value produced by this step.
        return "scn_quantity" if "scn_quantity" in columns else "insulation_quantity"
    # Check this condition before choosing the next branch.
    if target_key == "SCC":
        # Check this condition before choosing the next branch.
        if "scc_quantity" in columns:
            # Return the value produced by this step.
            return "scc_quantity"
        # Check this condition before choosing the next branch.
        if "fireproof_quantity" in columns:
            # Return the value produced by this step.
            return "fireproof_quantity"
        # Legacy files store SCC together with the inside quantity.
        return "acoustic_quantity"
    # Check this condition before choosing the next branch.
    if target_key in {"STA", "SCC/TA", "SCCTA"}:
        # Return the value produced by this step.
        return "sta_quantity" if "sta_quantity" in columns else "acoustic_quantity"
    # Return the value produced by this step.
    return None

PRODUCT_INPUT_FORMATS = {
    "rectangular_duct": ("W", "H", "K"),
    "round_duct": ("R_D", "K"),
    "end_cap": ("W", "H", "K"),
    "reducer": ("W", "H", "W1", "H1", "L_M", "K"),
    "square_round_reducer": ("W", "H", "L_M", "R_D", "K"),
    "tee": ("W", "H", "W1", "H1", "W2", "H2", "L_M", "K"),
    "takeoff": ("W", "H", "W1", "H1", "L_M", "K"),
    "elbow": ("W", "H", "L_M", "R_D", "K"),
    "square_elbow": ("W", "H", "R_D", "K"),
    "mfd_valve": ("W", "H", "L_M", "K"),
    "nrd_valve": ("W", "H", "L_M", "K"),
}

__all__ = [
    "FORMULA_INPUT_FIELDS",
    "SHEET_02_COLUMN_MAP",
    "PRODUCT_INPUT_FORMATS",
    "DEFAULT_SOURCE_COLUMNS",
    "detect_source_columns",
    "quantity_column_for_target",
]
