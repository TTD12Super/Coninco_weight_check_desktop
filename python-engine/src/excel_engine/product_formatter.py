"""Format one contractor row before sending it to the formula engine."""

from __future__ import annotations

from dataclasses import asdict, dataclass
import math
import re
from typing import Any, Mapping
import unicodedata

import pandas as pd

from .formula_engine import CalculationResult, FormulaInputs, calculate


def _text(value: Any) -> str:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    text = str(value).strip().lower()
    return "".join(
        character
        for character in unicodedata.normalize("NFD", text)
        if unicodedata.category(character) != "Mn"
    )


def _number(row: Mapping[str, Any], name: str) -> float:
    aliases = {
        "W1": ("W_prime",),
        "H1": ("H_prime",),
        "W2": ("W_double_prime",),
        "H2": ("H_double_prime",),
        "L_M": ("L",),
        "R_D": ("R", "D"),
    }
    value = row.get(name, None)
    if value is None:
        value = next((row.get(alias) for alias in aliases.get(name, ()) if alias in row), 0.0)
    if value is None or (not isinstance(value, str) and pd.isna(value)) or value == "":
        return 0.0
    number = float(value)
    if not math.isfinite(number):
        raise ValueError(f"Input {name} must be finite.")
    return number


def _section(row: Mapping[str, Any]) -> str:
    value = _text(row.get("section"))
    aliases = {
        "standard": "standard",
        "fireproof": "fireproof",
        "fire proof": "fireproof",
        "acoustic": "acoustic",
        "tieu am": "acoustic",
    }
    if value in aliases:
        return aliases[value]

    return "standard"


def calculate_section(row: Mapping[str, Any]) -> str:
    """Calculate the section column from the contractor formula labels."""
    explicit = _text(row.get("section"))
    if explicit:
        return _section(row)

    formula = _text(
        " ".join(
            str(row.get(name, ""))
            for name in (
                "formula_or_insulation_detail",
                "insulation_quantity_detail",
                "acoustic_fireproof_detail",
            )
        )
    )
    # Scc/ta and Sta are acoustic; Scn and standalone Scc are fireproof.
    if "scc/ta" in formula or "sta" in formula:
        return "acoustic"
    if "scn" in formula or "scc" in formula:
        return "fireproof"
    return "standard"


def _product_type(description: Any) -> str:
    text = _text(description)
    rules = (
        ("square_round_reducer", ("square round reducer", "giam vuong tron")),
        ("square_elbow", ("square elbow", "co vuong")),
        ("technical_louver_box", ("plenum box", "technical room", "louver")),
        ("end_cap", ("end cap", "bit dau")),
        ("reducer", ("reducer", "giam cap")),
        ("tee", ("tee", "three-way", "chac 3")),
        ("takeoff", ("take off", "takeoff", "chan re")),
        ("elbow", ("elbow", "co 90", "co 45", "co tron")),
        ("nrd_valve", ("vcd", "nrd", "van chinh gio")),
        ("mfd_valve", ("fire damper", "mfd", "msd", "van ngan chay")),
        ("round_duct", ("flexible air duct", "ong gio mem")),
        ("rectangular_duct", ("air duct", "ong gio")),
    )
    for product_type, keywords in rules:
        if any(keyword in text for keyword in keywords):
            return product_type

    if re.search(r"\b(fd)\b", text):
        return "mfd_valve"
    raise ValueError(f"Cannot determine product type: {description!r}")


@dataclass(frozen=True)
class ProductFormat:
    """Normalized section/product pair used by the formula engine."""

    section: str
    product_type: str
    description: str = ""

    @classmethod
    def from_row(
        cls,
        row: Mapping[str, Any],
        section: str | None = None,
        product_type: str | None = None,
    ) -> "ProductFormat":
        description = str(row.get("description", ""))
        return cls(
            section=section or _section(row),
            product_type=product_type or _product_type(description),
            description=description,
        )

    def inputs_from(self, row: Mapping[str, Any]) -> FormulaInputs:
        return FormulaInputs(
            w=_number(row, "W"),
            h=_number(row, "H"),
            w1=_number(row, "W1"),
            h1=_number(row, "H1"),
            w2=_number(row, "W2"),
            h2=_number(row, "H2"),
            L_M=_number(row, "L_M"),
            R_D=_number(row, "R_D"),
            a=_number(row, "a"),
            t=_number(row, "t"),
            Q=_number(row, "Q"),
            K=_number(row, "K_By_function" if "K_By_function" in row else "K"),
        )

    def calculate(self, row: Mapping[str, Any]) -> CalculationResult:
        inputs = self.inputs_from(row)
        return calculate(self.section, self.product_type, asdict(inputs))


__all__ = ["ProductFormat", "calculate_section"]
