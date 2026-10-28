"""Format one contractor row before sending it to the formula engine."""

from __future__ import annotations

from dataclasses import asdict, dataclass
import math
import re
from typing import Any, Mapping
import unicodedata

import pandas as pd

from .formula_engine import CalculationResult, FormulaInputs, calculate


# Purpose: Normalize text for case-insensitive product and section matching.
# Example output: "Square elbow 360x360"
def _text(value: Any) -> str:
    """Normalize text for case-insensitive product and section matching."""
    # Check this condition before choosing the next branch.
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        # Return the value produced by this step.
        return ""
    # Compute and store: text.
    text = str(value).strip().lower()
    # Return the value produced by this step.
    return "".join(
        character
        for character in unicodedata.normalize("NFD", text)
        if unicodedata.category(character) != "Mn"
    )


# Purpose: Read one numeric input, including common contractor aliases.
# Example output: 1.0
def _number(row: Mapping[str, Any], name: str) -> float:
    """Read one numeric input, including common contractor aliases."""
    # Compute and store: aliases.
    aliases = {
        "W1": ("W_prime",),
        "H1": ("H_prime",),
        "W2": ("W_double_prime",),
        "H2": ("H_double_prime",),
        "L_M": ("L",),
        "R_D": ("R", "D"),
    }
    # Compute and store: value.
    value = row.get(name, None)
    # Check this condition before choosing the next branch.
    if value is None:
        # Compute and store: value.
        value = next((row.get(alias) for alias in aliases.get(name, ()) if alias in row), 0.0)
    # Check this condition before choosing the next branch.
    if value is None or (not isinstance(value, str) and pd.isna(value)) or value == "":
        # Return the value produced by this step.
        return 0.0
    # Compute and store: number.
    number = float(value)
    # Check this condition before choosing the next branch.
    if not math.isfinite(number):
        # Stop with a clear validation error for invalid input.
        raise ValueError(f"Input {name} must be finite.")
    # Return the value produced by this step.
    return number


# Purpose: Normalize an explicit section name to an engine section.
# Example output: "fireproof"
def _section(row: Mapping[str, Any]) -> str:
    """Normalize an explicit section name to an engine section."""
    # Compute and store: value.
    value = _text(row.get("section"))
    # Compute and store: aliases.
    aliases = {
        "standard": "standard",
        "fireproof": "fireproof",
        "fire proof": "fireproof",
        "acoustic": "acoustic",
        "tieu am": "acoustic",
    }
    # Check this condition before choosing the next branch.
    if value in aliases:
        # Return the value produced by this step.
        return aliases[value]

    # Return the value produced by this step.
    return "standard"


# Purpose: Calculate the section column from the contractor formula labels.
# Example output: "acoustic"
def calculate_section(row: Mapping[str, Any]) -> str:
    """Calculate the section column from the contractor formula labels."""
    # Compute and store: explicit.
    explicit = _text(row.get("section"))
    # Check this condition before choosing the next branch.
    if explicit:
        # Return the value produced by this step.
        return _section(row)

    # Compute and store: formula.
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
        # Return the value produced by this step.
        return "acoustic"
    # Check this condition before choosing the next branch.
    if "scn" in formula or "scc" in formula:
        # Return the value produced by this step.
        return "fireproof"
    # Return the value produced by this step.
    return "standard"


# Purpose: Classify a product description into a formula product type.
# Example output: "square_elbow"
def _product_type(description: Any) -> str:
    """Classify a product description into a formula product type."""
    # Compute and store: text.
    text = _text(description)
    # Compute and store: rules.
    rules = (
        ("square_round_reducer", ("square round reducer", "giam vuong tron")),
        ("square_elbow", ("square elbow", "co vuong")),
        ("technical_louver_box", ("plenum box", "technical room", "box gio phong ky thuat", "louver")),
        ("fcu_reducer_box", ("box cap giam fcu",)),
        ("return_fcu_box", ("box hoi fcu",)),
        ("fcu_box", ("box cap fcu",)),
        ("air_outlet_box", ("box cap mieng gio",)),
        ("end_cap", ("end cap", "bit dau", "bit đau", "bi dau")),
        ("stiffener", ("stiffener", "boc tang cung")),
        ("reducer", ("reducer", "giam cap", "giam")),
        ("tee", ("tee", "three-way", "chac 3")),
        ("takeoff", ("take off", "takeoff", "chan re")),
        # Check valves before elbows.  A bare substring such as ``co`` must
        # never classify ``control damper`` as a square/round elbow.
        ("mfd_valve", ("fire damper", "mfd", "msd", "van ngan chay")),
        ("nrd_valve", ("vcd", "nrd", "van chinh gio", "volume control damper")),
        ("elbow", ("elbow", "co 90", "co 45", "co tron")),
        ("round_duct", ("flexible air duct", "ong gio mem", "ong gio thang tron")),
        ("rectangular_duct", ("air duct", "ong gio")),
    )
    # Process each item in this input sequence.
    for product_type, keywords in rules:
        # Check this condition before choosing the next branch.
        if any(keyword in text for keyword in keywords):
            # Return the value produced by this step.
            return product_type

    # The reference workbook has a short name ``Co``. Use a word boundary so
    # it remains distinct from ``control`` in Volume Control Damper.
    if re.search(r"\bco\b", text):
        # Return the value produced by this step.
        return "elbow"

    # Check this condition before choosing the next branch.
    if re.search(r"\b(fd)\b", text):
        # Return the value produced by this step.
        return "mfd_valve"
    # Stop with a clear validation error for invalid input.
    raise ValueError(f"Cannot determine product type: {description!r}")


@dataclass(frozen=True)
class ProductFormat:
    """Normalized section/product pair used by the formula engine."""

    section: str
    product_type: str
    description: str = ""

    @classmethod
    # Purpose: Build the normalized product descriptor for one worksheet row.
    # Example output: ProductFormat(section="standard", product_type="rectangular_duct", ...)
    def from_row(
        cls,
        row: Mapping[str, Any],
        section: str | None = None,
        product_type: str | None = None,
    ) -> "ProductFormat":
        """Build the normalized product descriptor for one worksheet row."""
        # Compute and store: description.
        description = str(row.get("description", ""))
        # Return the value produced by this step.
        return cls(
            section=section or _section(row),
            product_type=product_type or _product_type(description),
            description=description,
        )

    # Purpose: Extract the formula engine inputs from a normalized row.
    # Example output: FormulaInputs(w=800, h=400, K=5.325, ...)
    def inputs_from(self, row: Mapping[str, Any]) -> FormulaInputs:
        """Extract the formula engine inputs from a normalized row."""
        # Return the value produced by this step.
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
            K=_number(
                row,
                "K_for_formula"
                if "K_for_formula" in row
                else ("K_By_function" if "K_By_function" in row else "K"),
            ),
        )

    # Purpose: Run the legacy formula engine for this product row.
    # Example output: CalculationResult(section="standard", product_type="rectangular_duct", R=15.12, S=0.0, T=0.0)
    def calculate(self, row: Mapping[str, Any]) -> CalculationResult:
        """Run the legacy formula engine for this product row."""
        # Compute and store: inputs.
        inputs = self.inputs_from(row)
        # Return the value produced by this step.
        return calculate(self.section, self.product_type, asdict(inputs))


__all__ = ["ProductFormat", "calculate_section"]
