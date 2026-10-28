"""Python formula engine for the insulation workbook.

The Excel classifier is intentionally not here yet. The caller supplies the
section and product type, then this module returns R, S and T in square metres.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any, Callable, Mapping


########### Input types
@dataclass(frozen=True)
class FormulaInputs:
    """Normalized values from workbook columns C:N."""

    w: float = 0.0
    h: float = 0.0
    w1: float = 0.0
    h1: float = 0.0
    w2: float = 0.0
    h2: float = 0.0
    L_M: float = 0.0
    R_D: float = 0.0
    a: float = 0.0
    t: float = 0.0
    Q: float = 0.0
    K: float = 0.0

    # Readable aliases used by some of the longer formulas below.
    @property
    # Purpose: Return ``L_M`` using a readable name inside long formulas.
    # Example output: 1.0
    def length(self) -> float:
        """Return ``L_M`` using a readable name inside long formulas."""
        # Return the value produced by this step.
        return self.L_M

    @property
    # Purpose: Return ``R_D`` using a readable name inside long formulas.
    # Example output: 1.0
    def round_size(self) -> float:
        """Return ``R_D`` using a readable name inside long formulas."""
        # Return the value produced by this step.
        return self.R_D

    @property
    # Purpose: Return insulation thickness ``a``.
    # Example output: 1.0
    def insulation(self) -> float:
        """Return insulation thickness ``a``."""
        # Return the value produced by this step.
        return self.a

    @property
    # Purpose: Return fireproof thickness ``t``.
    # Example output: 1.0
    def fireproof(self) -> float:
        """Return fireproof thickness ``t``."""
        # Return the value produced by this step.
        return self.t

    @property
    # Purpose: Return joint quantity ``Q``.
    # Example output: 1.0
    def joints(self) -> float:
        """Return joint quantity ``Q``."""
        # Return the value produced by this step.
        return self.Q

    @property
    # Purpose: Return quantity multiplier ``K``.
    # Example output: 1.0
    def quantity(self) -> float:
        """Return quantity multiplier ``K``."""
        # Return the value produced by this step.
        return self.K


###########. Output 
@dataclass(frozen=True)
class CalculationResult:
    """The three result columns from the workbook."""

    section: str
    product_type: str
    R: float
    S: float
    T: float
    unit: str = "m2"

    @property
    # Purpose: Return result column R.
    # Example output: 1.0
    def r(self) -> float:
        """Return result column R."""
        # Return the value produced by this step.
        return self.R

    @property
    # Purpose: Return result column S.
    # Example output: 1.0
    def s(self) -> float:
        """Return result column S."""
        # Return the value produced by this step.
        return self.S

    @property
    # Purpose: Return result column T.
    # Example output: 1.0
    def t(self) -> float:
        """Return result column T."""
        # Return the value produced by this step.
        return self.T

    # Purpose: Serialize the result for the JSON boundary used by the desktop app.
    # Example output: {"section": "standard", "quantities": {"R": 15.12, "S": 0.0, "T": 0.0}}
    def to_dict(self) -> dict[str, Any]:
        """Serialize the result for the JSON boundary used by the desktop app."""
        # Return the value produced by this step.
        return {
            "section": self.section,
            "product_type": self.product_type,
            "unit": self.unit,
            "quantities": {"R": self.R, "S": self.S, "T": self.T},
        }


# Friendly names are accepted in addition to the original workbook names.
_ALIASES = {
    "W": ("width",),
    "H": ("height",),
    "W1": ("width_1",),
    "H1": ("height_1",),
    "W2": ("width_2",),
    "H2": ("height_2",),
    "L_M": ("L", "l", "length"),
    "R_D": ("D", "d", "R", "r", "diameter", "radius"),
    "a": ("insulation_thickness",),
    "t": ("fireproof_thickness",),
    "Q": ("joint_quantity",),
    "K": ("quantity",),
}


# Purpose: Read one numeric field and fail early when an input is invalid.
# Example output: 1.0
def _number(values: Mapping[str, Any], key: str) -> float:
    """Read one numeric field and fail early when an input is invalid."""
    # Compute and store: names.
    names = (key, key.lower(), *_ALIASES.get(key, ()))
    # Compute and store: value.
    value = next((values[name] for name in names if name and name in values), 0.0)

    # Check this condition before choosing the next branch.
    if value is None or value == "":
        # Return the value produced by this step.
        return 0.0
    # Protect parsing/evaluation so bad workbook data becomes a controlled error.
    try:
        # Compute and store: number.
        number = float(value)
    except (TypeError, ValueError) as error:
        # Stop with a clear validation error for invalid input.
        raise ValueError(f"Input {key} must be numeric, got {value!r}.") from error
    # Check this condition before choosing the next branch.
    if not math.isfinite(number):
        # Stop with a clear validation error for invalid input.
        raise ValueError(f"Input {key} must be finite, got {value!r}.")
    # Return the value produced by this step.
    return number


# Purpose: Convert JSON/workbook-style names into one simple Python object.
# Example output: FormulaInputs(w=800, h=400, K=5.325, ...)
def _read_inputs(values: Mapping[str, Any]) -> FormulaInputs:
    """Convert JSON/workbook-style names into one simple Python object."""
    # Check this condition before choosing the next branch.
    if not isinstance(values, Mapping):
        # Stop with a clear validation error for invalid input.
        raise ValueError("inputs must be an object/dictionary.")
    # Check this condition before choosing the next branch.
    if not any(name in values for name in ("K", "k", "quantity")):
        # Stop with a clear validation error for invalid input.
        raise ValueError("Input K (quantity) is required.")
    # Compute and store: numbers.
    numbers = [_number(values, key) for key in ("W", "H", "W1", "H1", "W2", "H2", "L_M", "R_D", "a", "t", "Q", "K")]
    # Return the value produced by this step.
    return FormulaInputs(*numbers)


# Purpose: Excel SQRT equivalent with a clear Python error for bad dimensions.
# Example output: 1.0
def _sqrt(value: float) -> float:
    """Excel SQRT equivalent with a clear Python error for bad dimensions."""
    # Check this condition before choosing the next branch.
    if value < 0:
        # Stop with a clear validation error for invalid input.
        raise ValueError(f"Square-root input cannot be negative: {value}.")
    # Return the value produced by this step.
    return math.sqrt(value)


# ---------------------------------------------------------------------------
# Standard / normal insulation formulas (Excel rows 6:18)
# ---------------------------------------------------------------------------


########## Ống gió
########## Ống gió chữ nhật
# Purpose: Standard rectangular duct: Excel row 6.
# Example output: (12.345, 0.0, 0.0)
def _standard_rectangular(x: FormulaInputs) -> tuple[float, float, float]:
    """Standard rectangular duct: Excel row 6."""
    # Compute and store: r.
    r = ((x.w + x.h) * 2 * x.K * 1000 + (x.w + x.h + 25 * 4) * 2 * x.Q * 100) / 1_000_000
    # Return the value produced by this step.
    return r, 0.0, 0.0

######### Ống gió thẳng tròn
########## Ống gió tròn
# Purpose: Standard round duct: Excel row 7.
# Example output: (12.345, 0.0, 0.0)
def _standard_round(x: FormulaInputs) -> tuple[float, float, float]:
    """Standard round duct: Excel row 7."""
    # Compute and store: r.
    r = (math.pi * x.R_D * x.K * 1000 + math.pi * (x.R_D + 25 * 2) * x.Q * 100) / 1_000_000
    # Return the value produced by this step.
    return r, 0.0, 0.0


########## Bít đầu
# Purpose: Standard end cap: Excel row 9.
# Example output: (12.345, 0.0, 0.0)
def _standard_end_cap(x: FormulaInputs) -> tuple[float, float, float]:
    """Standard end cap: Excel row 9."""
    # Compute and store: r.
    r = (x.w + 50) * (x.h + 50) * x.K / 1_000_000
    # Return the value produced by this step.
    return r, 0.0, 0.0


########## Giảm chữ nhật
# Purpose: Standard reducer: Excel row 10, including its existing parentheses.
# Example output: (12.345, 0.0, 0.0)
def _standard_reducer(x: FormulaInputs) -> tuple[float, float, float]:
    """Standard reducer: Excel row 10, including its existing parentheses."""
    # Compute and store: side_1.
    side_1 = (x.w + x.w1) * _sqrt(x.length**2 + (((x.h - x.h1) / 2) ** 2))
    # Keep the original Excel parentheses: SQRT(I^2) + (C-E/2)^2.
    side_2 = (x.h + x.h1) * _sqrt(x.length**2) + (x.w - x.w1 / 2) ** 2
    # Compute and store: seam.
    seam = (x.w + x.h + 25 * 4) * 2 * x.Q * 100 * x.K
    # Compute and store: r.
    r = (x.K * (side_1 + side_2) + seam) / 1_000_000
    # Return the value produced by this step.
    return r, 0.0, 0.0


########## Giảm vuông tròn
# Purpose: Standard square-to-round reducer: Excel row 11.
# Example output: (12.345, 0.0, 0.0)
def _standard_square_round_reducer(x: FormulaInputs) -> tuple[float, float, float]:
    """Standard square-to-round reducer: Excel row 11."""
    # Compute and store: d.
    d = x.round_size
    # Compute and store: side_1.
    side_1 = (x.w + math.pi * d / 4) * _sqrt(x.length**2 + (x.h - math.pi * d / 4) ** 2)
    # Compute and store: side_2.
    side_2 = (x.h + math.pi * d / 4) * _sqrt(x.length**2 + (x.w - math.pi * d / 4) ** 2)
    # Compute and store: seam.
    seam = (x.w + x.h + 25 * 4) * 2 * 100 * x.Q * x.K
    # Compute and store: round_seam.
    round_seam = math.pi * (d + 25 * 2) * x.K * x.Q * 100
    # Return the value produced by this step.
    return (x.K * (side_1 + side_2) + seam + round_seam) / 1_000_000, 0.0, 0.0


########## Chạc 3
# Purpose: Standard three-way branch: Excel row 12.
# Example output: (12.345, 0.0, 0.0)
def _standard_tee(x: FormulaInputs) -> tuple[float, float, float]:
    """Standard three-way branch: Excel row 12."""
    # Compute and store: branch_1.
    branch_1 = 2 * (x.w1 + x.h1) * (x.length - 1.5 * x.w2)
    # Compute and store: branch_2.
    branch_2 = (2 * math.pi * ((x.w + x.w / 2) ** 2 - (x.w / 2) ** 2) + 2 * math.pi * x.h * ((x.w + x.w / 2) + x.w / 2)) * 90 / 360
    # Compute and store: branch_3.
    branch_3 = (2 * math.pi * ((x.w2 + x.w2 / 2) ** 2 - x.w2 / 2**2) + 2 * math.pi * x.h2 * ((x.w2 + x.w2 / 2) + x.w2 / 2)) * 90 / 360
    # Compute and store: deduction.
    deduction = x.h * x.w / 2 + x.h2 * x.w2 / 2
    # Compute and store: seam.
    seam = (x.w1 + x.h1 + 25 * 4) * 2 * 2 * 100 * x.K * x.Q
    # Return the value produced by this step.
    return (x.K * (branch_1 + branch_2 + branch_3 - deduction) + seam) / 1_000_000, 0.0, 0.0


########## Chân rẽ
# Purpose: Standard takeoff branch: Excel row 13.
# Example output: (12.345, 0.0, 0.0)
def _standard_takeoff(x: FormulaInputs) -> tuple[float, float, float]:
    """Standard takeoff branch: Excel row 13."""
    # Compute and store: surface.
    surface = (x.w + x.w1) * x.length + x.h1 * x.length + x.h * _sqrt((x.w - x.w1) ** 2 + x.length**2) - x.w * x.h
    # Compute and store: seam.
    seam = (x.w1 + x.h1 + 25 * 4) * 2 * 100 * x.K * x.Q
    # Return the value produced by this step.
    return x.K * (surface + seam) / 1_000_000, 0.0, 0.0


########## Co tròn
# Purpose: Standard round elbow: Excel row 14.
# Example output: (12.345, 0.0, 0.0)
def _standard_elbow(x: FormulaInputs) -> tuple[float, float, float]:
    """Standard round elbow: Excel row 14."""
    # Compute and store: d.
    d = x.round_size
    # Compute and store: surface.
    surface = (2 * math.pi * ((x.w + d) ** 2 - d**2) + 2 * math.pi * x.h * (x.w + d + d)) * x.length / 360
    # Compute and store: seam.
    seam = (x.w + x.h + 25 * 4) * 2 * 100 * x.K * x.Q
    # Return the value produced by this step.
    return (x.K * surface + seam) / 1_000_000, 0.0, 0.0


########## Co vuông
# Purpose: Standard square elbow: Excel row 15.
# Example output: (12.345, 0.0, 0.0)
def _standard_square_elbow(x: FormulaInputs) -> tuple[float, float, float]:
    """Standard square elbow: Excel row 15."""
    # Compute and store: d.
    d = x.round_size
    # Compute and store: surface.
    surface = 2 * ((x.w + d) ** 2 - d**2) + 2 * (x.w + d) * x.h + 2 * d * x.h
    # Compute and store: seam.
    seam = (x.w + x.h + 25 * 4) * 2 * 100 * x.K * x.Q
    # Return the value produced by this step.
    return (x.K * surface + seam) / 1_000_000, 0.0, 0.0


########## Van MFD / FD / MSD
# Purpose: Standard MFD/FD/MSD valve: Excel row 16.
# Example output: (12.345, 0.0, 0.0)
def _standard_mfd(x: FormulaInputs) -> tuple[float, float, float]:
    """Standard MFD/FD/MSD valve: Excel row 16."""
    # Compute and store: surface.
    surface = 2 * ((x.w + 50) + (x.h + 50)) * x.length
    # Compute and store: seam.
    seam = (x.w + 50 + x.h + 50 + 25 * 4) * 2 * 100 * x.K * x.Q
    # Return the value produced by this step.
    return (x.K * surface + seam) / 1_000_000, 0.0, 0.0


########## Van NRD / VCD / MD
# Purpose: Standard NRD/VCD/MD valve: Excel row 17.
# Example output: (12.345, 0.0, 0.0)
def _standard_nrd(x: FormulaInputs) -> tuple[float, float, float]:
    """Standard NRD/VCD/MD valve: Excel row 17."""
    # Compute and store: surface.
    surface = 2 * (x.w + x.h) * x.length
    # Compute and store: seam.
    seam = (x.w + x.h + 25 * 4) * 2 * 100 * x.K * x.Q
    # Return the value produced by this step.
    return (x.K * surface + seam) / 1_000_000, 0.0, 0.0


########## Bọc tăng cứng
# Purpose: Standard stiffener wrap: Excel row 18.
# Example output: (12.345, 0.0, 0.0)
def _standard_stiffener(x: FormulaInputs) -> tuple[float, float, float]:
    """Standard stiffener wrap: Excel row 18."""
    # Return the value produced by this step.
    return (x.w + x.h + 25 * 4) * 2 * 100 * x.Q / 1_000_000, 0.0, 0.0


_STANDARD = {
    "rectangular_duct": _standard_rectangular,
    "round_duct": _standard_round,
    "end_cap": _standard_end_cap,
    "reducer": _standard_reducer,
    "square_round_reducer": _standard_square_round_reducer,
    "tee": _standard_tee,
    "takeoff": _standard_takeoff,
    "elbow": _standard_elbow,
    "square_elbow": _standard_square_elbow,
    "mfd_valve": _standard_mfd,
    "nrd_valve": _standard_nrd,
    "stiffener": _standard_stiffener,
}


# ---------------------------------------------------------------------------
# Fireproof formulas (Excel rows 21:31)
# ---------------------------------------------------------------------------


########## Ống gió chữ nhật - chống cháy
# Purpose: Fireproof rectangular duct: Excel row 21.
# Example output: (12.345, 0.0, 0.0)
def _fireproof_rectangular(x: FormulaInputs) -> tuple[float, float, float]:
    """Fireproof rectangular duct: Excel row 21."""
    # Compute and store: seams.
    seams = (x.w + x.h + 4 * x.insulation + 4 * x.fireproof) * 2 * x.Q * 100 + (x.w + x.h + 8 * x.insulation + 4 * x.fireproof) * 2 * x.Q * 100
    # Compute and store: r.
    r = ((x.w + x.h) * 2 * x.K * 1000 + seams) / 1_000_000 * x.insulation / 25
    # Compute and store: s.
    s = (x.w + x.h + 4 * x.insulation) * 2 * x.K * 1000 / 1_000_000
    # Return the value produced by this step.
    return r, s, 0.0


########## Bít đầu - chống cháy
# Purpose: Fireproof end cap: Excel row 23.
# Example output: (12.345, 0.0, 0.0)
def _fireproof_end_cap(x: FormulaInputs) -> tuple[float, float, float]:
    """Fireproof end cap: Excel row 23."""
    # Compute and store: r.
    r = (x.w + 2 * x.insulation) * (x.h + 2 * x.insulation) * x.K / 1_000_000 * x.insulation / 25
    # Compute and store: s.
    s = (x.w + 2 * x.insulation + 2 * x.fireproof) * (x.h + 2 * x.insulation + 2 * x.fireproof) * x.K / 1_000_000
    # Return the value produced by this step.
    return r, s, 0.0


########## Giảm chữ nhật - chống cháy
# Purpose: Fireproof reducer: Excel row 24.
# Example output: (12.345, 0.0, 0.0)
def _fireproof_reducer(x: FormulaInputs) -> tuple[float, float, float]:
    """Fireproof reducer: Excel row 24."""
    # Compute and store: side_1.
    side_1 = (x.w + x.w1) * _sqrt(x.length**2 + ((x.h - x.h1) / 2) ** 2)
    # Compute and store: side_2.
    side_2 = (x.h + x.h1) * _sqrt(x.length**2 + ((x.w - x.w1) / 2) ** 2)
    # Compute and store: seams.
    seams = ((x.w + x.h + 4 * x.insulation + 4 * x.fireproof) * 2 * x.Q * 100 * x.K + (x.w + x.h + 8 * x.insulation + 4 * x.fireproof) * 2 * x.Q * 100 * x.K)
    # Compute and store: r.
    r = (x.K * (side_1 + side_2) + seams) / 1_000_000 * x.insulation / 25
    # Compute and store: s.
    s = x.K * ((x.w + x.w1 + 4 * x.insulation) * _sqrt(x.length**2 + ((x.h - x.h1) / 2) ** 2) + (x.h + x.h1 + 4 * x.insulation) * _sqrt(x.length**2 + (x.w - x.w1 / 2) ** 2)) / 1_000_000
    # Return the value produced by this step.
    return r, s, 0.0


########## Giảm vuông tròn - chống cháy
# Purpose: Fireproof square-to-round reducer: Excel row 25.
# Example output: (12.345, 0.0, 0.0)
def _fireproof_square_round_reducer(x: FormulaInputs) -> tuple[float, float, float]:
    """Fireproof square-to-round reducer: Excel row 25."""
    # Compute and store: d.
    d = x.round_size
    # Compute and store: side_1.
    side_1 = (x.w + math.pi * d / 4) * _sqrt(x.length**2 + (x.h - math.pi * d / 4) ** 2)
    # Compute and store: side_2.
    side_2 = (x.h + math.pi * d / 4) * _sqrt(x.length**2 + (x.w - math.pi * d / 4) ** 2)
    # Compute and store: seams.
    seams = ((x.w + x.h + 4 * x.insulation + 4 * x.fireproof) * 2 * 100 * x.Q * x.K + (x.w + x.h + 8 * x.insulation + 4 * x.fireproof) * 2 * 100 * x.Q * x.K)
    # Compute and store: round_seam.
    round_seam = math.pi * (d + 4 * x.insulation + 4 * x.fireproof) * 2 * x.K * x.Q * 100
    # Compute and store: r.
    r = (x.K * (side_1 + side_2) + seams + round_seam) / 1_000_000 * x.insulation / 25
    # Compute and store: outer_d.
    outer_d = d + 2 * x.insulation
    # Compute and store: s.
    s = x.K * ((x.w + 2 * x.insulation + math.pi * outer_d / 4) * _sqrt(x.length**2 + (x.h + 2 * x.insulation - math.pi * outer_d / 4) ** 2) + (x.h + 2 * x.insulation + math.pi * outer_d / 4) * _sqrt(x.length**2 + (x.w + 2 * x.insulation - math.pi * outer_d / 4) ** 2)) / 1_000_000
    # Return the value produced by this step.
    return r, s, 0.0


########## Chạc 3 - chống cháy
# Purpose: Fireproof three-way branch: Excel row 26.
# Example output: (12.345, 0.0, 0.0)
def _fireproof_tee(x: FormulaInputs) -> tuple[float, float, float]:
    """Fireproof three-way branch: Excel row 26."""
    # Compute and store: branch_1.
    branch_1 = 2 * (x.w1 + x.h1) * (x.length - 1.5 * x.w2)
    # Compute and store: branch_2.
    branch_2 = (2 * math.pi * ((x.w + x.w / 2) ** 2 - (x.w / 2) ** 2) + 2 * math.pi * x.h * ((x.w + x.w / 2) + x.w / 2)) * 90 / 360
    # Compute and store: branch_3.
    branch_3 = (2 * math.pi * ((x.w2 + x.w2 / 2) ** 2 - x.w2 / 2**2) + 2 * math.pi * x.h2 * ((x.w2 + x.w2 / 2) + x.w2 / 2)) * 90 / 360
    # Compute and store: deduction.
    deduction = x.h * x.w / 2 + x.h2 * x.w2 / 2
    # Compute and store: seams.
    seams = ((x.w1 + x.h1 + 4 * x.insulation + 4 * x.fireproof) * 2 * 2 * 100 * x.K * x.Q + (x.w1 + x.h1 + 8 * x.insulation + 4 * x.fireproof) * 2 * 2 * 100 * x.K * x.Q)
    # Compute and store: r.
    r = (x.K * (branch_1 + branch_2 + branch_3 - deduction) + seams) / 1_000_000 * x.insulation / 25

    # Compute and store: (ow, oh).
    ow, oh = x.w + 2 * x.insulation, x.h + 2 * x.insulation
    # Compute and store: (ow2, oh2).
    ow2, oh2 = x.w2 + 2 * x.insulation, x.h2 + 2 * x.insulation
    # Compute and store: outer_2.
    outer_2 = (2 * math.pi * ((ow + ow / 2) ** 2 - (ow / 2) ** 2) + 2 * math.pi * oh * ((ow + ow / 2) + ow / 2)) * 90 / 360
    # Keep the source Excel precedence: (G+2K)/2^2, not ((G+2K)/2)^2.
    outer_3 = (2 * math.pi * ((ow2 + ow2 / 2) ** 2 - ow2 / 2**2) + 2 * math.pi * oh2 * ((ow2 + ow2 / 2) + ow2 / 2)) * 90 / 360
    # Compute and store: outer_deduction.
    outer_deduction = oh * ow / 2 + oh2 * ow2 / 2
    # Compute and store: s.
    s = x.K * (2 * (x.w1 + x.h1 + 4 * x.insulation) * (x.length + 2 * x.insulation - 1.5 * (x.w2 + 2 * x.insulation)) + outer_2 + outer_3 - outer_deduction) / 1_000_000
    # Return the value produced by this step.
    return r, s, 0.0


########## Chân rẽ - chống cháy
# Purpose: Fireproof takeoff branch: Excel row 27.
# Example output: (12.345, 0.0, 0.0)
def _fireproof_takeoff(x: FormulaInputs) -> tuple[float, float, float]:
    """Fireproof takeoff branch: Excel row 27."""
    # Compute and store: surface.
    surface = (x.w + x.w1) * x.length + x.h1 * x.length + x.h * _sqrt((x.w - x.w1) ** 2 + x.length**2) - x.w * x.h
    # Compute and store: seams.
    seams = ((x.w1 + x.h1 + 4 * x.insulation + 4 * x.fireproof) * 2 * 100 * x.K * x.Q + (x.w1 + x.h1 + 8 * x.insulation + 4 * x.fireproof) * 2 * 100 * x.K * x.Q)
    # Compute and store: r.
    r = x.K * (surface + seams) / 1_000_000 * x.insulation / 25
    # Compute and store: s.
    s = x.K * ((x.w + x.w1 + 4 * x.insulation) * x.length + (x.h1 + 2 * x.insulation) * x.length + (x.h + 2 * x.insulation) * _sqrt((x.w - x.w1) ** 2 + x.length**2) - x.w * x.h) / 1_000_000
    # Return the value produced by this step.
    return r, s, 0.0


########## Co tròn - chống cháy
# Purpose: Fireproof round elbow: Excel row 28.
# Example output: (12.345, 0.0, 0.0)
def _fireproof_elbow(x: FormulaInputs) -> tuple[float, float, float]:
    """Fireproof round elbow: Excel row 28."""
    # Compute and store: d.
    d = x.round_size
    # Compute and store: surface.
    surface = (2 * math.pi * ((x.w + d) ** 2 - d**2) + 2 * math.pi * x.h * (x.w + d + d)) * x.length / 360
    # Compute and store: seams.
    seams = ((x.w + x.h + 4 * x.insulation + 4 * x.fireproof) * 2 * 100 * x.Q * x.K + (x.w + x.h + 8 * x.insulation + 4 * x.fireproof) * 2 * 100 * x.Q * x.K)
    # Compute and store: r.
    r = (x.K * surface + seams) / 1_000_000 * x.insulation / 25
    # Compute and store: s.
    s = x.K * (2 * math.pi * ((x.w + 2 * x.insulation + d) ** 2 - d**2) + 2 * math.pi * (x.h + 2 * x.insulation) * (x.w + 2 * x.insulation + d + d)) * x.length / 360 / 1_000_000
    # Return the value produced by this step.
    return r, s, 0.0


########## Co vuông - chống cháy
# Purpose: Fireproof square elbow: Excel row 29.
# Example output: (12.345, 0.0, 0.0)
def _fireproof_square_elbow(x: FormulaInputs) -> tuple[float, float, float]:
    """Fireproof square elbow: Excel row 29."""
    # Compute and store: d.
    d = x.round_size
    # Compute and store: surface.
    surface = 2 * ((x.w + d) ** 2 - d**2) + 2 * (x.w + d) * x.h + 2 * d * x.h
    # Compute and store: seams.
    seams = ((x.w + x.h + 2 * x.insulation + 4 * x.fireproof) * 2 * 100 * x.K * x.Q + (x.w + x.h + 8 * x.insulation + 4 * x.fireproof) * 2 * 100 * x.K * x.Q)
    # Compute and store: r.
    r = (x.K * surface + seams) / 1_000_000 * x.insulation / 25
    # Compute and store: s.
    s = x.K * (2 * ((x.w + 2 * x.insulation + d) ** 2 - d**2) + 2 * (x.w + 2 * x.insulation + d) * (x.h + 2 * x.insulation) + 2 * d * (x.h + 2 * x.insulation)) / 1_000_000
    # Return the value produced by this step.
    return r, s, 0.0


########## Van MFD / FD / MSD - chống cháy
# Purpose: Fireproof MFD/FD/MSD valve: Excel row 30.
# Example output: (12.345, 0.0, 0.0)
def _fireproof_mfd(x: FormulaInputs) -> tuple[float, float, float]:
    """Fireproof MFD/FD/MSD valve: Excel row 30."""
    # Compute and store: surface.
    surface = 2 * (x.w + 50 + x.h + 50) * x.length
    # Compute and store: seam.
    seam = (x.w + 50 + x.h + 50 + x.insulation * 4) * 2 * 100 * x.K * x.Q
    # Return the value produced by this step.
    return (x.K * surface + seam) / 1_000_000 * x.insulation / 25, 0.0, 0.0


########## Van NRD / VCD / MD - chống cháy
# Purpose: Fireproof NRD/VCD/MD valve: Excel row 31.
# Example output: (12.345, 0.0, 0.0)
def _fireproof_nrd(x: FormulaInputs) -> tuple[float, float, float]:
    """Fireproof NRD/VCD/MD valve: Excel row 31."""
    # Compute and store: surface.
    surface = 2 * (x.w + x.h) * x.length
    # Keep the source row's first seam term exactly as written.
    seam = (x.w + x.h + 4 * x.insulation + 4 * x.fireproof) * 2 * 100 * x.K + (x.w + x.h + 8 * x.insulation + 4 * x.fireproof) * 2 * 100 * x.K * x.Q
    # Compute and store: r.
    r = (x.K * surface + seam) / 1_000_000 * x.insulation / 25
    # Compute and store: s.
    s = x.K * 2 * (x.w + x.h + 4 * x.insulation) * x.length / 1_000_000
    # Return the value produced by this step.
    return r, s, 0.0


_FIREPROOF = {
    "rectangular_duct": _fireproof_rectangular,
    "end_cap": _fireproof_end_cap,
    "reducer": _fireproof_reducer,
    "square_round_reducer": _fireproof_square_round_reducer,
    "tee": _fireproof_tee,
    "takeoff": _fireproof_takeoff,
    "elbow": _fireproof_elbow,
    "square_elbow": _fireproof_square_elbow,
    "mfd_valve": _fireproof_mfd,
    "nrd_valve": _fireproof_nrd,
}


# ---------------------------------------------------------------------------
# Acoustic / inner insulation formulas (Excel rows 34:49)
# ---------------------------------------------------------------------------


########## Ống gió chữ nhật - tiêu âm
# Purpose: Acoustic rectangular duct: Excel row 34.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_rectangular(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic rectangular duct: Excel row 34."""
    # Compute and store: first.
    first = (x.w + x.h + x.fireproof * 4) * 2 * x.Q * 100
    # Compute and store: second.
    second = (x.w + x.h + x.fireproof * 4 + 25 * 4) * 2 * x.Q * 100
    # Compute and store: r.
    r = first / 1_000_000 if x.fireproof == 0 and x.w <= 550 else (first + second) / 1_000_000
    # Compute and store: s.
    s = 0.0 if x.fireproof == 0 else (x.w + x.h) * 2 * x.K * 1000 / 1_000_000
    # Compute and store: t.
    t = (x.w + x.h) * 2 * x.K * 1000 / 1_000_000
    # Return the value produced by this step.
    return r, s, t


########## Bít đầu - tiêu âm
# Purpose: Acoustic end cap: Excel row 36.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_end_cap(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic end cap: Excel row 36."""
    # Compute and store: s.
    s = 0.0 if x.fireproof == 0 else (x.w + 2 * x.fireproof) * (x.h + 2 * x.fireproof) * x.K / 1_000_000
    # Return the value produced by this step.
    return 0.0, s, x.w * x.h * x.K / 1_000_000


########## Giảm chữ nhật - tiêu âm
# Purpose: Acoustic reducer: Excel row 37.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_reducer(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic reducer: Excel row 37."""
    # Compute and store: r.
    r = ((x.w + x.h + x.fireproof * 4) * 2 * 100 * x.K + (x.w + x.h + 4 * 25 + x.fireproof * 4) * 2 * x.Q * 100 * x.K) / 1_000_000
    # Compute and store: surface.
    surface = x.K * ((x.w + x.w1) * _sqrt(x.length**2 + ((x.h - x.h1) / 2) ** 2) + (x.h + x.h1) * _sqrt(x.length**2) + (x.w - x.w1 / 2) ** 2) / 1_000_000
    # Scc and T use the same formula for this product, so S = T.
    return r, surface, surface


########## Giảm vuông tròn - tiêu âm
# Purpose: Acoustic square-to-round reducer: Excel row 38.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_square_round_reducer(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic square-to-round reducer: Excel row 38."""
    # Compute and store: d.
    d = x.round_size
    # Compute and store: r.
    r = ((x.w + x.h + x.fireproof * 4) * 2 * 100 * x.K + (x.w + x.h + 4 * 25 + x.fireproof * 4) * 2 * 100 * x.Q * x.K * x.insulation / 25 + math.pi * (d + x.fireproof * 2) * 2 * x.K * 100 + math.pi * (d + x.fireproof * 2 + 2 * 25) * 2 * x.K * x.Q * 100) / 1_000_000
    # Compute and store: surface.
    surface = x.K * ((x.w + math.pi * d / 4) * _sqrt(x.length**2 + (x.h - math.pi * d / 4) ** 2) + (x.h + math.pi * d / 4) * _sqrt(x.length**2 + (x.w - math.pi * d / 4) ** 2)) / 1_000_000
    # Scc and T use the same formula for this product, so S = T.
    return r, surface, surface


########## Chạc 3 - tiêu âm
# Purpose: Acoustic three-way branch: Excel row 39.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_tee(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic three-way branch: Excel row 39."""
    # Compute and store: r.
    r = ((x.w1 + x.h1 + x.fireproof * 4) * 2 * 100 * x.K + (x.w1 + x.h1 + 4 * 25 + x.fireproof * 4) * 2 * 100 * x.K * x.Q + (x.w2 + x.h2 + x.fireproof * 4) * 2 * 100 * x.K + (x.w2 + x.h2 + x.fireproof * 4 + x.insulation * 4) * 100 * x.K * x.Q + (x.w + x.h + x.fireproof * 4) * 2 * 100 * x.K + (x.w + x.h + x.insulation * 4 + x.fireproof * 4) * 2 * 100 * x.K * x.Q) / 1_000_000
    # Compute and store: branch_1.
    branch_1 = 2 * (x.w1 + x.h1) * (x.length - 1.5 * x.w2)
    # Compute and store: branch_2.
    branch_2 = (2 * math.pi * ((x.w + x.w / 2) ** 2 - (x.w / 2) ** 2) + 2 * math.pi * x.h * ((x.w + x.w / 2) + x.w / 2)) * 90 / 360
    # Compute and store: branch_3.
    branch_3 = (2 * math.pi * ((x.w2 + x.w2 / 2) ** 2 - x.w2 / 2**2) + 2 * math.pi * x.h2 * ((x.w2 + x.w2 / 2) + x.w2 / 2)) * 90 / 360
    # Compute and store: surface.
    surface = x.K * (branch_1 + branch_2 + branch_3 - (x.h * x.w / 2 + x.h2 * x.w2 / 2)) / 1_000_000
    # Scc and T use the same formula for this product, so S = T.
    return r, surface, surface


########## Chân rẽ - tiêu âm
# Purpose: Acoustic takeoff branch: Excel row 40.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_takeoff(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic takeoff branch: Excel row 40."""
    # Compute and store: r.
    r = ((x.w1 + x.h1 + x.fireproof * 4) * 2 * 100 * x.K + (x.w1 + x.h1 + 4 * 25 + x.fireproof * 4) * 2 * 100 * x.K * x.Q) / 1_000_000
    # Compute and store: surface.
    surface = x.K * ((x.w + x.w1) * x.length + x.h1 * x.length + x.h * _sqrt((x.w - x.w1) ** 2 + x.length**2) - x.w * x.h) / 1_000_000
    # Scc and T use the same formula for this product, so S = T.
    return r, surface, surface


########## Co tròn - tiêu âm
# Purpose: Acoustic round elbow: Excel row 41.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_elbow(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic round elbow: Excel row 41."""
    # Compute and store: d.
    d = x.round_size
    # Compute and store: r.
    r = ((x.w + x.h + x.fireproof * 4) * 2 * 100 * x.K + (x.w + x.h + 4 * 25 + x.fireproof * 4) * 2 * 100 * x.K * x.Q) / 1_000_000
    # Compute and store: surface.
    surface = x.K * (2 * math.pi * ((x.w + d) ** 2 - d**2) + 2 * math.pi * x.h * (x.w + d + d)) * x.length / 360 / 1_000_000
    # Scc and T use the same formula for this product, so S = T.
    return r, surface, surface


########## Co vuông - tiêu âm
# Purpose: Acoustic square elbow: Excel row 42.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_square_elbow(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic square elbow: Excel row 42."""
    # Compute and store: d.
    d = x.round_size
    # Compute and store: r.
    r = ((x.w + x.h + x.fireproof * 4) * 2 * 100 * x.K + (x.w + x.h + 4 * 25 + x.fireproof * 4) * 2 * 100 * x.K * x.Q) / 1_000_000
    # Compute and store: surface.
    surface = x.K * (2 * ((x.w + d) ** 2 - d**2) + 2 * (x.w + d) * x.h + 2 * d * x.h) / 1_000_000
    # Scc and T use the same formula for this product, so S = T.
    return r, surface, surface


########## Van MFD / FD / MSD - tiêu âm
# Purpose: Acoustic MFD/FD/MSD valve: Excel row 43.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_mfd(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic MFD/FD/MSD valve: Excel row 43."""
    # Compute and store: r.
    r = ((x.w + x.h + 50 * 4) * 2 * x.length * x.K + (x.w + x.h + 50 * 4 + 25 * 4) * 2 * 100 * x.K * x.Q) / 1_000_000
    # Return the value produced by this step.
    return r, 0.0, 0.0


########## Van NRD / VCD / MD - tiêu âm
# Purpose: Acoustic NRD/VCD/MD valve: Excel row 44.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_nrd(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic NRD/VCD/MD valve: Excel row 44."""
    # Compute and store: r.
    r = ((x.w + x.h + x.fireproof * 4) * 2 * x.length * x.K + (x.w + x.h + 4 * 25 + x.fireproof * 4) * 2 * 100 * x.K * x.Q) / 1_000_000
    # Compute and store: s.
    s = 0.0 if x.fireproof == 0 else x.K * 2 * (x.w + x.h) * x.length / 1_000_000
    # Return the value produced by this step.
    return r, s, 0.0


########## Box cấp FCU
# Purpose: Acoustic FCU supply box: Excel row 45.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_fcu_box(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic FCU supply box: Excel row 45."""
    # Compute and store: r.
    r = ((x.w1 + x.h1) * 2 * 200 * x.K + (x.w1 + x.h1) * 2 * x.K * 100 + (x.w1 + x.h1 + 25 * 4) * 2 * 2 * x.K * 100) / 1_000_000
    # Compute and store: t.
    t = ((x.w + x.h + 4 * x.insulation) * 2 * x.length + 2 * (x.w + 2 * x.insulation) * (x.h + 2 * x.insulation) - x.w1 * x.h1 - x.Q * math.pi * x.round_size**2 + (x.w1 + x.h1) * 2 * 50) * x.K / 1_000_000
    # Return the value produced by this step.
    return r, 0.0, t


########## Box cấp giảm FCU
# Purpose: Acoustic reduced FCU box: Excel row 46.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_fcu_reducer_box(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic reduced FCU box: Excel row 46."""
    # Compute and store: r.
    r = ((x.fireproof + x.round_size) * 2 * x.K * 200 + (x.fireproof + x.round_size) * 2 * x.K * 100 + (x.fireproof + x.round_size + 25 * 4) * 2 * 2 * x.K * 100 + (x.w2 + x.h2) * 2 * x.K * x.Q * 100 + (x.w2 + x.h2 + 25 * 4) * 2 * x.K * x.Q * 100) / 1_000_000
    # Compute and store: t.
    t = ((x.w + x.w1 + 4 * x.insulation) * _sqrt((x.length + 2 * x.insulation) ** 2 + ((x.h - x.h1) / 2) ** 2) + (x.h + x.h1 + 4 * x.insulation) * _sqrt((x.length + 2 * x.insulation) ** 2) + (x.w - x.w1 / 2) ** 2 + (x.w + 2 * x.insulation) * (x.h + 2 * x.insulation) + (x.w1 + 2 * x.insulation) * (x.h1 + 2 * x.insulation) - x.w2 * x.h2 - x.fireproof * x.round_size + (x.w2 + x.h2) * 2 * 50 + (x.fireproof + x.round_size) * 2 * 50) * x.K / 1_000_000
    # Return the value produced by this step.
    return r, 0.0, t


########## Box hồi FCU
# Purpose: Acoustic FCU return box: Excel row 47.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_return_fcu_box(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic FCU return box: Excel row 47."""
    # Compute and store: r.
    r = ((x.w1 + x.h1) * 2 * 200 * x.K + (x.w1 + x.h1) * 2 * x.K * 100 + (x.w1 + x.h1 + 25 * 4) * 2 * 2 * x.K * 100) / 1_000_000
    # Compute and store: t.
    t = (2 * x.w * x.h + (x.w + x.h) * 2 * x.length - x.w1 * x.h1 - x.w2 * x.h2 - x.fireproof * math.pi * x.round_size**2 + (x.w1 + x.h1) * 2 * 50) * x.K / 1_000_000
    # Return the value produced by this step.
    return r, 0.0, t


########## Box gió phòng kỹ thuật / Louver
# Purpose: Acoustic technical-room/Louver box: Excel row 48.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_louver_box(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic technical-room/Louver box: Excel row 48."""
    # Compute and store: r.
    r = ((x.w1 + x.h1) * 2 * x.K * x.Q * 100 + (x.w1 + x.h1 + 25 * 4) * 2 * x.K * x.Q * 100 + (x.w2 + x.h2) * 2 * x.K * x.Q * 100 + (x.w2 + x.h2 + 25 * 4) * 2 * x.K * x.Q * 100) / 1_000_000
    # Compute and store: t.
    t = (2 * x.w * x.h + (x.w + x.h) * 2 * x.length - x.w1 * x.h1 - x.w2 * x.h2 + (x.w1 + x.h1) * 2 * 50 + (x.w2 + x.h2) * 2 * 50) * x.K / 1_000_000
    # Return the value produced by this step.
    return r, 0.0, t


########## Box cấp miệng gió
# Purpose: Acoustic air-outlet box: Excel row 49.
# Example output: (12.345, 0.0, 0.0)
def _acoustic_air_outlet_box(x: FormulaInputs) -> tuple[float, float, float]:
    """Acoustic air-outlet box: Excel row 49."""
    # Compute and store: t.
    t = (2 * x.w * x.h + (x.w + x.h) * 2 * x.length - x.w1 * x.h1 - math.pi * x.round_size**2) * x.K / 1_000_000
    # Return the value produced by this step.
    return 0.0, 0.0, t






############################
_ACOUSTIC = {
    "rectangular_duct": _acoustic_rectangular,
    "end_cap": _acoustic_end_cap,
    "reducer": _acoustic_reducer,
    "square_round_reducer": _acoustic_square_round_reducer,
    "tee": _acoustic_tee,
    "takeoff": _acoustic_takeoff,
    "elbow": _acoustic_elbow,
    "square_elbow": _acoustic_square_elbow,
    "mfd_valve": _acoustic_mfd,
    "nrd_valve": _acoustic_nrd,
    "fcu_box": _acoustic_fcu_box,
    "fcu_reducer_box": _acoustic_fcu_reducer_box,
    "return_fcu_box": _acoustic_return_fcu_box,
    "technical_louver_box": _acoustic_louver_box,
    "air_outlet_box": _acoustic_air_outlet_box,
}


_CALCULATORS = {
    "standard": _STANDARD,
    "fireproof": _FIREPROOF,
    "acoustic": _ACOUSTIC,
}


# Purpose: Calculate R/S/T for one explicitly selected product type.
# Example output: CalculationResult(section="standard", product_type="rectangular_duct", R=15.12, S=0.0, T=0.0)
def calculate(section: str, product_type: str, inputs: Mapping[str, Any]) -> CalculationResult:
    """Calculate R/S/T for one explicitly selected product type."""
    # Compute and store: normalized_section.
    normalized_section = str(section).strip().lower()
    # Compute and store: normalized_product.
    normalized_product = str(product_type).strip().lower()
    # Compute and store: products.
    products = _CALCULATORS.get(normalized_section)

    # Check this condition before choosing the next branch.
    if products is None:
        # Stop with a clear validation error for invalid input.
        raise ValueError(f"Unsupported section: {section}")
    
    # Compute and store: calculator.
    calculator = products.get(normalized_product)

    # Check this condition before choosing the next branch.
    if calculator is None:
        # Stop with a clear validation error for invalid input.
        raise ValueError(f"Unsupported {normalized_section} product_type: {product_type}")
    
    # Compute and store: (r, s, t).
    r, s, t = calculator(_read_inputs(inputs))
    # Return the value produced by this step.
    return CalculationResult(
        normalized_section,
        normalized_product,
        round(r, 3),
        round(s, 3),
        round(t, 3),
    )


# Formula text is kept in a separate module so numeric calculations remain
# untouched, while the public engine API exposes one consistent dispatcher.
from . import formula_strings as _formula_strings

excel_input_references = _formula_strings.excel_input_references
formula_strings = _formula_strings.formula_strings
for _formula_name in _formula_strings.__all__:
    if _formula_name not in {"excel_input_references", "formula_strings"}:
        globals()[_formula_name] = getattr(_formula_strings, _formula_name)


__all__ = [
    "CalculationResult",
    "FormulaInputs",
    "calculate",
    *_formula_strings.__all__,
]
