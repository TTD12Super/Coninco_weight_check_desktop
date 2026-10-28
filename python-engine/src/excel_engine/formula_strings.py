"""Excel formula strings corresponding to the numeric formula engine.

The numeric functions remain the source of truth for calculation.  This
module mirrors their expressions for reporting/auditing and never evaluates a
formula.  By default formulas use workbook-style symbolic names (``W``, ``H``
...).  ``formula_strings(..., references=...)`` can replace those names with
real cell references for a particular row.
"""

from __future__ import annotations

import re
from typing import Mapping


# Each entry contains formulas for the numeric function's outputs in R/S/T
# order.  A zero output is represented as the valid Excel formula ``=0``.
_FORMULAS: dict[str, tuple[str, str, str]] = {
    # Standard
    "_standard_rectangular": (
        "=((W+H)*2*K*1000+(W+H+25*4)*2*Q*100)/1000000", "=0", "=0",
    ),
    "_standard_round": (
        "=(PI()*R_D*K*1000+PI()*(R_D+25*2)*Q*100)/1000000", "=0", "=0",
    ),
    "_standard_end_cap": ("=(W+50)*(H+50)*K/1000000", "=0", "=0"),
    "_standard_reducer": (
        "=(K*((W+W1)*SQRT(L_M^2+((H-H1)/2)^2)+(H+H1)*SQRT(L_M^2)+(W-W1/2)^2)+(W+H+25*4)*2*Q*100*K)/1000000",
        "=0", "=0",
    ),
    "_standard_square_round_reducer": (
        "=(K*((W+PI()*R_D/4)*SQRT(L_M^2+(H-PI()*R_D/4)^2)+(H+PI()*R_D/4)*SQRT(L_M^2+(W-PI()*R_D/4)^2))+(W+H+25*4)*2*100*Q*K+PI()*(R_D+25*2)*K*Q*100)/1000000",
        "=0", "=0",
    ),
    "_standard_tee": (
        "=(K*(2*(W1+H1)*(L_M-1.5*W2)+(2*PI()*((W+W/2)^2-(W/2)^2)+2*PI()*H*((W+W/2)+W/2))*90/360+(2*PI()*((W2+W2/2)^2-W2/2^2)+2*PI()*H2*((W2+W2/2)+W2/2))*90/360-(H*W/2+H2*W2/2))+(W1+H1+25*4)*2*2*100*K*Q)/1000000",
        "=0", "=0",
    ),
    "_standard_takeoff": (
        "=K*((W+W1)*L_M+H1*L_M+H*SQRT((W-W1)^2+L_M^2)-W*H+(W1+H1+25*4)*2*100*K*Q)/1000000",
        "=0", "=0",
    ),
    "_standard_elbow": (
        "=(K*(2*PI()*((W+R_D)^2-R_D^2)+2*PI()*H*(W+R_D+R_D))*L_M/360+(W+H+25*4)*2*100*K*Q)/1000000",
        "=0", "=0",
    ),
    "_standard_square_elbow": (
        "=(K*(2*((W+R_D)^2-R_D^2)+2*(W+R_D)*H+2*R_D*H)+(W+H+25*4)*2*100*K*Q)/1000000",
        "=0", "=0",
    ),
    "_standard_mfd": (
        "=(K*(2*((W+50)+(H+50))*L_M)+(W+50+H+50+25*4)*2*100*K*Q)/1000000",
        "=0", "=0",
    ),
    "_standard_nrd": (
        "=(K*2*(W+H)*L_M+(W+H+25*4)*2*100*K*Q)/1000000", "=0", "=0",
    ),
    "_standard_stiffener": ("=(W+H+25*4)*2*100*Q/1000000", "=0", "=0"),

    # Fireproof
    "_fireproof_rectangular": (
        "=(((W+H)*2*K*1000+(W+H+4*a+4*t)*2*Q*100+(W+H+8*a+4*t)*2*Q*100)/1000000)*a/25",
        "=(W+H+4*a)*2*K*1000/1000000", "=0",
    ),
    "_fireproof_end_cap": (
        "=(W+2*a)*(H+2*a)*K/1000000*a/25",
        "=(W+2*a+2*t)*(H+2*a+2*t)*K/1000000", "=0",
    ),
    "_fireproof_reducer": (
        "=((K*((W+W1)*SQRT(L_M^2+((H-H1)/2)^2)+(H+H1)*SQRT(L_M^2+((W-W1)/2)^2))+(W+H+4*a+4*t)*2*Q*100*K+(W+H+8*a+4*t)*2*Q*100*K)/1000000)*a/25",
        "=K*((W+W1+4*a)*SQRT(L_M^2+((H-H1)/2)^2)+(H+H1+4*a)*SQRT(L_M^2+(W-W1/2)^2))/1000000", "=0",
    ),
    "_fireproof_square_round_reducer": (
        "=(K*((W+PI()*R_D/4)*SQRT(L_M^2+(H-PI()*R_D/4)^2)+(H+PI()*R_D/4)*SQRT(L_M^2+(W-PI()*R_D/4)^2))+(W+H+4*a+4*t)*2*100*Q*K+(W+H+8*a+4*t)*2*100*Q*K+PI()*(R_D+4*a+4*t)*2*K*Q*100)/1000000*a/25",
        "=K*((W+2*a+PI()*(R_D+2*a)/4)*SQRT(L_M^2+(H+2*a-PI()*(R_D+2*a)/4)^2)+(H+2*a+PI()*(R_D+2*a)/4)*SQRT(L_M^2+(W+2*a-PI()*(R_D+2*a)/4)^2))/1000000", "=0",
    ),
    "_fireproof_tee": (
        "=((K*(2*(W1+H1)*(L_M-1.5*W2)+(2*PI()*((W+W/2)^2-(W/2)^2)+2*PI()*H*((W+W/2)+W/2))*90/360+(2*PI()*((W2+W2/2)^2-W2/2^2)+2*PI()*H2*((W2+W2/2)+W2/2))*90/360-(H*W/2+H2*W2/2))+(W1+H1+4*a+4*t)*2*2*100*K*Q+(W1+H1+8*a+4*t)*2*2*100*K*Q)/1000000)*a/25",
        "=K*(2*(W1+H1+4*a)*(L_M+2*a-1.5*(W2+2*a))+(2*PI()*((W+2*a+(W+2*a)/2)^2-((W+2*a)/2)^2)+2*PI()*(H+2*a)*((W+2*a+(W+2*a)/2)+(W+2*a)/2))*90/360+(2*PI()*((W2+2*a+(W2+2*a)/2)^2-(W2+2*a)/2^2)+2*PI()*(H2+2*a)*((W2+2*a+(W2+2*a)/2)+(W2+2*a)/2))*90/360-((H+2*a)*(W+2*a)/2+(H2+2*a)*(W2+2*a)/2))/1000000", "=0",
    ),
    "_fireproof_takeoff": (
        "=K*((W+W1)*L_M+H1*L_M+H*SQRT((W-W1)^2+L_M^2)-W*H+(W1+H1+4*a+4*t)*2*100*K*Q+(W1+H1+8*a+4*t)*2*100*K*Q)/1000000*a/25",
        "=K*((W+W1+4*a)*L_M+(H1+2*a)*L_M+(H+2*a)*SQRT((W-W1)^2+L_M^2)-W*H)/1000000", "=0",
    ),
    "_fireproof_elbow": (
        "=(K*(2*PI()*((W+R_D)^2-R_D^2)+2*PI()*H*(W+2*R_D))*L_M/360+(W+H+4*a+4*t)*2*100*K*Q+(W+H+8*a+4*t)*2*100*K*Q)/1000000*a/25",
        "=K*(2*PI()*((W+2*a+R_D)^2-R_D^2)+2*PI()*(H+2*a)*(W+2*a+2*R_D))*L_M/360/1000000", "=0",
    ),
    "_fireproof_square_elbow": (
        "=((K*(2*((W+R_D)^2-R_D^2)+2*(W+R_D)*H+2*R_D*H)+(W+H+2*a+4*t)*2*100*K*Q+(W+H+8*a+4*t)*2*100*K*Q)/1000000)*a/25",
        "=K*(2*((W+2*a+R_D)^2-R_D^2)+2*(W+2*a+R_D)*(H+2*a)+2*R_D*(H+2*a))/1000000", "=0",
    ),
    "_fireproof_mfd": (
        "=(K*2*(W+50+H+50)*L_M+(W+50+H+50+4*a)*2*100*K*Q)/1000000*a/25", "=0", "=0",
    ),
    "_fireproof_nrd": (
        "=(K*2*(W+H)*L_M+(W+H+4*a+4*t)*2*100*K+(W+H+8*a+4*t)*2*100*K*Q)/1000000*a/25",
        "=K*2*(W+H+4*a)*L_M/1000000", "=0",
    ),

    # Acoustic
    "_acoustic_rectangular": (
        "=IF(AND(t=0,W<=550),(W+H+4*t)*2*Q*100/1000000,((W+H+4*t)*2*Q*100+(W+H+4*t+25*4)*2*Q*100)/1000000)",
        "=IF(t=0,0,(W+H)*2*K*1000/1000000)", "=(W+H)*2*K*1000/1000000",
    ),
    "_acoustic_end_cap": (
        "=0", "=IF(t=0,0,(W+2*t)*(H+2*t)*K/1000000)", "=W*H*K/1000000",
    ),
    "_acoustic_reducer": (
        "=((W+H+4*t)*2*100*K+(W+H+4*25+4*t)*2*Q*100*K)/1000000",
        "=K*((W+W1)*SQRT(L_M^2+((H-H1)/2)^2)+(H+H1)*SQRT(L_M^2)+(W-W1/2)^2)/1000000",
        "=K*((W+W1)*SQRT(L_M^2+((H-H1)/2)^2)+(H+H1)*SQRT(L_M^2)+(W-W1/2)^2)/1000000",
    ),
    "_acoustic_square_round_reducer": (
        "=((W+H+4*t)*2*100*K+(W+H+4*25+4*t)*2*100*Q*K*a/25+PI()*(R_D+2*t)*2*K*100+PI()*(R_D+2*t+2*25)*2*K*Q*100)/1000000",
        "=K*((W+PI()*R_D/4)*SQRT(L_M^2+(H-PI()*R_D/4)^2)+(H+PI()*R_D/4)*SQRT(L_M^2+(W-PI()*R_D/4)^2))/1000000",
        "=K*((W+PI()*R_D/4)*SQRT(L_M^2+(H-PI()*R_D/4)^2)+(H+PI()*R_D/4)*SQRT(L_M^2+(W-PI()*R_D/4)^2))/1000000",
    ),
    "_acoustic_tee": (
        "=((W1+H1+4*t)*2*100*K+(W1+H1+4*25+4*t)*2*100*K*Q+(W2+H2+4*t)*2*100*K+(W2+H2+4*t+4*a)*100*K*Q+(W+H+4*t)*2*100*K+(W+H+4*a+4*t)*2*100*K*Q)/1000000",
        "=K*(2*(W1+H1)*(L_M-1.5*W2)+(2*PI()*((W+W/2)^2-(W/2)^2)+2*PI()*H*((W+W/2)+W/2))*90/360+(2*PI()*((W2+W2/2)^2-W2/2^2)+2*PI()*H2*((W2+W2/2)+W2/2))*90/360-(H*W/2+H2*W2/2))/1000000",
        "=K*(2*(W1+H1)*(L_M-1.5*W2)+(2*PI()*((W+W/2)^2-(W/2)^2)+2*PI()*H*((W+W/2)+W/2))*90/360+(2*PI()*((W2+W2/2)^2-W2/2^2)+2*PI()*H2*((W2+W2/2)+W2/2))*90/360-(H*W/2+H2*W2/2))/1000000",
    ),
    "_acoustic_takeoff": (
        "=((W1+H1+4*t)*2*100*K+(W1+H1+4*25+4*t)*2*100*K*Q)/1000000",
        "=K*((W+W1)*L_M+H1*L_M+H*SQRT((W-W1)^2+L_M^2)-W*H)/1000000",
        "=K*((W+W1)*L_M+H1*L_M+H*SQRT((W-W1)^2+L_M^2)-W*H)/1000000",
    ),
    "_acoustic_elbow": (
        "=((W+H+4*t)*2*100*K+(W+H+4*25+4*t)*2*100*K*Q)/1000000",
        "=K*(2*PI()*((W+R_D)^2-R_D^2)+2*PI()*H*(W+2*R_D))*L_M/360/1000000",
        "=K*(2*PI()*((W+R_D)^2-R_D^2)+2*PI()*H*(W+2*R_D))*L_M/360/1000000",
    ),
    "_acoustic_square_elbow": (
        "=((W+H+4*t)*2*100*K+(W+H+4*25+4*t)*2*100*K*Q)/1000000",
        "=K*(2*((W+R_D)^2-R_D^2)+2*(W+R_D)*H+2*R_D*H)/1000000",
        "=K*(2*((W+R_D)^2-R_D^2)+2*(W+R_D)*H+2*R_D*H)/1000000",
    ),
    "_acoustic_mfd": (
        "=((W+H+50*4)*2*L_M*K+(W+H+50*4+25*4)*2*100*K*Q)/1000000", "=0", "=0",
    ),
    "_acoustic_nrd": (
        "=((W+H+4*t)*2*L_M*K+(W+H+4*25+4*t)*2*100*K*Q)/1000000",
        "=IF(t=0,0,K*2*(W+H)*L_M/1000000)", "=0",
    ),
    "_acoustic_fcu_box": (
        "=((W1+H1)*2*200*K+(W1+H1)*2*K*100+(W1+H1+25*4)*2*2*K*100)/1000000", "=0",
        "=((W+H+4*a)*2*L_M+2*(W+2*a)*(H+2*a)-W1*H1-Q*PI()*R_D^2+(W1+H1)*2*50)*K/1000000",
    ),
    "_acoustic_fcu_reducer_box": (
        "=((t+R_D)*2*K*200+(t+R_D)*2*K*100+(t+R_D+25*4)*2*2*K*100+(W2+H2)*2*K*Q*100+(W2+H2+25*4)*2*K*Q*100)/1000000", "=0",
        "=((W+W1+4*a)*SQRT((L_M+2*a)^2+((H-H1)/2)^2)+(H+H1+4*a)*SQRT((L_M+2*a)^2)+(W-W1/2)^2+(W+2*a)*(H+2*a)+(W1+2*a)*(H1+2*a)-W2*H2-t*R_D+(W2+H2)*2*50+(t+R_D)*2*50)*K/1000000",
    ),
    "_acoustic_return_fcu_box": (
        "=((W1+H1)*2*200*K+(W1+H1)*2*K*100+(W1+H1+25*4)*2*2*K*100)/1000000", "=0",
        "=(2*W*H+(W+H)*2*L_M-W1*H1-W2*H2-t*PI()*R_D^2+(W1+H1)*2*50)*K/1000000",
    ),
    "_acoustic_louver_box": (
        "=((W1+H1)*2*K*Q*100+(W1+H1+25*4)*2*K*Q*100+(W2+H2)*2*K*Q*100+(W2+H2+25*4)*2*K*Q*100)/1000000", "=0",
        "=(2*W*H+(W+H)*2*L_M-W1*H1-W2*H2+(W1+H1)*2*50+(W2+H2)*2*50)*K/1000000",
    ),
    "_acoustic_air_outlet_box": ("=0", "=0", "=(2*W*H+(W+H)*2*L_M-W1*H1-PI()*R_D^2)*K/1000000"),
}


_SYMBOLS = ("W", "H", "W1", "H1", "W2", "H2", "L_M", "R_D", "a", "t", "Q", "K")
_SOURCE_COLUMNS = {"W": "F", "H": "G", "W1": "H", "H1": "I", "W2": "J", "H2": "K", "L_M": "L", "R_D": "M", "a": "N", "t": "O", "Q": "P", "K": "Q"}

_PRODUCT_SUFFIX = {
    "rectangular_duct": "rectangular",
    "round_duct": "round",
    "end_cap": "end_cap",
    "reducer": "reducer",
    "square_round_reducer": "square_round_reducer",
    "tee": "tee",
    "takeoff": "takeoff",
    "elbow": "elbow",
    "square_elbow": "square_elbow",
    "mfd_valve": "mfd",
    "nrd_valve": "nrd",
    "stiffener": "stiffener",
    "fcu_box": "fcu_box",
    "fcu_reducer_box": "fcu_reducer_box",
    "return_fcu_box": "return_fcu_box",
    "technical_louver_box": "louver_box",
    "air_outlet_box": "air_outlet_box",
}


# Purpose: Replace symbolic input names with row-specific Excel cell references.
# Example output: "2*F21*G21"
def _replace_references(formula: str, references: Mapping[str, str] | None) -> str:
    """Replace symbolic input names with row-specific Excel cell references."""
    # Check this condition before choosing the next branch.
    if not references:
        # Return the value produced by this step.
        return formula
    # Compute and store: result.
    result = formula
    # Process each item in this input sequence.
    for symbol in sorted(_SYMBOLS, key=len, reverse=True):
        # Compute and store: reference.
        reference = references.get(symbol)
        # Check this condition before choosing the next branch.
        if reference:
            # Compute and store: result.
            result = re.sub(rf"(?<![A-Za-z0-9_]){re.escape(symbol)}(?![A-Za-z0-9_])", reference, result)
    # Return the value produced by this step.
    return result


# Purpose: Return source-sheet cell references for ``FormulaInputs`` at ``row``.
# Example output: {"W": "F21", "H": "G21", "K": "Q21"}
def excel_input_references(row: int, column_map: Mapping[str, str] | None = None) -> dict[str, str]:
    """Return source-sheet cell references for ``FormulaInputs`` at ``row``."""
    # Compute and store: columns.
    columns = dict(_SOURCE_COLUMNS)
    # Check this condition before choosing the next branch.
    if column_map:
        # Execute this expression for its result or side effect.
        columns.update(column_map)
    # Return the value produced by this step.
    return {symbol: f"{columns[symbol]}{row}" for symbol in _SYMBOLS}


# Purpose: Return the Excel R/S/T formulas for one section and product type.
# Example output: ("=2*F21*G21", "=0", "=0")
def formula_strings(
    section: str,
    product_type: str,
    references: Mapping[str, str] | None = None,
) -> tuple[str, str, str]:
    """Return the Excel R/S/T formulas for one section and product type."""
    # Compute and store: normalized_product.
    normalized_product = str(product_type).strip().lower()
    # Compute and store: suffix.
    suffix = _PRODUCT_SUFFIX.get(normalized_product, normalized_product)
    # Compute and store: key.
    key = f"_{str(section).strip().lower()}_{suffix}"
    # Protect parsing/evaluation so bad workbook data becomes a controlled error.
    try:
        # Compute and store: formulas.
        formulas = _FORMULAS[key]
    except KeyError as error:
        # Stop with a clear validation error for invalid input.
        raise ValueError(f"No Excel formula strings for {section}/{product_type}") from error
    # Return the value produced by this step.
    return tuple(_replace_references(formula, references) for formula in formulas)  # type: ignore[return-value]


# Purpose: Create one public wrapper for an R, S, or T formula string.
# Example output: <function _standard_rectangular_formula>
def _make_formula_function(function_name: str, output_index: int):
    """Create one public wrapper for an R, S, or T formula string."""
    # Purpose: Return the selected output formula with optional cell references.
    # Example output: "2*F21*G21"
    def formula_function(references: Mapping[str, str] | None = None) -> str:
        """Return the selected output formula with optional cell references."""
        # Compute and store: formulas.
        formulas = _FORMULAS[function_name]
        # Return the value produced by this step.
        return _replace_references(formulas[output_index], references)

    # Compute and store: formula_function.__name__.
    formula_function.__name__ = f"{function_name}_formula" if output_index == 0 else f"{function_name}_{'s' if output_index == 1 else 't'}_formula"
    # Compute and store: formula_function.__doc__.
    formula_function.__doc__ = f"Return the Excel formula for {function_name} output {output_index}."
    # Return the value produced by this step.
    return formula_function


for _function_name in _FORMULAS:
    globals()[f"{_function_name}_formula"] = _make_formula_function(_function_name, 0)
    globals()[f"{_function_name}_s_formula"] = _make_formula_function(_function_name, 1)
    globals()[f"{_function_name}_t_formula"] = _make_formula_function(_function_name, 2)


__all__ = [
    "excel_input_references",
    "formula_strings",
    *[f"{name}_formula" for name in _FORMULAS],
    *[f"{name}_s_formula" for name in _FORMULAS],
    *[f"{name}_t_formula" for name in _FORMULAS],
]
