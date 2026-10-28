"""Small safe evaluators for quantity expressions stored in Excel cells."""

from __future__ import annotations

import re
from typing import Any, Mapping

import pandas as pd


_ALLOWED_EXPRESSION = re.compile(r"[0-9.+*/()\\-]+")
_K_COLUMNS = ("plan", "up_down", "vertical_axis")


# Purpose: Evaluate numbers with +, -, *, / and parentheses from one cell.
# Example output: 0.92
def evaluate_expression(expression: Any) -> float:
    """Evaluate numbers with +, -, *, / and parentheses from one cell."""
    # Check this condition before choosing the next branch.
    if expression is None or (not isinstance(expression, str) and pd.isna(expression)):
        # Return the value produced by this step.
        return 0.0

    # Compute and store: text.
    text = str(expression).replace(",", ".").replace(" ", "")
    # Check this condition before choosing the next branch.
    if not text:
        # Return the value produced by this step.
        return 0.0
    # Check this condition before choosing the next branch.
    if not _ALLOWED_EXPRESSION.fullmatch(text):
        # Stop with a clear validation error for invalid input.
        raise ValueError(f"Invalid quantity expression: {expression}")

    # Return the value produced by this step.
    return float(eval(text, {"__builtins__": {}}, {}))


# Purpose: Sum plan, up/down, and vertical-axis quantities for one row.
# Example output: 1.985
def calculate_K(row: Mapping[str, Any]) -> float:
    """Sum plan, up/down, and vertical-axis quantities for one row."""
    # Return the value produced by this step.
    return sum(evaluate_expression(row.get(column)) for column in _K_COLUMNS)


__all__ = ["calculate_K", "evaluate_expression"]
