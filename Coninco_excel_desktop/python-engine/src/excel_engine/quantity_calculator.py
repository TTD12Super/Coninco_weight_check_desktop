"""Small safe evaluators for quantity expressions stored in Excel cells."""

from __future__ import annotations

import re
from typing import Any, Mapping

import pandas as pd


_ALLOWED_EXPRESSION = re.compile(r"[0-9.+*/()\\-]+")
_K_COLUMNS = ("plan", "up_down", "vertical_axis")


def evaluate_expression(expression: Any) -> float:
    """Evaluate numbers with +, -, *, / and parentheses from one cell."""
    if expression is None or (not isinstance(expression, str) and pd.isna(expression)):
        return 0.0

    text = str(expression).replace(",", ".").replace(" ", "")
    if not text:
        return 0.0
    if not _ALLOWED_EXPRESSION.fullmatch(text):
        raise ValueError(f"Invalid quantity expression: {expression}")

    return float(eval(text, {"__builtins__": {}}, {}))


def calculate_K(row: Mapping[str, Any]) -> float:
    """Sum plan, up/down, and vertical-axis quantities for one row."""
    return sum(evaluate_expression(row.get(column)) for column in _K_COLUMNS)


__all__ = ["calculate_K", "evaluate_expression"]
