"""Small, safe parser for contractor formula descriptions.

The workbook is allowed to change the expression on each row.  This module
parses the supported Excel-like expression language instead of evaluating
untrusted strings with ``eval``.
"""

from __future__ import annotations

from dataclasses import dataclass
import math
import re
import unicodedata
from functools import lru_cache
from typing import Any, Mapping

import pandas as pd


class FormulaParseError(ValueError):
    """Raised when a formula cannot be tokenized or parsed."""


class FormulaEvaluationError(ValueError):
    """Raised when a valid formula cannot be evaluated for a row."""


@dataclass(frozen=True)
class Token:
    kind: str
    value: str
    position: int


@dataclass(frozen=True)
class Node:
    kind: str
    value: Any = None
    children: tuple["Node", ...] = ()


_TOKEN_RE = re.compile(
    r"\s*(?:(?P<number>(?:\d+(?:\.\d*)?|\.\d+))|"
    # W1/H1/W2/H2 are dimension names in contractor formulas. Other letter+
    # digit strings (for example E6) remain Excel cell references.
    r"(?P<dimension>[WH][12](?![A-Za-z0-9_']))|"
    r"(?P<cell>\$?[A-Za-z]{1,3}\$?\d+)|"
    r"(?P<name>[A-Za-z_][A-Za-z0-9_']*)|"
    r"(?P<compare><>|<=|>=|=|<|>)|"
    r"(?P<operator>[+\-*/^(),]))"
)   



##### Tim cac cong thuc dien giai
_ASSIGNMENT_RE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<label>SCC/TA|SCN|SCC|STA|S)\s*=",
    re.IGNORECASE,
)

# Intermediate quantities in contractor descriptions are normally written as
# ``S_V1=...`` on their own line.  They are not output columns, but they are
# part of the same formula block and can be referenced by the final ``S=...``.
_INTERMEDIATE_ASSIGNMENT_RE = re.compile(
    r"(?<![A-Za-z0-9_])(?P<label>S_[A-Za-z][A-Za-z0-9_']*)\s*=",
    re.IGNORECASE,
)

_VARIABLE_ALIASES = {
    "W": "W",
    "H": "H",
    "W'": "W1",
    "H'": "H1",
    "W''": "W2",
    "H''": "H2",
    "L": "L_M",
    "M": "L_M",
    "R": "R_D",
    "D": "R_D",
    "A": "a",
    "T": "t",
    "Q": "Q",
    "K": "K",
}

_OUTPUT_TARGET_LABELS = frozenset({"S", "SCC", "SCN", "STA", "SCC/TA"})


# Purpose: Normalize workbook symbols before tokenization.
# Example output: "2*PI()*W"
def _normalize_source(text: str) -> str:
    """Normalize workbook symbols before tokenization."""
    # Compute and store: source.
    source = unicodedata.normalize("NFKC", str(text))
    # Compute and store: source.
    source = (
        source
        .replace("\u00d7", "*")
        .replace("\u00f7", "/")
        .replace("\u2019", "'")
        .replace("\u2018", "'")
        .replace("\u2212", "-")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\r", "")
        .replace("\u00a0", " ")
    )

    # Natural-language multiplication: ASCII x, uppercase X, or Cyrillic х.
    # Restrict replacement to an operator position so names such as MAX stay intact.
    source = re.sub(
        r"(?<=[\w.)])\s*[xXхХ]\s*(?=[\w(])",
        "*",
        source,
    )

    # Compute and store: source.
    source = re.sub(
        r"(?<![A-Za-z0-9_,(])(\d+),(\d+)(?!\s*,)",
        r"\1.\2",
        source,
    )


    # Return the value produced by this step.
    return source.replace(";", ",")


# Purpose: Extract a formula block's labelled expressions.
# Example output: [("S_V1", "2*W"), ("S", "K*S_V1")]
def extract_assignments(text: Any) -> list[tuple[str, str]]:
    """Extract a formula block's labelled expressions.

    Besides output labels such as ``SCN`` and ``STA``, this recognizes
    intermediate labels such as ``S_V1``.  The latter are kept in the returned
    block so the evaluator can resolve dependencies before calculating ``S``.
    """
    

    ####### Kiem tra xem text dau vao co bi rong hoac khong phai chuoi hay khong, Neu rong thi tra ve List rong
    if text is None or (not isinstance(text, str) and pd.isna(text)):
        # Return the value produced by this step.
        return []
    
    #### Chuan hoa chuoi de dung voi cong thuc trong excel
    source = _normalize_source(str(text))

    #### Tim cac cong thuc dien giai
    matches = list(_ASSIGNMENT_RE.finditer(source))

    
    # Execute this expression for its result or side effect.
    matches.extend(_INTERMEDIATE_ASSIGNMENT_RE.finditer(source))

    # Execute this expression for its result or side effect.
    matches.sort(key=lambda match: match.start())
    # Check this condition before choosing the next branch.
    if not matches:
        # Compute and store: expression.
        expression = source.strip()
        # Return the value produced by this step.
        return [("S", expression)] if expression else []

    # Compute and store: assignments.
    assignments: list[tuple[str, str]] = []
    
    # Process each item in this input sequence.
    for position, match in enumerate(matches):
        # Compute and store: start.
        start = match.end()
        # Compute and store: end.
        end = matches[position + 1].start() if position + 1 < len(matches) else len(source)
        # Compute and store: expression.
        expression = source[start:end].strip().strip(";,")
        # Check this condition before choosing the next branch.
        if expression:
            # Execute this expression for its result or side effect.
            assignments.append((match.group("label").upper(), expression))
    # Return the value produced by this step.
    return assignments


# Purpose: Repair a missing opening parenthesis without changing precedence.
# Example output: "(W+H)*K/1000000"
def _repair_extra_closing_parentheses(source: str) -> str:
    """Repair a missing opening parenthesis without changing precedence.

    Contractor workbooks occasionally contain formulas such as
    ``(A+B)/1000000)``.  Dropping the final ``)`` would turn it into
    ``A+B/1000000`` and change the calculation.  When the whole expression has
    more closing than opening parentheses, prepend the missing opening groups
    so the original closing parenthesis still groups the numerator.  Expressions
    with missing closing parentheses are left untouched and remain visible as
    parse errors.
    """
    # Compute and store: missing_openings.
    missing_openings = source.count(")") - source.count("(")
    # Check this condition before choosing the next branch.
    if missing_openings > 0:
        # Return the value produced by this step.
        return "(" * missing_openings + source
    # Return the value produced by this step.
    return source


# Purpose: Convert one normalized expression into parser tokens.
# Example output: [Token(kind="number", value="2", position=0), Token(kind="op", value="*", position=2)]
def _tokens(expression: str) -> list[Token]:
    """Convert one normalized expression into parser tokens."""
    # Compute and store: source.
    source = _repair_extra_closing_parentheses(_normalize_source(expression))
    # Compute and store: result.
    result: list[Token] = []
    # Compute and store: position.
    position = 0
    # Repeat while this condition remains true.
    while position < len(source):
        # Compute and store: match.
        match = _TOKEN_RE.match(source, position)
        # Check this condition before choosing the next branch.
        if not match:
            # Stop with a clear validation error for invalid input.
            raise FormulaParseError(f"Unexpected character {source[position]!r} at {position}.")
        # Compute and store: kind.
        kind = next(name for name, value in match.groupdict().items() if value is not None)
        # Execute this expression for its result or side effect.
        result.append(Token(kind, match.group(kind), position))
        # Compute and store: position.
        position = match.end()
    # Execute this expression for its result or side effect.
    result.append(Token("eof", "", len(source)))
    # Return the value produced by this step.
    return result


class _Parser:
    # Purpose: Tokenize one expression and start the recursive-descent parser.
    # Example output: constructor returns None; ``self.tokens`` is initialized.
    def __init__(self, expression: str):
        """Tokenize one expression and start the recursive-descent parser."""
        # Compute and store: self.expression.
        self.expression = expression
        # Compute and store: self.tokens.
        self.tokens = _tokens(expression)
        # Compute and store: self.index.
        self.index = 0

    @property
    # Purpose: Return the token currently being parsed.
    # Example output: Token(kind="number", value="2", position=0)
    def current(self) -> Token:
        """Return the token currently being parsed."""
        # Return the value produced by this step.
        return self.tokens[self.index]

    # Purpose: Consume and return the current token.
    # Example output: Token(kind="number", value="2", position=0)
    def advance(self) -> Token:
        """Consume and return the current token."""
        # Compute and store: token.
        token = self.current
        # Update the accumulated value: self.index.
        self.index += 1
        # Return the value produced by this step.
        return token

    # Purpose: Consume a token when its value matches.
    # Example output: True
    def accept(self, value: str) -> bool:
        """Consume a token when its value matches."""
        # Check this condition before choosing the next branch.
        if self.current.value.upper() == value.upper():
            # Execute this expression for its result or side effect.
            self.advance()
            # Return the value produced by this step.
            return True
        # Return the value produced by this step.
        return False

    # Purpose: Require a specific token or raise a parse error.
    def expect(self, value: str) -> None:
        """Require a specific token or raise a parse error."""
        # Check this condition before choosing the next branch.
        if not self.accept(value):
            # Stop with a clear validation error for invalid input.
            raise FormulaParseError(
                f"Expected {value!r} at {self.current.position}, got {self.current.value!r}."
            )

    # Purpose: Parse the complete expression and require end of input.
    # Example output: Node(kind="binary", value="+", children=(...))
    def parse(self) -> Node:
        """Parse the complete expression and require end of input."""
        # Compute and store: node.
        node = self.parse_comparison()
        # Check this condition before choosing the next branch.
        if self.current.kind != "eof":
            # Stop with a clear validation error for invalid input.
            raise FormulaParseError(
                f"Unexpected token {self.current.value!r} at {self.current.position}."
            )
        # Return the value produced by this step.
        return node

    # Purpose: Parse comparison operators with the lowest precedence.
    # Example output: Node(kind="binary", value="<=", children=(...))
    def parse_comparison(self) -> Node:
        """Parse comparison operators with the lowest precedence."""
        # Compute and store: node.
        node = self.parse_additive()
        # Check this condition before choosing the next branch.
        if self.current.kind == "compare":
            # Compute and store: operator.
            operator = self.advance().value
            # Compute and store: right.
            right = self.parse_additive()
            # Compute and store: node.
            node = Node("binary", operator, (node, right))
            # Check this condition before choosing the next branch.
            if self.current.kind == "compare":
                # Stop with a clear validation error for invalid input.
                raise FormulaParseError("Chained comparisons are not supported.")
        # Return the value produced by this step.
        return node

    # Purpose: Parse addition and subtraction nodes.
    # Example output: Node(kind="binary", value="+", children=(...))
    def parse_additive(self) -> Node:
        """Parse addition and subtraction nodes."""
        # Compute and store: node.
        node = self.parse_multiplicative()
        # Repeat while this condition remains true.
        while self.current.value in ("+", "-"):
            # Compute and store: operator.
            operator = self.advance().value
            # Compute and store: node.
            node = Node("binary", operator, (node, self.parse_multiplicative()))
        # Return the value produced by this step.
        return node

    # Purpose: Parse multiplication, division, and implicit multiplication.
    # Example output: Node(kind="binary", value="*", children=(...))
    def parse_multiplicative(self) -> Node:
        """Parse multiplication, division, and implicit multiplication."""
        # Compute and store: node.
        node = self.parse_unary()
        # Repeat while this condition remains true.
        while True:
            # Check this condition before choosing the next branch.
            if self.current.value in ("*", "/"):
                # Compute and store: operator.
                operator = self.advance().value
                # Compute and store: node.
                node = Node("binary", operator, (node, self.parse_unary()))
                # Execute this statement as part of the current workflow.
                continue
            # Excel descriptions commonly omit the multiplication sign: 4t,
            # 2PI(), or (W+H)(K+Q).  Insert it only before a primary expression.
            if self.current.kind in ("number", "cell", "name", "dimension") or self.current.value == "(":
                # Compute and store: node.
                node = Node("binary", "*", (node, self.parse_unary()))
                # Execute this statement as part of the current workflow.
                continue
            # Return the value produced by this step.
            return node

    # Purpose: Parse exponentiation.
    # Example output: Node(kind="binary", value="^", children=(...))
    def parse_power(self) -> Node:
        """Parse exponentiation."""
        # Compute and store: node.
        node = self.parse_primary()
        # Check this condition before choosing the next branch.
        if self.current.value == "^":
            # Execute this expression for its result or side effect.
            self.advance()
            # Compute and store: node.
            node = Node("binary", "^", (node, self.parse_unary()))
        # Return the value produced by this step.
        return node

    # Purpose: Parse unary plus and minus.
    # Example output: Node(kind="unary", value="-", children=(...))
    def parse_unary(self) -> Node:
        """Parse unary plus and minus."""
        # Check this condition before choosing the next branch.
        if self.current.value in ("+", "-"):
            # Compute and store: operator.
            operator = self.advance().value
            # Return the value produced by this step.
            return Node("unary", operator, (self.parse_unary(),))
        # Return the value produced by this step.
        return self.parse_power()

    # Purpose: Parse numbers, names, calls, cells, and grouped expressions.
    # Example output: Node(kind="number", value=2.0)
    def parse_primary(self) -> Node:
        """Parse numbers, names, calls, cells, and grouped expressions."""
        # Compute and store: token.
        token = self.advance()
        # Check this condition before choosing the next branch.
        if token.kind == "number":
            # Return the value produced by this step.
            return Node("number", float(token.value))
        # Check this condition before choosing the next branch.
        if token.kind == "cell":
            # Return the value produced by this step.
            return Node("cell", token.value.upper())
        # Check this condition before choosing the next branch.
        if token.kind in ("name", "dimension"):
            # Compute and store: name.
            name = token.value
            # Check this condition before choosing the next branch.
            if self.accept("("):
                # Compute and store: arguments.
                arguments: list[Node] = []
                # Check this condition before choosing the next branch.
                if not self.accept(")"):
                    # Repeat while this condition remains true.
                    while True:
                        # Execute this expression for its result or side effect.
                        arguments.append(self.parse_comparison())
                        # Check this condition before choosing the next branch.
                        if self.accept(")"):
                            # Execute this statement as part of the current workflow.
                            break
                        # Execute this expression for its result or side effect.
                        self.expect(",")
                # Return the value produced by this step.
                return Node("call", name.upper(), tuple(arguments))
            # Return the value produced by this step.
            return Node("name", name)
        # Check this condition before choosing the next branch.
        if token.value == "(":
            # Compute and store: node.
            node = self.parse_comparison()
            # Execute this expression for its result or side effect.
            self.expect(")")
            # Return the value produced by this step.
            return node
        # Stop with a clear validation error for invalid input.
        raise FormulaParseError(f"Unexpected token {token.value!r} at {token.position}.")


@lru_cache(maxsize=4096)
# Purpose: Parse one expression without its optional ``S=`` label.
# Example output: Node(kind="binary", value="+", children=(...))
def parse_formula(expression: str) -> Node:
    """Parse one expression without its optional ``S=`` label."""
    # Compute and store: text.
    text = str(expression).strip()
    # Check this condition before choosing the next branch.
    if not text:
        # Stop with a clear validation error for invalid input.
        raise FormulaParseError("Formula is empty.")
    # Return the value produced by this step.
    return _Parser(text).parse()


# Purpose: Map formula aliases such as W' and L to normalized input names.
# Example output: "W1"
def _variable_name(name: str) -> str:
    """Map formula aliases such as W' and L to normalized input names."""
    # Return the value produced by this step.
    return _VARIABLE_ALIASES.get(name.upper(), name)


# Purpose: Evaluate a parsed expression with row variables and Excel cell values.
# Example output: 1300.0
def evaluate(node: Node, variables: Mapping[str, float], cells: Mapping[str, float] | None = None) -> float | bool:
    """Evaluate a parsed expression with row variables and Excel cell values."""
    # Compute and store: cells.
    cells = cells or {}
    # Check this condition before choosing the next branch.
    if node.kind == "number":
        # Return the value produced by this step.
        return node.value
    # Check this condition before choosing the next branch.
    if node.kind == "name":
        # Compute and store: name.
        name = _variable_name(str(node.value))
        # Check this condition before choosing the next branch.
        if name.upper() == "PI":
            # Return the value produced by this step.
            return math.pi
        # Process each item in this input sequence.
        for key in (name, name.upper(), name.lower()):
            # Check this condition before choosing the next branch.
            if key in variables:
                # Return the value produced by this step.
                return float(variables[key])
        # Stop with a clear validation error for invalid input.
        raise FormulaEvaluationError(f"Unknown variable {node.value}.")
    # Check this condition before choosing the next branch.
    if node.kind == "cell":
        # Compute and store: key.
        key = str(node.value).replace("$", "").upper()
        # Check this condition before choosing the next branch.
        if key not in cells:
            # Stop with a clear validation error for invalid input.
            raise FormulaEvaluationError(f"Missing cell value {key}.")
        # Return the value produced by this step.
        return float(cells[key])
    # Check this condition before choosing the next branch.
    if node.kind == "unary":
        # Compute and store: value.
        value = evaluate(node.children[0], variables, cells)
        # Return the value produced by this step.
        return +float(value) if node.value == "+" else -float(value)
    # Check this condition before choosing the next branch.
    if node.kind == "binary":
        # Compute and store: left.
        left = evaluate(node.children[0], variables, cells)
        # Compute and store: right.
        right = evaluate(node.children[1], variables, cells)
        # Compute and store: operator.
        operator = node.value
        # Check this condition before choosing the next branch.
        if operator == "+":
            # Return the value produced by this step.
            return float(left) + float(right)
        # Check this condition before choosing the next branch.
        if operator == "-":
            # Return the value produced by this step.
            return float(left) - float(right)
        # Check this condition before choosing the next branch.
        if operator == "*":
            # Return the value produced by this step.
            return float(left) * float(right)
        # Check this condition before choosing the next branch.
        if operator == "/":
            # Check this condition before choosing the next branch.
            if float(right) == 0:
                # Stop with a clear validation error for invalid input.
                raise FormulaEvaluationError("Division by zero.")
            # Return the value produced by this step.
            return float(left) / float(right)
        # Check this condition before choosing the next branch.
        if operator == "^":
            # Return the value produced by this step.
            return float(left) ** float(right)
        # Check this condition before choosing the next branch.
        if operator == "=":
            # Return the value produced by this step.
            return float(left) == float(right)
        # Check this condition before choosing the next branch.
        if operator == "<>":
            # Return the value produced by this step.
            return float(left) != float(right)
        # Check this condition before choosing the next branch.
        if operator == "<":
            # Return the value produced by this step.
            return float(left) < float(right)
        # Check this condition before choosing the next branch.
        if operator == "<=":
            # Return the value produced by this step.
            return float(left) <= float(right)
        # Check this condition before choosing the next branch.
        if operator == ">":
            # Return the value produced by this step.
            return float(left) > float(right)
        # Check this condition before choosing the next branch.
        if operator == ">=":
            # Return the value produced by this step.
            return float(left) >= float(right)
    # Check this condition before choosing the next branch.
    if node.kind == "call":
        # Compute and store: name.
        name = str(node.value).upper()
        # Check this condition before choosing the next branch.
        if name == "IF":
            # Check this condition before choosing the next branch.
            if len(node.children) != 3:
                # Stop with a clear validation error for invalid input.
                raise FormulaEvaluationError("IF requires three arguments.")
            # Compute and store: condition.
            condition = bool(evaluate(node.children[0], variables, cells))
            # Return the value produced by this step.
            return evaluate(node.children[1 if condition else 2], variables, cells)
        # Check this condition before choosing the next branch.
        if name in ("AND", "OR"):
            # Check this condition before choosing the next branch.
            if not node.children:
                # Stop with a clear validation error for invalid input.
                raise FormulaEvaluationError(f"{name} requires at least one argument.")
            # Compute and store: values.
            values = [bool(evaluate(child, variables, cells)) for child in node.children]
            # Return the value produced by this step.
            return all(values) if name == "AND" else any(values)
        # Compute and store: values.
        values = [float(evaluate(child, variables, cells)) for child in node.children]
        # Check this condition before choosing the next branch.
        if name == "PI" and not values:
            # Return the value produced by this step.
            return math.pi
        # Check this condition before choosing the next branch.
        if name == "SQRT" and len(values) == 1:
            # Check this condition before choosing the next branch.
            if values[0] < 0:
                # Stop with a clear validation error for invalid input.
                raise FormulaEvaluationError("SQRT received a negative value.")
            # Return the value produced by this step.
            return math.sqrt(values[0])
        # Check this condition before choosing the next branch.
        if name == "ABS" and len(values) == 1:
            # Return the value produced by this step.
            return abs(values[0])
        # Check this condition before choosing the next branch.
        if name == "ROUND" and len(values) in (1, 2):
            # Return the value produced by this step.
            return round(values[0], int(values[1]) if len(values) == 2 else 0)
        # Check this condition before choosing the next branch.
        if name == "MIN" and values:
            # Return the value produced by this step.
            return min(values)
        # Check this condition before choosing the next branch.
        if name == "MAX" and values:
            # Return the value produced by this step.
            return max(values)
        # Stop with a clear validation error for invalid input.
        raise FormulaEvaluationError(f"Unsupported function {name}.")
    # Stop with a clear validation error for invalid input.
    raise FormulaEvaluationError(f"Unsupported expression node {node.kind}.")


_BINARY_PRECEDENCE = {
    "=": 1,
    "<>": 1,
    "<": 1,
    "<=": 1,
    ">": 1,
    ">=": 1,
    "+": 2,
    "-": 2,
    "*": 3,
    "/": 3,
    "^": 4,
}


# Purpose: Keep only the parentheses needed to preserve AST evaluation order.
# Example output: True
def _needs_parentheses(parent_operator: str, child: Node, *, right: bool) -> bool:
    """Keep only the parentheses needed to preserve AST evaluation order."""
    # Check this condition before choosing the next branch.
    if child.kind == "unary":
        # Return the value produced by this step.
        return parent_operator == "^"
    # Check this condition before choosing the next branch.
    if child.kind != "binary":
        # Return the value produced by this step.
        return False

    # Compute and store: child_operator.
    child_operator = str(child.value)
    # Compute and store: parent_precedence.
    parent_precedence = _BINARY_PRECEDENCE[parent_operator]
    # Compute and store: child_precedence.
    child_precedence = _BINARY_PRECEDENCE[child_operator]
    # Check this condition before choosing the next branch.
    if child_precedence < parent_precedence:
        # Return the value produced by this step.
        return True
    # Check this condition before choosing the next branch.
    if child_precedence > parent_precedence:
        # Return the value produced by this step.
        return False

    # Same-precedence expressions are left associative except exponentiation.
    # Keep the right side grouped for non-associative operators (a-(b-c),
    # a/(b*c), ...), while allowing the common a+b+c and a*b*c forms.
    if parent_operator == "^":
        # Return the value produced by this step.
        return True
    # Check this condition before choosing the next branch.
    if not right:
        # Return the value produced by this step.
        return False
    # Check this condition before choosing the next branch.
    if parent_operator == "+" and child_operator == "+":
        # Return the value produced by this step.
        return False
    # Check this condition before choosing the next branch.
    if parent_operator == "*" and child_operator == "*":
        # Return the value produced by this step.
        return False
    # Return the value produced by this step.
    return True


# Purpose: Render an AST without wrapping every binary operation in parentheses.
# Example output: "=E6+F6"
def _render_excel_node(
    node: Node,
    references: Mapping[str, str],
    *,
    parent_operator: str | None = None,
    right: bool = False,
) -> str:
    """Render an AST without wrapping every binary operation in parentheses."""
    # Check this condition before choosing the next branch.
    if node.kind == "number":
        # Compute and store: value.
        value = float(node.value)
        # Return the value produced by this step.
        return str(int(value)) if value.is_integer() else format(value, ".15g")
    # Check this condition before choosing the next branch.
    if node.kind == "cell":
        # Return the value produced by this step.
        return str(node.value).upper()
    # Check this condition before choosing the next branch.
    if node.kind == "name":
        # Compute and store: name.
        name = str(node.value)
        # Compute and store: key.
        key = _variable_name(name)
        # Check this condition before choosing the next branch.
        if key.upper() == "PI":
            # Return the value produced by this step.
            return "PI()"
        # Check this condition before choosing the next branch.
        if key in references:
            # Return the value produced by this step.
            return references[key]
        # Check this condition before choosing the next branch.
        if name.upper() in references:
            # Return the value produced by this step.
            return references[name.upper()]
        # Stop with a clear validation error for invalid input.
        raise FormulaParseError(f"No Excel reference for variable {name}.")
    # Check this condition before choosing the next branch.
    if node.kind == "unary":
        # Compute and store: child.
        child = _render_excel_node(node.children[0], references)
        # Check this condition before choosing the next branch.
        if node.children[0].kind == "binary":
            # Compute and store: child.
            child = f"({child})"
        # Compute and store: text.
        text = f"{node.value}{child}"
        # Check this condition before choosing the next branch.
        if parent_operator == "^":
            # Return the value produced by this step.
            return f"({text})"
        # Return the value produced by this step.
        return text
    # Check this condition before choosing the next branch.
    if node.kind == "binary":
        # Compute and store: operator.
        operator = str(node.value)
        # Compute and store: left.
        left = _render_excel_node(
            node.children[0],
            references,
            parent_operator=operator,
            right=False,
        )
        # Compute and store: right_text.
        right_text = _render_excel_node(
            node.children[1],
            references,
            parent_operator=operator,
            right=True,
        )
        # Compute and store: text.
        text = f"{left}{operator}{right_text}"
        # Check this condition before choosing the next branch.
        if parent_operator is not None and _needs_parentheses(
            parent_operator, node, right=right
        ):
            # Return the value produced by this step.
            return f"({text})"
        # Return the value produced by this step.
        return text
    # Check this condition before choosing the next branch.
    if node.kind == "call":
        # Return the value produced by this step.
        return f"{node.value}({','.join(_render_excel_node(child, references) for child in node.children)})"
    # Stop with a clear validation error for invalid input.
    raise FormulaParseError(f"Unsupported expression node {node.kind}.")


# Purpose: Render the AST as a readable Excel formula using row references.
# Example output: "=E6+F6"
def render_excel(
    node: Node,
    references: Mapping[str, str],
    *,
    top_level: bool = True,
) -> str:
    """Render the AST as a readable Excel formula using row references."""
    # Return the value produced by this step.
    return _render_excel_node(node, references)


@lru_cache(maxsize=16384)
# Purpose: Return a stable representation for commutative + and * blocks.
# Example output: ("binary", "+", ...)
def canonical(node: Node) -> Any:
    """Return a stable representation for commutative + and * blocks."""
    # Check this condition before choosing the next branch.
    if node.kind in ("number", "name", "cell"):
        # Compute and store: value.
        value = float(node.value) if node.kind == "number" else str(node.value).upper()
        # Check this condition before choosing the next branch.
        if node.kind == "name":
            # Compute and store: value.
            value = _variable_name(str(node.value)).upper()
        # Return the value produced by this step.
        return (node.kind, value)
    # Check this condition before choosing the next branch.
    if node.kind == "unary":
        # Return the value produced by this step.
        return ("unary", node.value, canonical(node.children[0]))
    # Check this condition before choosing the next branch.
    if node.kind == "call":
        # Return the value produced by this step.
        return ("call", str(node.value).upper(), tuple(canonical(child) for child in node.children))
    # Compute and store: children.
    children = []
    # Check this condition before choosing the next branch.
    if node.value in ("+", "*"):
        # Process each item in this input sequence.
        for child in node.children:
            # Compute and store: child_key.
            child_key = canonical(child)
            # Check this condition before choosing the next branch.
            if child.kind == "binary" and child.value == node.value:
                # Execute this expression for its result or side effect.
                children.extend(child_key[2])
            else:
                # Execute this expression for its result or side effect.
                children.append(child_key)
        # Execute this expression for its result or side effect.
        children.sort(key=repr)
        # Return the value produced by this step.
        return ("binary", node.value, tuple(children))
    # Return the value produced by this step.
    return ("binary", node.value, tuple(canonical(child) for child in node.children))


# Purpose: Parse, evaluate and optionally render one row formula.
# Example output: {"value": 1300.0, "excel_formula": "=E6+F6"}
def formula_result(
    expression: str,
    variables: Mapping[str, float],
    cells: Mapping[str, float] | None = None,
    references: Mapping[str, str] | None = None,
) -> dict[str, Any]:
    """Parse, evaluate and optionally render one row formula."""
    # Compute and store: node.
    node = parse_formula(expression)
    # Compute and store: result.
    result = evaluate(node, variables, cells)
    # Return the value produced by this step.
    return {
        "value": float(result),
        "excel_formula": f"={render_excel(node, references or {})}" if references is not None else "",
        "canonical": canonical(node),
    }


# Purpose: Use a case-insensitive key for intermediate formula variables.
# Example output: "S_V1"
def _assignment_key(label: str) -> str:
    """Use a case-insensitive key for intermediate formula variables."""
    # Return the value produced by this step.
    return str(label).strip().upper()


# Purpose: Build an evaluation scope containing row values and prior assignments.
# Example output: {"S": Node(...), "S_V1": Node(...)}
def _with_assignment_values(
    variables: Mapping[str, float],
    assignment_values: Mapping[str, float],
) -> dict[str, float]:
    """Build an evaluation scope containing row values and prior assignments."""
    # Compute and store: scope.
    scope = dict(variables)
    # Process each item in this input sequence.
    for key, value in assignment_values.items():
        # Compute and store: scope[key].
        scope[key] = value
        # Compute and store: scope[key.upper()].
        scope[key.upper()] = value
        # Compute and store: scope[key.lower()].
        scope[key.lower()] = value
    # Return the value produced by this step.
    return scope


# Purpose: Inline intermediate variables so an output cell gets one real formula.
# Example output: Node(kind="binary", value="+", children=(...))
def _expand_assignments(
    node: Node,
    definitions: Mapping[str, Node],
    stack: tuple[str, ...] = (),
) -> Node:
    """Inline intermediate variables so an output cell gets one real formula."""
    # Check this condition before choosing the next branch.
    if node.kind != "name":
        # Check this condition before choosing the next branch.
        if node.kind == "unary":
            # Return the value produced by this step.
            return Node("unary", node.value, (_expand_assignments(node.children[0], definitions, stack),))
        # Check this condition before choosing the next branch.
        if node.kind in ("binary", "call"):
            # Return the value produced by this step.
            return Node(
                node.kind,
                node.value,
                tuple(_expand_assignments(child, definitions, stack) for child in node.children),
            )
        # Return the value produced by this step.
        return node

    # Compute and store: key.
    key = _assignment_key(str(node.value))
    # Compute and store: replacement.
    replacement = definitions.get(key)
    # Check this condition before choosing the next branch.
    if replacement is None:
        # Return the value produced by this step.
        return node
    # Check this condition before choosing the next branch.
    if key in stack:
        # Stop with a clear validation error for invalid input.
        raise FormulaEvaluationError(f"Circular formula reference: {' -> '.join(stack + (key,))}.")
    # Return the value produced by this step.
    return _expand_assignments(replacement, definitions, stack + (key,))


# Purpose: Evaluate labelled formulas with intermediate variables.
# Example output: {"value": 28.0, "target": "S", "excel_formula": "=Q6*(2*E6+2*F6)"}
def formula_block_result(
    assignments: list[tuple[str, str]],
    variables: Mapping[str, float],
    cells: Mapping[str, float] | None = None,
    references: Mapping[str, str] | None = None,
    target: str | None = None,
) -> dict[str, Any]:
    """Evaluate labelled formulas with intermediate variables.

    Example input::

        [("S", "K*(S_V1+S_V2)"), ("S_V1", "2*PI()*W")]

    Definitions may appear in any order.  The evaluator repeatedly resolves
    definitions whose dependencies are available, then expands the target
    expression inline for Excel output.  A genuinely unknown variable or a
    circular dependency remains an explicit error instead of becoming zero.

    ``target`` must be one of the final output labels ``S``, ``SCC``, ``SCN``,
    ``STA`` or ``SCC/TA``.  When omitted, the only output label in the block
    is selected automatically; blocks with multiple output labels require an
    explicit target.
    """
    # Check this condition before choosing the next branch.
    if not assignments:
        # Stop with a clear validation error for invalid input.
        raise FormulaParseError("Formula block is empty.")

    # Compute and store: parsed.
    parsed: dict[str, Node] = {}
    # Compute and store: source_expressions.
    source_expressions: dict[str, str] = {}
    
    # Process each item in this input sequence.
    for label, expression in assignments:
        # Compute and store: key.
        key = _assignment_key(label)
        # Compute and store: parsed[key].
        parsed[key] = parse_formula(expression)
        # Compute and store: source_expressions[key].
        source_expressions[key] = expression

    # Compute and store: values.
    values: dict[str, float] = {}
    # Compute and store: pending.
    pending = list(parsed)
    # Compute and store: errors.
    errors: dict[str, Exception] = {}
    # Compute and store: cells.
    cells = cells or {}

    # Repeat while this condition remains true.
    while pending:
        # Compute and store: progress.
        progress = False
        # Compute and store: next_pending.
        next_pending: list[str] = []
        # Process each item in this input sequence.
        for key in pending:
            # Protect parsing/evaluation so bad workbook data becomes a controlled error.
            try:
                # Compute and store: value.
                value = evaluate(
                    parsed[key],
                    _with_assignment_values(variables, values),
                    cells,
                )
            except FormulaEvaluationError as error:
                # Unknown variables may be intermediate definitions that have
                # not been resolved yet.  Other errors are real errors now.
                if str(error).startswith("Unknown variable "):
                    # Execute this expression for its result or side effect.
                    next_pending.append(key)
                    # Compute and store: errors[key].
                    errors[key] = error
                    # Execute this statement as part of the current workflow.
                    continue
                # Stop with a clear validation error for invalid input.
                raise
            # Compute and store: values[key].
            values[key] = float(value)
            # Compute and store: progress.
            progress = True
        # Check this condition before choosing the next branch.
        if not progress:
            # Compute and store: unresolved.
            unresolved = ", ".join(next_pending)
            # Compute and store: detail.
            detail = str(errors[next_pending[0]]) if next_pending else "unknown dependency"
            # Stop with a clear validation error for invalid input.
            raise FormulaEvaluationError(f"Unresolved formula variables: {unresolved}. {detail}")
        # Compute and store: pending.
        pending = next_pending

    # Check this condition before choosing the next branch.
    if target is None:
        # Compute and store: output_targets.
        output_targets = [
            _assignment_key(label)
            for label, _expression in assignments
            if _assignment_key(label) in _OUTPUT_TARGET_LABELS
        ]
        # Check this condition before choosing the next branch.
        if len(output_targets) != 1:
            # Check this condition before choosing the next branch.
            if not output_targets:
                # Stop with a clear validation error for invalid input.
                raise FormulaEvaluationError(
                    "Formula block has no output target. "
                    "Expected one of S, SCC, SCN, STA, SCC/TA."
                )
            # Stop with a clear validation error for invalid input.
            raise FormulaEvaluationError(
                "Formula block has multiple output targets; pass target explicitly."
            )
        # Compute and store: target_key.
        target_key = output_targets[0]
    else:
        # Compute and store: target_key.
        target_key = _assignment_key(target)
        # Check this condition before choosing the next branch.
        if target_key not in _OUTPUT_TARGET_LABELS:
            # Stop with a clear validation error for invalid input.
            raise FormulaEvaluationError(
                f"Formula target {target!r} must be one of "
                "S, SCC, SCN, STA, SCC/TA."
            )
    # Check this condition before choosing the next branch.
    if target_key not in parsed:
        # Stop with a clear validation error for invalid input.
        raise FormulaEvaluationError(f"Formula target {target!r} was not defined.")
    # Compute and store: expanded_nodes.
    expanded_nodes = {
        key: _expand_assignments(node, parsed)
        for key, node in parsed.items()
    }
    # Compute and store: target_node.
    target_node = expanded_nodes[target_key]
    # Return the value produced by this step.
    return {
        "value": values[target_key],
        "excel_formula": (
            f"={render_excel(target_node, references or {})}" if references is not None else ""
        ),
        "canonical": canonical(target_node),
        "values": values,
        "expressions": source_expressions,
        "expanded_nodes": expanded_nodes,
    }


__all__ = [
    "FormulaEvaluationError",
    "FormulaParseError",
    "Node",
    "canonical",
    "evaluate",
    "extract_assignments",
    "formula_block_result",
    "formula_result",
    "parse_formula",
    "render_excel",
]
