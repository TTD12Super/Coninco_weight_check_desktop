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


def _normalize_source(text: str) -> str:
    source = unicodedata.normalize("NFKC", str(text))
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

    source = re.sub(
        r"(?<![A-Za-z0-9_,(])(\d+),(\d+)(?!\s*,)",
        r"\1.\2",
        source,
    )


    return source.replace(";", ",")


def extract_assignments(text: Any) -> list[tuple[str, str]]:
    """Extract a formula block's labelled expressions.

    Besides output labels such as ``SCN`` and ``STA``, this recognizes
    intermediate labels such as ``S_V1``.  The latter are kept in the returned
    block so the evaluator can resolve dependencies before calculating ``S``.
    """
    

    ####### Kiem tra xem text dau vao co bi rong hoac khong phai chuoi hay khong, Neu rong thi tra ve List rong
    if text is None or (not isinstance(text, str) and pd.isna(text)):
        return []
    
    #### Chuan hoa chuoi de dung voi cong thuc trong excel
    source = _normalize_source(str(text))

    #### Tim cac cong thuc dien giai
    matches = list(_ASSIGNMENT_RE.finditer(source))

    
    matches.extend(_INTERMEDIATE_ASSIGNMENT_RE.finditer(source))

    matches.sort(key=lambda match: match.start())
    if not matches:
        expression = source.strip()
        return [("S", expression)] if expression else []

    assignments: list[tuple[str, str]] = []
    
    for position, match in enumerate(matches):
        start = match.end()
        end = matches[position + 1].start() if position + 1 < len(matches) else len(source)
        expression = source[start:end].strip().strip(";,")
        if expression:
            assignments.append((match.group("label").upper(), expression))
    return assignments


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
    missing_openings = source.count(")") - source.count("(")
    if missing_openings > 0:
        return "(" * missing_openings + source
    return source


def _tokens(expression: str) -> list[Token]:
    source = _repair_extra_closing_parentheses(_normalize_source(expression))
    result: list[Token] = []
    position = 0
    while position < len(source):
        match = _TOKEN_RE.match(source, position)
        if not match:
            raise FormulaParseError(f"Unexpected character {source[position]!r} at {position}.")
        kind = next(name for name, value in match.groupdict().items() if value is not None)
        result.append(Token(kind, match.group(kind), position))
        position = match.end()
    result.append(Token("eof", "", len(source)))
    return result


class _Parser:
    def __init__(self, expression: str):
        self.expression = expression
        self.tokens = _tokens(expression)
        self.index = 0

    @property
    def current(self) -> Token:
        return self.tokens[self.index]

    def advance(self) -> Token:
        token = self.current
        self.index += 1
        return token

    def accept(self, value: str) -> bool:
        if self.current.value.upper() == value.upper():
            self.advance()
            return True
        return False

    def expect(self, value: str) -> None:
        if not self.accept(value):
            raise FormulaParseError(
                f"Expected {value!r} at {self.current.position}, got {self.current.value!r}."
            )

    def parse(self) -> Node:
        node = self.parse_comparison()
        if self.current.kind != "eof":
            raise FormulaParseError(
                f"Unexpected token {self.current.value!r} at {self.current.position}."
            )
        return node

    def parse_comparison(self) -> Node:
        node = self.parse_additive()
        if self.current.kind == "compare":
            operator = self.advance().value
            right = self.parse_additive()
            node = Node("binary", operator, (node, right))
            if self.current.kind == "compare":
                raise FormulaParseError("Chained comparisons are not supported.")
        return node

    def parse_additive(self) -> Node:
        node = self.parse_multiplicative()
        while self.current.value in ("+", "-"):
            operator = self.advance().value
            node = Node("binary", operator, (node, self.parse_multiplicative()))
        return node

    def parse_multiplicative(self) -> Node:
        node = self.parse_unary()
        while True:
            if self.current.value in ("*", "/"):
                operator = self.advance().value
                node = Node("binary", operator, (node, self.parse_unary()))
                continue
            # Excel descriptions commonly omit the multiplication sign: 4t,
            # 2PI(), or (W+H)(K+Q).  Insert it only before a primary expression.
            if self.current.kind in ("number", "cell", "name", "dimension") or self.current.value == "(":
                node = Node("binary", "*", (node, self.parse_unary()))
                continue
            return node

    def parse_power(self) -> Node:
        node = self.parse_primary()
        if self.current.value == "^":
            self.advance()
            node = Node("binary", "^", (node, self.parse_unary()))
        return node

    def parse_unary(self) -> Node:
        if self.current.value in ("+", "-"):
            operator = self.advance().value
            return Node("unary", operator, (self.parse_unary(),))
        return self.parse_power()

    def parse_primary(self) -> Node:
        token = self.advance()
        if token.kind == "number":
            return Node("number", float(token.value))
        if token.kind == "cell":
            return Node("cell", token.value.upper())
        if token.kind in ("name", "dimension"):
            name = token.value
            if self.accept("("):
                arguments: list[Node] = []
                if not self.accept(")"):
                    while True:
                        arguments.append(self.parse_comparison())
                        if self.accept(")"):
                            break
                        self.expect(",")
                return Node("call", name.upper(), tuple(arguments))
            return Node("name", name)
        if token.value == "(":
            node = self.parse_comparison()
            self.expect(")")
            return node
        raise FormulaParseError(f"Unexpected token {token.value!r} at {token.position}.")


def parse_formula(expression: str) -> Node:
    """Parse one expression without its optional ``S=`` label."""
    text = str(expression).strip()
    if not text:
        raise FormulaParseError("Formula is empty.")
    return _Parser(text).parse()


def _variable_name(name: str) -> str:
    return _VARIABLE_ALIASES.get(name.upper(), name)


def evaluate(node: Node, variables: Mapping[str, float], cells: Mapping[str, float] | None = None) -> float | bool:
    """Evaluate a parsed expression with row variables and Excel cell values."""
    cells = cells or {}
    if node.kind == "number":
        return node.value
    if node.kind == "name":
        name = _variable_name(str(node.value))
        if name.upper() == "PI":
            return math.pi
        for key in (name, name.upper(), name.lower()):
            if key in variables:
                return float(variables[key])
        raise FormulaEvaluationError(f"Unknown variable {node.value}.")
    if node.kind == "cell":
        key = str(node.value).replace("$", "").upper()
        if key not in cells:
            raise FormulaEvaluationError(f"Missing cell value {key}.")
        return float(cells[key])
    if node.kind == "unary":
        value = evaluate(node.children[0], variables, cells)
        return +float(value) if node.value == "+" else -float(value)
    if node.kind == "binary":
        left = evaluate(node.children[0], variables, cells)
        right = evaluate(node.children[1], variables, cells)
        operator = node.value
        if operator == "+":
            return float(left) + float(right)
        if operator == "-":
            return float(left) - float(right)
        if operator == "*":
            return float(left) * float(right)
        if operator == "/":
            if float(right) == 0:
                raise FormulaEvaluationError("Division by zero.")
            return float(left) / float(right)
        if operator == "^":
            return float(left) ** float(right)
        if operator == "=":
            return float(left) == float(right)
        if operator == "<>":
            return float(left) != float(right)
        if operator == "<":
            return float(left) < float(right)
        if operator == "<=":
            return float(left) <= float(right)
        if operator == ">":
            return float(left) > float(right)
        if operator == ">=":
            return float(left) >= float(right)
    if node.kind == "call":
        name = str(node.value).upper()
        if name == "IF":
            if len(node.children) != 3:
                raise FormulaEvaluationError("IF requires three arguments.")
            condition = bool(evaluate(node.children[0], variables, cells))
            return evaluate(node.children[1 if condition else 2], variables, cells)
        if name in ("AND", "OR"):
            if not node.children:
                raise FormulaEvaluationError(f"{name} requires at least one argument.")
            values = [bool(evaluate(child, variables, cells)) for child in node.children]
            return all(values) if name == "AND" else any(values)
        values = [float(evaluate(child, variables, cells)) for child in node.children]
        if name == "PI" and not values:
            return math.pi
        if name == "SQRT" and len(values) == 1:
            if values[0] < 0:
                raise FormulaEvaluationError("SQRT received a negative value.")
            return math.sqrt(values[0])
        if name == "ABS" and len(values) == 1:
            return abs(values[0])
        if name == "ROUND" and len(values) in (1, 2):
            return round(values[0], int(values[1]) if len(values) == 2 else 0)
        if name == "MIN" and values:
            return min(values)
        if name == "MAX" and values:
            return max(values)
        raise FormulaEvaluationError(f"Unsupported function {name}.")
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


def _needs_parentheses(parent_operator: str, child: Node, *, right: bool) -> bool:
    """Keep only the parentheses needed to preserve AST evaluation order."""
    if child.kind == "unary":
        return parent_operator == "^"
    if child.kind != "binary":
        return False

    child_operator = str(child.value)
    parent_precedence = _BINARY_PRECEDENCE[parent_operator]
    child_precedence = _BINARY_PRECEDENCE[child_operator]
    if child_precedence < parent_precedence:
        return True
    if child_precedence > parent_precedence:
        return False

    # Same-precedence expressions are left associative except exponentiation.
    # Keep the right side grouped for non-associative operators (a-(b-c),
    # a/(b*c), ...), while allowing the common a+b+c and a*b*c forms.
    if parent_operator == "^":
        return True
    if not right:
        return False
    if parent_operator == "+" and child_operator == "+":
        return False
    if parent_operator == "*" and child_operator == "*":
        return False
    return True


def _render_excel_node(
    node: Node,
    references: Mapping[str, str],
    *,
    parent_operator: str | None = None,
    right: bool = False,
) -> str:
    """Render an AST without wrapping every binary operation in parentheses."""
    if node.kind == "number":
        value = float(node.value)
        return str(int(value)) if value.is_integer() else format(value, ".15g")
    if node.kind == "cell":
        return str(node.value).upper()
    if node.kind == "name":
        name = str(node.value)
        key = _variable_name(name)
        if key.upper() == "PI":
            return "PI()"
        if key in references:
            return references[key]
        if name.upper() in references:
            return references[name.upper()]
        raise FormulaParseError(f"No Excel reference for variable {name}.")
    if node.kind == "unary":
        child = _render_excel_node(node.children[0], references)
        if node.children[0].kind == "binary":
            child = f"({child})"
        text = f"{node.value}{child}"
        if parent_operator == "^":
            return f"({text})"
        return text
    if node.kind == "binary":
        operator = str(node.value)
        left = _render_excel_node(
            node.children[0],
            references,
            parent_operator=operator,
            right=False,
        )
        right_text = _render_excel_node(
            node.children[1],
            references,
            parent_operator=operator,
            right=True,
        )
        text = f"{left}{operator}{right_text}"
        if parent_operator is not None and _needs_parentheses(
            parent_operator, node, right=right
        ):
            return f"({text})"
        return text
    if node.kind == "call":
        return f"{node.value}({','.join(_render_excel_node(child, references) for child in node.children)})"
    raise FormulaParseError(f"Unsupported expression node {node.kind}.")


def render_excel(
    node: Node,
    references: Mapping[str, str],
    *,
    top_level: bool = True,
) -> str:
    """Render the AST as a readable Excel formula using row references."""
    return _render_excel_node(node, references)


def canonical(node: Node) -> Any:
    """Return a stable representation for commutative + and * blocks."""
    if node.kind in ("number", "name", "cell"):
        value = float(node.value) if node.kind == "number" else str(node.value).upper()
        if node.kind == "name":
            value = _variable_name(str(node.value)).upper()
        return (node.kind, value)
    if node.kind == "unary":
        return ("unary", node.value, canonical(node.children[0]))
    if node.kind == "call":
        return ("call", str(node.value).upper(), tuple(canonical(child) for child in node.children))
    children = []
    if node.value in ("+", "*"):
        for child in node.children:
            child_key = canonical(child)
            if child.kind == "binary" and child.value == node.value:
                children.extend(child_key[2])
            else:
                children.append(child_key)
        children.sort(key=repr)
        return ("binary", node.value, tuple(children))
    return ("binary", node.value, tuple(canonical(child) for child in node.children))


def formula_result(
    expression: str,
    variables: Mapping[str, float],
    cells: Mapping[str, float] | None = None,
    references: Mapping[str, str] | None = None,
) -> dict[str, Any]:
    """Parse, evaluate and optionally render one row formula."""
    node = parse_formula(expression)
    result = evaluate(node, variables, cells)
    return {
        "value": float(result),
        "excel_formula": f"={render_excel(node, references or {})}" if references is not None else "",
        "canonical": canonical(node),
    }


def _assignment_key(label: str) -> str:
    """Use a case-insensitive key for intermediate formula variables."""
    return str(label).strip().upper()


def _with_assignment_values(
    variables: Mapping[str, float],
    assignment_values: Mapping[str, float],
) -> dict[str, float]:
    """Build an evaluation scope containing row values and prior assignments."""
    scope = dict(variables)
    for key, value in assignment_values.items():
        scope[key] = value
        scope[key.upper()] = value
        scope[key.lower()] = value
    return scope


def _expand_assignments(
    node: Node,
    definitions: Mapping[str, Node],
    stack: tuple[str, ...] = (),
) -> Node:
    """Inline intermediate variables so an output cell gets one real formula."""
    if node.kind != "name":
        if node.kind == "unary":
            return Node("unary", node.value, (_expand_assignments(node.children[0], definitions, stack),))
        if node.kind in ("binary", "call"):
            return Node(
                node.kind,
                node.value,
                tuple(_expand_assignments(child, definitions, stack) for child in node.children),
            )
        return node

    key = _assignment_key(str(node.value))
    replacement = definitions.get(key)
    if replacement is None:
        return node
    if key in stack:
        raise FormulaEvaluationError(f"Circular formula reference: {' -> '.join(stack + (key,))}.")
    return _expand_assignments(replacement, definitions, stack + (key,))


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
    if not assignments:
        raise FormulaParseError("Formula block is empty.")

    parsed: dict[str, Node] = {}
    source_expressions: dict[str, str] = {}
    
    for label, expression in assignments:
        key = _assignment_key(label)
        parsed[key] = parse_formula(expression)
        source_expressions[key] = expression

    values: dict[str, float] = {}
    pending = list(parsed)
    errors: dict[str, Exception] = {}
    cells = cells or {}

    while pending:
        progress = False
        next_pending: list[str] = []
        for key in pending:
            try:
                value = evaluate(
                    parsed[key],
                    _with_assignment_values(variables, values),
                    cells,
                )
            except FormulaEvaluationError as error:
                # Unknown variables may be intermediate definitions that have
                # not been resolved yet.  Other errors are real errors now.
                if str(error).startswith("Unknown variable "):
                    next_pending.append(key)
                    errors[key] = error
                    continue
                raise
            values[key] = float(value)
            progress = True
        if not progress:
            unresolved = ", ".join(next_pending)
            detail = str(errors[next_pending[0]]) if next_pending else "unknown dependency"
            raise FormulaEvaluationError(f"Unresolved formula variables: {unresolved}. {detail}")
        pending = next_pending

    if target is None:
        output_targets = [
            _assignment_key(label)
            for label, _expression in assignments
            if _assignment_key(label) in _OUTPUT_TARGET_LABELS
        ]
        if len(output_targets) != 1:
            if not output_targets:
                raise FormulaEvaluationError(
                    "Formula block has no output target. "
                    "Expected one of S, SCC, SCN, STA, SCC/TA."
                )
            raise FormulaEvaluationError(
                "Formula block has multiple output targets; pass target explicitly."
            )
        target_key = output_targets[0]
    else:
        target_key = _assignment_key(target)
        if target_key not in _OUTPUT_TARGET_LABELS:
            raise FormulaEvaluationError(
                f"Formula target {target!r} must be one of "
                "S, SCC, SCN, STA, SCC/TA."
            )
    if target_key not in parsed:
        raise FormulaEvaluationError(f"Formula target {target!r} was not defined.")
    expanded_nodes = {
        key: _expand_assignments(node, parsed)
        for key, node in parsed.items()
    }
    target_node = expanded_nodes[target_key]
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
