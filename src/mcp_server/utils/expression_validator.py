"""Shared AST-based expression validator for safe math evaluation."""

from __future__ import annotations

import ast
import math

_DEFAULT_FORBIDDEN_NAMES: frozenset[str] = frozenset(
    {
        "__builtins__",
        "__import__",
        "exec",
        "eval",
        "open",
        "system",
        "getattr",
        "setattr",
        "delattr",
        "globals",
        "locals",
        "compile",
        "breakpoint",
        "input",
        "print",
        "exit",
        "quit",
    }
)

_SAFE_NODES = (
    ast.Expression,
    ast.BinOp,
    ast.UnaryOp,
    ast.Constant,
    ast.Name,
    ast.Call,
    ast.Attribute,
    ast.Compare,
    ast.BoolOp,
    ast.IfExp,
    ast.Subscript,
    ast.Index,
    ast.Load,
    ast.Store,
    ast.Tuple,
    ast.List,
    ast.Add,
    ast.Sub,
    ast.Mult,
    ast.Div,
    ast.Pow,
    ast.Mod,
    ast.FloorDiv,
    ast.USub,
    ast.UAdd,
    ast.And,
    ast.Or,
    ast.Gt,
    ast.Lt,
    ast.GtE,
    ast.LtE,
    ast.Eq,
    ast.NotEq,
)

SAFE_MATH_FUNCS: dict[str, object] = {
    "sqrt": math.sqrt,
    "log": math.log,
    "log10": math.log10,
    "exp": math.exp,
    "sin": math.sin,
    "cos": math.cos,
    "tan": math.tan,
    "abs": abs,
    "pow": pow,
}


def validate_expression(
    expression: str,
    *,
    allowed_names: frozenset[str] | None = None,
    forbidden_names: frozenset[str] | None = None,
) -> ast.Expression:
    """Validate a mathematical expression via AST inspection.

    Raises ValueError if the expression contains unsafe constructs.
    """
    if forbidden_names is None:
        forbidden_names = _DEFAULT_FORBIDDEN_NAMES

    try:
        tree = ast.parse(expression, mode="eval")
    except SyntaxError as e:
        raise ValueError(f"Invalid expression syntax: {e}") from e

    for node in ast.walk(tree):
        if not isinstance(node, _SAFE_NODES):
            raise ValueError(
                f"Unsafe expression node: {type(node).__name__}. "
                "Only arithmetic operations and basic math functions are allowed."
            )
        if isinstance(node, ast.Name):
            if node.id in forbidden_names:
                raise ValueError(f"Expression references forbidden name '{node.id}'.")
            if allowed_names is not None and node.id not in allowed_names and node.id not in SAFE_MATH_FUNCS:
                raise ValueError(
                    f"Unknown variable '{node.id}'. Allowed: {sorted(allowed_names)}, "
                    f"math functions: {list(SAFE_MATH_FUNCS)}"
                )
        if isinstance(node, ast.Call):
            if not isinstance(node.func, ast.Name) or node.func.id not in SAFE_MATH_FUNCS:
                raise ValueError(f"Unsafe function call. Only these are allowed: {list(SAFE_MATH_FUNCS)}")

    return tree
