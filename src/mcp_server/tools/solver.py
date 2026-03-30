"""Constrained optimization solver using scipy.optimize."""

from __future__ import annotations

import ast
import math
import re
from logging import Logger

from mcp_server.models.solver import SolverResult
from mcp_server.utils.excel_helpers import get_sheet, load_workbook_safe, save_workbook_safe
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

# Safe AST nodes — arithmetic, comparisons, and basic math calls
_SAFE_NODES = (
    ast.Expression,
    ast.BinOp,
    ast.UnaryOp,
    ast.Constant,
    ast.Name,
    ast.Call,
    ast.Add,
    ast.Sub,
    ast.Mult,
    ast.Div,
    ast.Pow,
    ast.Mod,
    ast.FloorDiv,
    ast.USub,
    ast.UAdd,
    ast.Load,
    ast.Compare,
    ast.BoolOp,
    ast.And,
    ast.Or,
    ast.Gt,
    ast.Lt,
    ast.GtE,
    ast.LtE,
    ast.Eq,
    ast.NotEq,
)

_SAFE_MATH_FUNCS: dict[str, object] = {
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

_CMP_SPLIT = re.compile(r"(<=|>=|==|!=|<|>)")


def _sanitize_var_name(cell_ref: str) -> str:
    """Convert a cell reference to a valid Python identifier."""
    name = re.sub(r"[^a-zA-Z0-9]", "_", cell_ref)
    if name and name[0].isdigit():
        name = "_" + name
    return name


def _validate_expr_vars(expr: str, allowed_vars: set[str]) -> ast.Expression:
    """Parse and validate an arithmetic expression, allowing only specified variable names."""
    try:
        tree = ast.parse(expr, mode="eval")
    except SyntaxError as e:
        raise ValueError(f"Invalid expression syntax: {e}") from e

    for node in ast.walk(tree):
        if not isinstance(node, _SAFE_NODES):
            raise ValueError(
                f"Unsafe expression node: {type(node).__name__}. "
                "Only arithmetic operations and basic math functions are allowed."
            )
        if isinstance(node, ast.Name) and node.id not in allowed_vars and node.id not in _SAFE_MATH_FUNCS:
            raise ValueError(
                f"Unknown variable '{node.id}'. Allowed: {sorted(allowed_vars)}, "
                f"math functions: {list(_SAFE_MATH_FUNCS)}"
            )
        if isinstance(node, ast.Call):
            if not isinstance(node.func, ast.Name) or node.func.id not in _SAFE_MATH_FUNCS:
                raise ValueError(f"Unsafe function call. Only these are allowed: {list(_SAFE_MATH_FUNCS)}")
    return tree


def _eval_expr_vars(expr: str, variables: dict[str, float]) -> float:
    """Evaluate a validated arithmetic expression with named variables."""
    tree = _validate_expr_vars(expr, set(variables.keys()))
    code = compile(tree, "<solver_expression>", "eval")
    namespace = {**variables, **_SAFE_MATH_FUNCS}
    return float(eval(code, {"__builtins__": {}}, namespace))  # noqa: S307


def _remap_vars(expr: str, var_map: dict[str, str]) -> str:
    """Replace cell refs in an expression with sanitized Python identifiers."""
    result = expr
    for cell_ref, var_name in sorted(var_map.items(), key=lambda x: -len(x[0])):
        result = result.replace(cell_ref, var_name)
    return result


def _build_constraint(
    c_dict: dict,
    var_map: dict[str, str],
    var_names_ordered: list[str],
    allowed_vars: set[str],
) -> dict:
    """Parse a constraint dict and return a scipy-compatible constraint dict."""
    expr = c_dict.get("expression", "")
    fallback_type = c_dict.get("type", "ineq")

    match = _CMP_SPLIT.search(expr)
    if not match:
        # No comparison operator — treat the whole expression as the constraint function
        expr_mapped = _remap_vars(expr, var_map)
        _validate_expr_vars(expr_mapped, allowed_vars)

        def make_plain(e: str) -> object:
            def f(x_arr: list) -> float:
                vd = {v: float(x_arr[i]) for i, v in enumerate(var_names_ordered)}
                return _eval_expr_vars(e, vd)

            return f

        return {"type": fallback_type, "fun": make_plain(expr_mapped)}

    op = match.group(0)
    lhs_raw = _remap_vars(expr[: match.start()].strip(), var_map)
    rhs_raw = _remap_vars(expr[match.end() :].strip(), var_map)

    _validate_expr_vars(lhs_raw, allowed_vars)
    _validate_expr_vars(rhs_raw, allowed_vars)

    if op in ("<=", "<"):
        # LHS <= RHS  →  RHS - LHS >= 0  (scipy ineq: fun(x) >= 0)
        def make_le(lhs: str, rhs: str) -> object:
            def f(x_arr: list) -> float:
                vd = {v: float(x_arr[i]) for i, v in enumerate(var_names_ordered)}
                return _eval_expr_vars(rhs, vd) - _eval_expr_vars(lhs, vd)

            return f

        return {"type": "ineq", "fun": make_le(lhs_raw, rhs_raw)}

    if op in (">=", ">"):
        # LHS >= RHS  →  LHS - RHS >= 0
        def make_ge(lhs: str, rhs: str) -> object:
            def f(x_arr: list) -> float:
                vd = {v: float(x_arr[i]) for i, v in enumerate(var_names_ordered)}
                return _eval_expr_vars(lhs, vd) - _eval_expr_vars(rhs, vd)

            return f

        return {"type": "ineq", "fun": make_ge(lhs_raw, rhs_raw)}

    # == or !=  →  equality constraint: LHS - RHS = 0
    def make_eq(lhs: str, rhs: str) -> object:
        def f(x_arr: list) -> float:
            vd = {v: float(x_arr[i]) for i, v in enumerate(var_names_ordered)}
            return _eval_expr_vars(lhs, vd) - _eval_expr_vars(rhs, vd)

        return f

    return {"type": "eq", "fun": make_eq(lhs_raw, rhs_raw)}


def run_solver(
    file_path: str,
    sheet_name: str,
    objective_expression: str,
    variable_cells: dict[str, tuple[float, float]],
    constraints: list[dict[str, str]] | None = None,
    maximize: bool = False,
    tolerance: float = 1e-6,
    max_iterations: int = 1000,
) -> SolverResult:
    """Multi-variable constrained optimization using scipy SLSQP.

    objective_expression: arithmetic expression using cell refs as variable names
        e.g. "B2 * B3 - B4"
    variable_cells: {cell_ref: (lower_bound, upper_bound)}
        e.g. {"B2": (0.0, 100.0), "B3": (0.0, 50.0)}
    constraints: list of {"expression": "B2 + B3 <= 100", "type": "ineq"|"eq"}
    maximize: negate the objective to maximise instead of minimise
    """
    from scipy.optimize import minimize

    if not variable_cells:
        raise ValueError("variable_cells must not be empty.")

    # Build sanitized variable name mapping
    var_map: dict[str, str] = {ref: _sanitize_var_name(ref) for ref in variable_cells}
    cell_refs = list(variable_cells.keys())
    var_names_ordered = [var_map[ref] for ref in cell_refs]
    allowed_vars: set[str] = set(var_names_ordered)

    # Validate and remap objective expression
    obj_expr = _remap_vars(objective_expression, var_map)
    _validate_expr_vars(obj_expr, allowed_vars)

    # Build scipy constraints
    scipy_constraints = [_build_constraint(c, var_map, var_names_ordered, allowed_vars) for c in (constraints or [])]

    # Read initial values from workbook
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        x0 = []
        for ref in cell_refs:
            lo, hi = variable_cells[ref]
            v = ws[ref].value
            if v is not None:
                try:
                    x0.append(float(v))
                    continue
                except (TypeError, ValueError):
                    pass
            x0.append((lo + hi) / 2.0)
    finally:
        wb.close()

    bounds = [variable_cells[ref] for ref in cell_refs]

    def objective(x_arr: list) -> float:
        vd = {v: float(x_arr[i]) for i, v in enumerate(var_names_ordered)}
        val = _eval_expr_vars(obj_expr, vd)
        return -val if maximize else val

    result = minimize(
        objective,
        x0,
        method="SLSQP",
        bounds=bounds,
        constraints=scipy_constraints,
        tol=tolerance,
        options={"maxiter": max_iterations},
    )

    found_values = {ref: round(float(result.x[i]), 6) for i, ref in enumerate(cell_refs)}
    obj_val = round(float(-result.fun if maximize else result.fun), 6)

    # Write found values back to workbook
    wb2 = load_workbook_safe(file_path)
    try:
        ws2 = get_sheet(wb2, sheet_name)
        for ref, val in found_values.items():
            ws2[ref] = val
        save_workbook_safe(wb2, file_path)
    finally:
        wb2.close()

    logger.info(
        "Solver: converged=%s, objective=%.6f, iterations=%d",
        result.success,
        obj_val,
        result.nit,
    )
    return {
        "found_values": found_values,
        "objective_value": obj_val,
        "converged": bool(result.success),
        "iterations": int(result.nit),
        "message": result.message,
    }
