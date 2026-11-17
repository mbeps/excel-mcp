"""Constrained optimization solver using scipy.optimize."""

from __future__ import annotations

import math
import re
from logging import Logger
from types import CodeType

from openpyxl.utils.cell import coordinate_from_string

from mcp_server.models.solver import SolverResult
from mcp_server.utils.excel_helpers import get_sheet, load_workbook_safe, save_workbook_safe
from mcp_server.utils.expression_validator import SAFE_MATH_FUNCS, validate_expression
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

_CMP_SPLIT = re.compile(r"(<=|>=|==|!=|<|>)")


def _sanitize_var_name(cell_ref: str) -> str:
    """Convert a cell reference to a valid Python identifier."""
    name = re.sub(r"[^a-zA-Z0-9]", "_", cell_ref)
    if name and name[0].isdigit():
        name = "_" + name
    return name


def _compile_expr(expr: str, allowed_vars: set[str]) -> CodeType:
    """Parse, validate, and compile an arithmetic expression once."""
    tree = validate_expression(expr, allowed_names=frozenset(allowed_vars | set(SAFE_MATH_FUNCS.keys())))
    return compile(tree, "<solver_expression>", "eval")


def _eval_compiled(code: CodeType, variables: dict[str, float]) -> float:
    """Evaluate a pre-compiled arithmetic expression with named variables."""
    namespace = {**SAFE_MATH_FUNCS, **variables}
    return float(eval(code, {"__builtins__": {}}, namespace))  # noqa: S307


def _remap_vars(expr: str, var_map: dict[str, str]) -> str:
    """Replace cell refs in an expression with sanitized Python identifiers on token boundaries."""
    if not var_map:
        return expr
    pattern = re.compile(
        r"(?<![A-Za-z0-9_])(?:"
        + "|".join(re.escape(k) for k in sorted(var_map, key=len, reverse=True))
        + r")(?![A-Za-z0-9_])"
    )
    return pattern.sub(lambda m: var_map[m.group(0)], expr)


def _build_constraint(
    c_dict: dict,
    var_map: dict[str, str],
    var_names_ordered: list[str],
    allowed_vars: set[str],
) -> dict:
    """Parse a constraint dict and return a scipy-compatible constraint dict."""
    expr = c_dict.get("expression", "")
    fallback_type = c_dict.get("type", "ineq")
    if fallback_type not in ("ineq", "eq"):
        raise ValueError(f"Unsupported constraint type '{fallback_type}'. Must be 'ineq' or 'eq'.")

    matches = list(_CMP_SPLIT.finditer(expr))
    if len(matches) > 1:
        raise ValueError(f"Chained comparisons are not supported in constraint '{expr}'.")

    if not matches:
        # No comparison operator — treat the whole expression as the constraint function
        expr_mapped = _remap_vars(expr, var_map)
        code = _compile_expr(expr_mapped, allowed_vars)

        def make_plain(compiled: CodeType) -> object:
            def f(x_arr: list) -> float:
                vd = {v: float(x_arr[i]) for i, v in enumerate(var_names_ordered)}
                return _eval_compiled(compiled, vd)

            return f

        return {"type": fallback_type, "fun": make_plain(code)}

    match = matches[0]
    op = match.group(0)
    if op == "!=":
        raise ValueError("Constraint operator '!=' is not supported by SLSQP; use <=, >=, or ==.")

    lhs_raw = _remap_vars(expr[: match.start()].strip(), var_map)
    rhs_raw = _remap_vars(expr[match.end() :].strip(), var_map)

    lhs_code = _compile_expr(lhs_raw, allowed_vars)
    rhs_code = _compile_expr(rhs_raw, allowed_vars)

    if op in ("<=", "<"):
        # LHS <= RHS  →  RHS - LHS >= 0  (scipy ineq: fun(x) >= 0)
        def make_le(lhs_c: CodeType, rhs_c: CodeType) -> object:
            def f(x_arr: list) -> float:
                vd = {v: float(x_arr[i]) for i, v in enumerate(var_names_ordered)}
                return _eval_compiled(rhs_c, vd) - _eval_compiled(lhs_c, vd)

            return f

        return {"type": "ineq", "fun": make_le(lhs_code, rhs_code)}

    if op in (">=", ">"):
        # LHS >= RHS  →  LHS - RHS >= 0
        def make_ge(lhs_c: CodeType, rhs_c: CodeType) -> object:
            def f(x_arr: list) -> float:
                vd = {v: float(x_arr[i]) for i, v in enumerate(var_names_ordered)}
                return _eval_compiled(lhs_c, vd) - _eval_compiled(rhs_c, vd)

            return f

        return {"type": "ineq", "fun": make_ge(lhs_code, rhs_code)}

    # ==  →  equality constraint: LHS - RHS = 0
    def make_eq(lhs_c: CodeType, rhs_c: CodeType) -> object:
        def f(x_arr: list) -> float:
            vd = {v: float(x_arr[i]) for i, v in enumerate(var_names_ordered)}
            return _eval_compiled(lhs_c, vd) - _eval_compiled(rhs_c, vd)

        return f

    return {"type": "eq", "fun": make_eq(lhs_code, rhs_code)}


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

    for ref, (lo, hi) in variable_cells.items():
        if lo > hi:
            raise ValueError(f"Invalid bounds for '{ref}': lower bound {lo} > upper bound {hi}.")

    # Build sanitized variable name mapping
    var_map: dict[str, str] = {ref: _sanitize_var_name(ref) for ref in variable_cells}
    cell_refs = list(variable_cells.keys())
    var_names_ordered = [var_map[ref] for ref in cell_refs]
    allowed_vars: set[str] = set(var_names_ordered)

    # Validate, remap, and pre-compile objective expression once
    obj_expr = _remap_vars(objective_expression, var_map)
    obj_code = _compile_expr(obj_expr, allowed_vars)

    # Build scipy constraints
    scipy_constraints = [_build_constraint(c, var_map, var_names_ordered, allowed_vars) for c in (constraints or [])]

    # Validate that all variable_cells keys are single cell references (e.g. "B2"), not columns/rows
    for ref in cell_refs:
        try:
            coordinate_from_string(ref)
        except Exception as exc:
            raise ValueError(
                f"Invalid cell reference '{ref}' in variable_cells. "
                f"Keys must be single cell references like 'A1' or 'B2', not column/row names."
            ) from exc

    # Read initial values from workbook and clamp into [lo, hi]
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        x0 = []
        for ref in cell_refs:
            lo, hi = variable_cells[ref]
            v = ws[ref].value
            if v is not None and not isinstance(v, bool):
                try:
                    val_f = float(v)
                    if math.isfinite(val_f):
                        x0.append(max(lo, min(hi, val_f)))
                        continue
                except (TypeError, ValueError):
                    pass
            x0.append((lo + hi) / 2.0)
    finally:
        wb.close()

    bounds = [variable_cells[ref] for ref in cell_refs]

    def objective(x_arr: list) -> float:
        vd = {v: float(x_arr[i]) for i, v in enumerate(var_names_ordered)}
        val = _eval_compiled(obj_code, vd)
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

    converged = (
        bool(result.success) and all(math.isfinite(float(v)) for v in result.x) and math.isfinite(float(result.fun))
    )
    found_values = {
        ref: round(float(result.x[i]), 6) if math.isfinite(float(result.x[i])) else 0.0
        for i, ref in enumerate(cell_refs)
    }
    obj_val = round(float(-result.fun if maximize else result.fun), 6) if math.isfinite(float(result.fun)) else 0.0

    # Write found values back to workbook only when converged with finite results
    if converged:
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
        converged,
        obj_val,
        result.nit,
    )
    return {
        "found_values": found_values,
        "objective_value": obj_val,
        "converged": converged,
        "iterations": int(result.nit),
        "message": result.message,
    }
