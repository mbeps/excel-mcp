"""Financial calculations: NPV, IRR, PMT, goal seek, and loan amortization."""

from __future__ import annotations

import ast
import math
from logging import Logger

import numpy_financial as npf

from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

# Safe AST nodes for goal_seek expression parsing
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
)

_SAFE_MATH_FUNCS = {
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


def _validate_expression(expr: str) -> ast.Expression:
    """Parse and validate a math expression, allowing only safe operations."""
    try:
        tree = ast.parse(expr, mode="eval")
    except SyntaxError as e:
        raise ValueError(f"Invalid expression syntax: {e}") from e

    for node in ast.walk(tree):
        if not isinstance(node, _SAFE_NODES):
            raise ValueError(
                f"Unsafe expression node: {type(node).__name__}."
                " Only arithmetic operations and basic math functions are allowed."
            )
        if isinstance(node, ast.Name) and node.id != "x" and node.id not in _SAFE_MATH_FUNCS:
            raise ValueError(
                f"Unknown variable '{node.id}'. Only 'x' and math functions ({list(_SAFE_MATH_FUNCS)}) are allowed."
            )
        if isinstance(node, ast.Call):
            if not isinstance(node.func, ast.Name) or node.func.id not in _SAFE_MATH_FUNCS:
                raise ValueError(f"Unsafe function call. Only these are allowed: {list(_SAFE_MATH_FUNCS)}")
    return tree


def _eval_expression(expr: str, x_val: float) -> float:
    tree = _validate_expression(expr)
    code = compile(tree, "<expression>", "eval")
    namespace = {"x": x_val, **_SAFE_MATH_FUNCS}
    return float(eval(code, {"__builtins__": {}}, namespace))


def calculate_npv(discount_rate: float, cash_flows: list[float]) -> dict:
    """Calculate Net Present Value."""
    result = npf.npv(discount_rate, cash_flows)
    logger.info("Calculated NPV: %.2f", result)
    return {
        "npv": float(result),
        "discount_rate": discount_rate,
        "cash_flows": cash_flows,
    }


def calculate_irr(cash_flows: list[float]) -> dict:
    """Calculate Internal Rate of Return."""
    result = npf.irr(cash_flows)
    if result is None or (hasattr(result, "__float__") and math.isnan(float(result))):
        return {"irr": None, "message": "IRR could not be computed for the given cash flows."}
    logger.info("Calculated IRR: %.4f", result)
    return {"irr": float(result)}


def calculate_pmt(rate: float, nper: int, pv: float, fv: float = 0) -> dict:
    """Calculate periodic payment for a loan or annuity."""
    result = npf.pmt(rate, nper, pv, fv)
    logger.info("Calculated PMT: %.2f", result)
    return {
        "payment": float(result),
        "rate": rate,
        "nper": nper,
        "pv": pv,
        "fv": fv,
    }


def goal_seek(
    target_formula: str,
    target_value: float,
    initial_guess: float = 1.0,
) -> dict:
    """Find x such that target_formula(x) = target_value using numerical optimization.

    target_formula is a math expression in terms of 'x', e.g. '1000 * (1 + x)**10'.
    Only basic arithmetic and math functions (sqrt, log, exp, sin, cos, etc.) are allowed.
    """
    from scipy.optimize import minimize_scalar, root_scalar

    _validate_expression(target_formula)

    def objective(x_val: float) -> float:
        return _eval_expression(target_formula, x_val) - target_value

    try:
        x1 = initial_guess * 1.1 if initial_guess != 0 else 0.1
        result = root_scalar(objective, x0=initial_guess, x1=x1, method="secant", maxiter=1000)
        x_found = result.root
        converged = result.converged
    except Exception:
        try:
            result = minimize_scalar(lambda x: objective(x) ** 2)
            x_found = result.x
            converged = abs(objective(x_found)) < 1e-6
        except Exception as e:
            raise ValueError(f"Goal seek failed to converge: {e}") from e

    achieved = _eval_expression(target_formula, x_found)
    logger.info("Goal seek: x=%.6f, achieved=%.6f, target=%.6f", x_found, achieved, target_value)
    return {
        "result": float(x_found),
        "target_value": target_value,
        "achieved_value": float(achieved),
        "converged": bool(converged),
    }


def loan_amortization(
    principal: float,
    annual_rate: float,
    years: int,
    payments_per_year: int = 12,
) -> dict:
    """Generate a loan amortization schedule."""
    periodic_rate = annual_rate / payments_per_year
    total_periods = years * payments_per_year
    payment = float(-npf.pmt(periodic_rate, total_periods, principal))

    schedule = []
    balance = principal
    total_interest = 0.0

    max_periods = min(total_periods, 24)
    for period in range(1, total_periods + 1):
        interest = balance * periodic_rate
        principal_paid = payment - interest
        balance -= principal_paid
        total_interest += interest

        if period <= max_periods:
            schedule.append(
                {
                    "period": period,
                    "payment": round(payment, 2),
                    "principal": round(principal_paid, 2),
                    "interest": round(interest, 2),
                    "balance": round(max(balance, 0), 2),
                }
            )

    total_paid = payment * total_periods
    logger.info(
        "Loan amortization: principal=%.2f, payment=%.2f, total_interest=%.2f",
        principal,
        payment,
        total_interest,
    )
    return {
        "monthly_payment": round(payment, 2),
        "total_interest": round(total_interest, 2),
        "total_paid": round(total_paid, 2),
        "total_periods": total_periods,
        "periods_shown": max_periods,
        "schedule": schedule,
    }
