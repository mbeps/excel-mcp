"""Financial calculations: NPV, IRR, PMT, goal seek, loan amortization, and advanced analytics."""

from __future__ import annotations

import ast
import math
from logging import Logger

import numpy as np
import numpy_financial as npf

from mcp_server.utils.excel_helpers import (
    col_letter_to_index,
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
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


def _validate_expression_vars(expr: str, allowed_vars: set[str]) -> ast.Expression:
    """Parse and validate a math expression, allowing specified variable names."""
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
        if isinstance(node, ast.Name) and node.id not in allowed_vars and node.id not in _SAFE_MATH_FUNCS:
            raise ValueError(
                f"Unknown variable '{node.id}'. Allowed variables: {sorted(allowed_vars)},"
                f" math functions: {list(_SAFE_MATH_FUNCS)}"
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


def _eval_expression_vars(expr: str, variables: dict[str, float]) -> float:
    """Evaluate a validated expression with multiple named variables."""
    tree = _validate_expression_vars(expr, set(variables.keys()))
    code = compile(tree, "<expression>", "eval")
    namespace = {**variables, **_SAFE_MATH_FUNCS}
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
    max_periods: int | None = None,
) -> dict:
    """Generate a loan amortization schedule."""
    periodic_rate = annual_rate / payments_per_year
    total_periods = years * payments_per_year
    payment = float(-npf.pmt(periodic_rate, total_periods, principal))

    schedule = []
    balance = principal
    total_interest = 0.0

    periods_to_show = total_periods if max_periods is None else min(max_periods, total_periods)
    for period in range(1, total_periods + 1):
        interest = balance * periodic_rate
        principal_paid = payment - interest
        balance -= principal_paid
        total_interest += interest

        if period <= periods_to_show:
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
    result: dict = {
        "monthly_payment": round(payment, 2),
        "total_interest": round(total_interest, 2),
        "total_paid": round(total_paid, 2),
        "total_periods": total_periods,
        "periods_shown": periods_to_show,
        "schedule": schedule,
    }
    if max_periods is not None and max_periods < total_periods:
        result["truncated"] = True
    return result


def dcf_analysis(
    cash_flows: list[float],
    discount_rate: float,
    terminal_growth_rate: float = 0.02,
    initial_investment: float = 0.0,
) -> dict:
    """Discounted Cash Flow valuation with terminal value (Gordon Growth Model)."""
    if discount_rate <= terminal_growth_rate:
        raise ValueError("Discount rate must be greater than terminal growth rate.")
    if not cash_flows:
        raise ValueError("At least one cash flow is required.")

    pv_cash_flows = [cf / (1 + discount_rate) ** (i + 1) for i, cf in enumerate(cash_flows)]
    total_pv = sum(pv_cash_flows)

    last_cf = cash_flows[-1]
    terminal_value = (last_cf * (1 + terminal_growth_rate)) / (discount_rate - terminal_growth_rate)
    n = len(cash_flows)
    pv_terminal_value = terminal_value / (1 + discount_rate) ** n

    enterprise_value = total_pv + pv_terminal_value
    net_value = enterprise_value - initial_investment

    logger.info("DCF analysis: enterprise_value=%.2f, net_value=%.2f", enterprise_value, net_value)
    return {
        "pv_cash_flows": [round(pv, 2) for pv in pv_cash_flows],
        "total_pv": round(total_pv, 2),
        "terminal_value": round(terminal_value, 2),
        "pv_terminal_value": round(pv_terminal_value, 2),
        "enterprise_value": round(enterprise_value, 2),
        "net_value": round(net_value, 2),
    }


def budget_variance_analysis(
    file_path: str,
    sheet_name: str = "Sheet1",
    category_column: str = "A",
    budget_column: str = "B",
    actual_column: str = "C",
    header_row: int = 1,
    output_file: str | None = None,
) -> dict:
    """Analyze budget vs actual spending from an Excel file."""
    wb = load_workbook_safe(file_path, data_only=True)
    ws = get_sheet(wb, sheet_name)

    cat_idx = col_letter_to_index(category_column)
    bud_idx = col_letter_to_index(budget_column)
    act_idx = col_letter_to_index(actual_column)

    items: list[dict] = []
    for row in range(header_row + 1, ws.max_row + 1):
        cat_val = ws.cell(row=row, column=cat_idx).value
        bud_val = ws.cell(row=row, column=bud_idx).value
        act_val = ws.cell(row=row, column=act_idx).value

        if cat_val is None or bud_val is None or act_val is None:
            continue

        budget = float(bud_val)
        actual = float(act_val)
        variance = actual - budget
        variance_pct = (variance / budget * 100) if budget != 0 else 0.0

        if abs(variance_pct) < 0.5:
            status = "on_budget"
        elif variance > 0:
            status = "over_budget"
        else:
            status = "under_budget"

        items.append(
            {
                "category": str(cat_val),
                "budget": round(budget, 2),
                "actual": round(actual, 2),
                "variance": round(variance, 2),
                "variance_pct": round(variance_pct, 2),
                "status": status,
            }
        )

    wb.close()

    total_budget = sum(i["budget"] for i in items)
    total_actual = sum(i["actual"] for i in items)
    total_variance = total_actual - total_budget
    total_variance_pct = (total_variance / total_budget * 100) if total_budget != 0 else 0.0

    result = {
        "items": items,
        "summary": {
            "total_budget": round(total_budget, 2),
            "total_actual": round(total_actual, 2),
            "total_variance": round(total_variance, 2),
            "total_variance_pct": round(total_variance_pct, 2),
        },
    }

    if output_file:
        import openpyxl

        out_wb = openpyxl.Workbook()
        out_ws = out_wb.active
        out_ws.title = "Variance Analysis"
        headers = ["Category", "Budget", "Actual", "Variance", "Variance %", "Status"]
        out_ws.append(headers)
        for item in items:
            out_ws.append(
                [
                    item["category"],
                    item["budget"],
                    item["actual"],
                    item["variance"],
                    item["variance_pct"],
                    item["status"],
                ]
            )
        out_ws.append([])
        out_ws.append(["TOTAL", total_budget, total_actual, total_variance, total_variance_pct])
        save_workbook_safe(out_wb, output_file)
        result["output_file"] = output_file

    logger.info("Budget variance analysis: %d categories processed", len(items))
    return result


def financial_ratio_analysis(
    ratios: dict,
    industry_benchmarks: dict | None = None,
) -> dict:
    """Compute financial ratios from provided data and optionally compare to benchmarks."""
    computed: dict[str, dict] = {}

    ratio_defs: list[tuple[str, str, str, str]] = [
        ("current_ratio", "current_assets", "current_liabilities", "divide"),
        ("debt_to_equity", "total_debt", "total_equity", "divide"),
        ("roe", "net_income", "total_equity", "divide"),
        ("roa", "net_income", "total_assets", "divide"),
        ("gross_margin", "gross_profit", "revenue", "divide"),
        ("net_margin", "net_income", "revenue", "divide"),
        ("interest_coverage", "ebitda", "interest_expense", "divide"),
    ]

    for ratio_name, numerator_key, denominator_key, _ in ratio_defs:
        if numerator_key in ratios and denominator_key in ratios:
            denominator = float(ratios[denominator_key])
            if denominator == 0:
                computed[ratio_name] = {"value": None, "error": "Division by zero"}
                continue
            value = round(float(ratios[numerator_key]) / denominator, 4)
            entry: dict = {"value": value}

            if industry_benchmarks and ratio_name in industry_benchmarks:
                benchmark = float(industry_benchmarks[ratio_name])
                entry["benchmark"] = benchmark
                if abs(value - benchmark) / max(abs(benchmark), 1e-9) < 0.05:
                    entry["comparison"] = "at_benchmark"
                elif value > benchmark:
                    entry["comparison"] = "above_benchmark"
                else:
                    entry["comparison"] = "below_benchmark"

            computed[ratio_name] = entry

    logger.info("Financial ratio analysis: %d ratios computed", len(computed))
    return {"ratios": computed}


def scenario_analysis(
    base_case: dict,
    scenarios: list[dict],
    formula: str,
    periods: int = 1,
) -> dict:
    """Evaluate a formula across base case and multiple scenarios.

    Uses AST-validated safe expression evaluation (same pattern as goal_seek).
    """
    all_vars: set[str] = set(base_case.keys())
    for sc in scenarios:
        all_vars.update(sc.get("adjustments", {}).keys())
    _validate_expression_vars(formula, all_vars)

    def _evaluate_scenario(variables: dict[str, float], num_periods: int) -> list[float]:
        results = []
        current = dict(variables)
        for _ in range(num_periods):
            value = _eval_expression_vars(formula, current)
            results.append(round(value, 4))
        return results

    base_values = {k: float(v) for k, v in base_case.items()}
    base_results = _evaluate_scenario(base_values, periods)

    scenario_results = []
    for sc in scenarios:
        name = sc.get("name", "Unnamed")
        adjustments = sc.get("adjustments", {})
        sc_vars = {**base_values, **{k: float(v) for k, v in adjustments.items()}}
        sc_results = _evaluate_scenario(sc_vars, periods)
        scenario_results.append(
            {
                "name": name,
                "variables": sc_vars,
                "results": sc_results,
            }
        )

    logger.info("Scenario analysis: %d scenarios evaluated over %d periods", len(scenarios), periods)
    return {
        "base_case": {"variables": base_values, "results": base_results},
        "scenarios": scenario_results,
        "formula": formula,
        "periods": periods,
    }


def trend_analysis(
    file_path: str,
    sheet_name: str = "Sheet1",
    date_column: str = "A",
    value_column: str = "B",
    header_row: int = 1,
    periods_to_forecast: int = 3,
) -> dict:
    """Analyze trends in time-series data from an Excel file with linear regression forecasting."""
    wb = load_workbook_safe(file_path, data_only=True)
    ws = get_sheet(wb, sheet_name)

    date_idx = col_letter_to_index(date_column)
    val_idx = col_letter_to_index(value_column)

    dates: list[str] = []
    values: list[float] = []

    for row in range(header_row + 1, ws.max_row + 1):
        d = ws.cell(row=row, column=date_idx).value
        v = ws.cell(row=row, column=val_idx).value
        if d is None or v is None:
            continue
        dates.append(str(d))
        values.append(float(v))

    wb.close()

    if len(values) < 2:
        raise ValueError("At least 2 data points are required for trend analysis.")

    x = np.arange(len(values), dtype=float)
    y = np.array(values, dtype=float)

    coeffs = np.polyfit(x, y, 1)
    slope = float(coeffs[0])
    intercept = float(coeffs[1])

    y_pred = np.polyval(coeffs, x)
    ss_res = float(np.sum((y - y_pred) ** 2))
    ss_tot = float(np.sum((y - np.mean(y)) ** 2))
    r_squared = 1 - (ss_res / ss_tot) if ss_tot != 0 else 0.0

    window = max(2, min(3, len(values) // 3))
    moving_averages: list[float | None] = [None] * (window - 1)
    for i in range(window - 1, len(values)):
        ma = sum(values[i - window + 1 : i + 1]) / window
        moving_averages.append(round(ma, 4))

    growth_rates: list[float | None] = [None]
    for i in range(1, len(values)):
        if values[i - 1] != 0:
            gr = (values[i] - values[i - 1]) / abs(values[i - 1])
            growth_rates.append(round(gr, 4))
        else:
            growth_rates.append(None)

    forecast: list[dict] = []
    for i in range(1, periods_to_forecast + 1):
        fx = len(values) - 1 + i
        fv = float(np.polyval(coeffs, fx))
        forecast.append({"period": len(values) + i, "value": round(fv, 4)})

    if abs(slope) < 1e-9:
        trend_direction = "stable"
    elif slope > 0:
        trend_direction = "increasing"
    else:
        trend_direction = "decreasing"

    logger.info("Trend analysis: direction=%s, slope=%.4f, r²=%.4f", trend_direction, slope, r_squared)
    return {
        "trend_direction": trend_direction,
        "slope": round(slope, 4),
        "intercept": round(intercept, 4),
        "r_squared": round(r_squared, 4),
        "data_points": len(values),
        "dates": dates,
        "values": values,
        "moving_averages": moving_averages,
        "growth_rates": growth_rates,
        "forecast": forecast,
    }


def calculate_xnpv(
    discount_rate: float,
    cash_flows: list[float],
    dates: list[str],
) -> dict:
    """Calculate XNPV (NPV with irregular cash flow dates).

    discount_rate: annual discount rate (e.g. 0.10 for 10%)
    cash_flows: list of cash flow amounts
    dates: list of ISO-format date strings ("YYYY-MM-DD")
    """
    from datetime import date as _date

    if len(cash_flows) != len(dates):
        raise ValueError("cash_flows and dates must have the same length")
    if discount_rate <= -1:
        raise ValueError("discount_rate must be greater than -1")

    parsed = [_date.fromisoformat(d) for d in dates]
    d0 = parsed[0]
    xnpv = sum(cf / (1 + discount_rate) ** ((d - d0).days / 365) for cf, d in zip(cash_flows, parsed))
    logger.info("XNPV: %.4f (rate=%.4f, n=%d)", xnpv, discount_rate, len(cash_flows))
    return {
        "xnpv": round(xnpv, 4),
        "discount_rate": discount_rate,
        "num_cash_flows": len(cash_flows),
    }


def calculate_xirr(
    cash_flows: list[float],
    dates: list[str],
    guess: float = 0.1,
) -> dict:
    """Calculate XIRR (IRR for irregular cash flow dates).

    cash_flows: list of cash flow amounts (must have at least one positive and one negative)
    dates: list of ISO-format date strings ("YYYY-MM-DD")
    guess: initial rate guess
    """
    from datetime import date as _date

    from scipy.optimize import brentq, fsolve

    if len(cash_flows) != len(dates):
        raise ValueError("cash_flows and dates must have the same length")

    parsed = [_date.fromisoformat(d) for d in dates]
    d0 = parsed[0]

    def _xnpv_objective(r: float) -> float:
        return sum(cf / (1 + r) ** ((d - d0).days / 365) for cf, d in zip(cash_flows, parsed))

    xirr_value: float | None = None
    try:
        xirr_value = brentq(_xnpv_objective, -0.999, 100.0, maxiter=1000)
    except Exception:
        try:
            result = fsolve(_xnpv_objective, guess, full_output=True)
            sol = float(result[0][0])
            if abs(_xnpv_objective(sol)) < 1e-6:
                xirr_value = sol
        except Exception:
            pass

    if xirr_value is None:
        return {"xirr": None, "message": "XIRR could not be computed: convergence failed"}

    logger.info("XIRR: %.6f (n=%d)", xirr_value, len(cash_flows))
    return {
        "xirr": round(xirr_value, 6),
        "dates": dates,
        "cash_flows": cash_flows,
    }


def calculate_cagr(
    beginning_value: float,
    ending_value: float,
    periods: float,
) -> dict:
    """Calculate Compound Annual Growth Rate (CAGR).

    beginning_value: starting value (must be > 0)
    ending_value: ending value
    periods: number of periods (years)
    """
    if beginning_value <= 0:
        raise ValueError("beginning_value must be greater than 0")
    if periods <= 0:
        raise ValueError("periods must be greater than 0")

    cagr = (ending_value / beginning_value) ** (1 / periods) - 1
    total_growth = (ending_value - beginning_value) / beginning_value
    logger.info("CAGR: %.6f over %.2f periods", cagr, periods)
    return {
        "cagr": round(cagr, 6),
        "beginning_value": beginning_value,
        "ending_value": ending_value,
        "periods": periods,
        "total_growth": round(total_growth, 4),
    }


def break_even_analysis(
    fixed_costs: float,
    price_per_unit: float,
    variable_cost_per_unit: float,
) -> dict:
    """Calculate break-even point in units and revenue.

    fixed_costs: total fixed costs
    price_per_unit: selling price per unit
    variable_cost_per_unit: variable cost per unit
    """
    if fixed_costs < 0:
        raise ValueError("fixed_costs must be non-negative")
    if price_per_unit <= 0:
        raise ValueError("price_per_unit must be positive")
    if variable_cost_per_unit < 0:
        raise ValueError("variable_cost_per_unit must be non-negative")
    if price_per_unit <= variable_cost_per_unit:
        raise ValueError("price_per_unit must be greater than variable_cost_per_unit")

    contribution_margin = price_per_unit - variable_cost_per_unit
    contribution_margin_ratio = contribution_margin / price_per_unit
    break_even_units = fixed_costs / contribution_margin
    break_even_revenue = break_even_units * price_per_unit

    logger.info(
        "Break-even: units=%.2f, revenue=%.2f, contribution_margin=%.4f",
        break_even_units,
        break_even_revenue,
        contribution_margin,
    )
    return {
        "break_even_units": round(break_even_units, 2),
        "break_even_revenue": round(break_even_revenue, 2),
        "contribution_margin": contribution_margin,
        "contribution_margin_ratio": round(contribution_margin_ratio, 4),
    }
