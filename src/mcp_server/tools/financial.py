"""Financial calculations: NPV, IRR, PMT, goal seek, loan amortization, and advanced analytics."""

from __future__ import annotations

import os
from logging import Logger

import numpy_financial as npf

from mcp_server.utils.excel_helpers import (
    col_letter_to_index,
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.expression_validator import SAFE_MATH_FUNCS, validate_expression
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def _eval_expression(expr: str, x_val: float) -> float:
    tree = validate_expression(expr, allowed_names=frozenset({"x"}))
    code = compile(tree, "<expression>", "eval")
    namespace = {"x": x_val, **SAFE_MATH_FUNCS}
    return float(eval(code, {"__builtins__": {}}, namespace))


def _eval_expression_vars(expr: str, variables: dict[str, float]) -> float:
    """Evaluate a validated expression with multiple named variables."""
    tree = validate_expression(expr, allowed_names=frozenset(variables.keys()))
    code = compile(tree, "<expression>", "eval")
    namespace = {**variables, **SAFE_MATH_FUNCS}
    return float(eval(code, {"__builtins__": {}}, namespace))


def goal_seek(
    file_path: str,
    sheet_name: str,
    variable_cell: str,
    expression: str,
    target_value: float,
    initial_value: float = 0.0,
    tolerance: float = 1e-6,
    max_iterations: int = 1000,
) -> dict:
    """Find x such that expression(x) = target_value and write the result to variable_cell.

    expression is a math expression in terms of 'x', e.g. '1000 * (1 + x)**10'.
    Only basic arithmetic and basic math functions (sqrt, log, exp, sin, cos, etc.) are allowed.
    """
    from scipy.optimize import root_scalar

    validate_expression(expression, allowed_names=frozenset({"x"}))

    def objective(x_val: float) -> float:
        return _eval_expression(expression, x_val) - target_value

    x1 = initial_value * 1.1 if initial_value != 0 else 0.1
    try:
        result = root_scalar(
            objective,
            x0=initial_value,
            x1=x1,
            method="secant",
            maxiter=max_iterations,
            xtol=tolerance,
        )
        x_found = result.root
        converged = bool(result.converged)
        iterations = int(result.iterations)
    except Exception as e:
        raise ValueError(f"Goal seek failed to converge: {e}") from e

    achieved = _eval_expression(expression, x_found)

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws[variable_cell] = float(x_found)
        save_workbook_safe(wb, file_path)
    finally:
        wb.close()

    logger.info("Goal seek: x=%.6f, achieved=%.6f, target=%.6f", x_found, achieved, target_value)
    return {
        "variable_cell": variable_cell,
        "found_value": float(x_found),
        "achieved_result": float(achieved),
        "target_value": target_value,
        "converged": converged,
        "iterations": iterations,
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
    try:
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
    finally:
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

        if os.path.realpath(output_file) == os.path.realpath(file_path):
            out_wb = load_workbook_safe(file_path)
            if "Variance Analysis" in out_wb.sheetnames:
                del out_wb["Variance Analysis"]
            out_ws = out_wb.create_sheet(title="Variance Analysis")
        else:
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
        out_wb.close()

    logger.info("Budget variance analysis: %d categories processed", len(items))
    return result


def financial_ratio_analysis(
    financial_data: dict,
    industry_benchmarks: dict | None = None,
) -> dict:
    """Compute financial ratios from raw financial metric values and optionally compare to benchmarks.

    ``financial_data`` must be a flat dict of financial metric values keyed by the component names
    listed below.  Do NOT pass computed ratio names (e.g. ``current_ratio``) — those are outputs,
    not inputs.

    Valid input keys:
        current_assets       – total current assets
        current_liabilities  – total current liabilities
        total_debt           – total debt (for debt-to-equity)
        total_equity         – shareholders' equity  (also called ``equity``)
        net_income           – net income / net profit
        total_assets         – total assets
        revenue              – total revenue / net sales
        gross_profit         – gross profit (revenue minus COGS)
        operating_income     – operating income / EBIT
        ebitda               – EBITDA (used as numerator for interest-coverage)
        interest_expense     – interest expense

    Computed ratios (returned when both component keys are present):
        current_ratio        = current_assets / current_liabilities
        debt_to_equity       = total_debt / total_equity
        roe                  = net_income / total_equity
        roa                  = net_income / total_assets
        gross_margin         = gross_profit / revenue
        net_margin           = net_income / revenue
        interest_coverage    = ebitda / interest_expense
    """
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
        if numerator_key in financial_data and denominator_key in financial_data:
            denominator = float(financial_data[denominator_key])
            if denominator == 0:
                computed[ratio_name] = {"value": None, "error": "Division by zero"}
                continue
            value = round(float(financial_data[numerator_key]) / denominator, 4)
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


def calculate_fv(
    rate: float,
    nper: int,
    pmt: float,
    pv: float = 0.0,
    when: str = "end",
) -> dict:
    """Calculate future value of an annuity or lump sum.

    rate: periodic interest rate (e.g. 0.05 for 5%)
    nper: number of periods
    pmt: payment per period (negative = outflow)
    pv:  present value (negative = outflow)
    when: 'end' (ordinary annuity) or 'begin' (annuity due)
    """
    when_int = 0 if when == "end" else 1
    result = float(npf.fv(rate, nper, pmt, pv, when=when_int))
    logger.info("calculate_fv: rate=%.4f nper=%d pmt=%.2f pv=%.2f -> fv=%.4f", rate, nper, pmt, pv, result)
    return {"rate": rate, "nper": nper, "pmt": pmt, "pv": pv, "when": when, "fv": round(result, 4)}


def calculate_pv(
    rate: float,
    nper: int,
    pmt: float,
    fv: float = 0.0,
    when: str = "end",
) -> dict:
    """Calculate present value of an annuity or lump sum.

    rate: periodic interest rate
    nper: number of periods
    pmt: payment per period
    fv:  future value
    when: 'end' or 'begin'
    """
    when_int = 0 if when == "end" else 1
    result = float(npf.pv(rate, nper, pmt, fv, when=when_int))
    logger.info("calculate_pv: rate=%.4f nper=%d pmt=%.2f fv=%.2f -> pv=%.4f", rate, nper, pmt, fv, result)
    return {"rate": rate, "nper": nper, "pmt": pmt, "fv": fv, "when": when, "pv": round(result, 4)}


def calculate_nper(
    rate: float,
    pmt: float,
    pv: float,
    fv: float = 0.0,
    when: str = "end",
) -> dict:
    """Calculate number of periods for an annuity.

    rate: periodic interest rate
    pmt: payment per period
    pv:  present value
    fv:  future value
    when: 'end' or 'begin'
    """
    when_int = 0 if when == "end" else 1
    result = float(npf.nper(rate, pmt, pv, fv, when=when_int))
    logger.info("calculate_nper: rate=%.4f pmt=%.2f pv=%.2f fv=%.2f -> nper=%.4f", rate, pmt, pv, fv, result)
    return {"rate": rate, "pmt": pmt, "pv": pv, "fv": fv, "when": when, "nper": round(result, 4)}


def calculate_rate(
    nper: int,
    pmt: float,
    pv: float,
    fv: float = 0.0,
    when: str = "end",
    guess: float = 0.1,
) -> dict:
    """Calculate the periodic interest rate for an annuity.

    nper: number of periods
    pmt: payment per period
    pv:  present value
    fv:  future value
    when: 'end' or 'begin'
    guess: starting guess for iteration
    """
    when_int = 0 if when == "end" else 1
    result = float(npf.rate(nper, pmt, pv, fv, when=when_int, guess=guess))
    if math.isnan(result):
        raise ValueError(
            "Rate calculation did not converge. Try a different 'guess' value "
            "(e.g., guess=0.05 for low rates or guess=0.2 for high rates)."
        )
    # round() can produce IEEE-754 -0.0 when result is a tiny negative value
    # (e.g., -4e-9 rounds to -0.0 at 6 d.p.).  Adding 0.0 canonicalizes to 0.0.
    rate_rounded = round(result, 6) + 0.0
    logger.info("calculate_rate: nper=%d pmt=%.2f pv=%.2f fv=%.2f -> rate=%.6f", nper, pmt, pv, fv, rate_rounded)
    return {"nper": nper, "pmt": pmt, "pv": pv, "fv": fv, "when": when, "rate": rate_rounded}


_DEPRECIATION_METHOD_ALIASES: dict[str, str] = {
    "straight_line": "sln",
    "sum_of_years": "syd",
    "sum_of_years_digits": "syd",
    "double_declining": "ddb",
    "double_declining_balance": "ddb",
}


def calculate_depreciation(
    cost: float,
    salvage: float,
    life: int,
    method: str = "sln",
    period: int | None = None,
) -> dict:
    """Calculate depreciation using SLN, SYD, or DDB method.

    cost:    initial asset cost
    salvage: residual value at end of life
    life:    useful life in periods
    method:  'sln' / 'straight_line', 'syd' / 'sum_of_years' / 'sum_of_years_digits',
             'ddb' / 'double_declining' / 'double_declining_balance'
    period:  required for 'syd' and 'ddb' (1-based period number)
    """
    method = _DEPRECIATION_METHOD_ALIASES.get(method.lower(), method.lower())
    if method == "sln":
        result = (cost - salvage) / life
        logger.info("depreciation SLN: %.2f", result)
        return {
            "method": method,
            "cost": cost,
            "salvage": salvage,
            "life": life,
            "period": None,
            "depreciation": round(result, 2),
        }
    elif method in ("syd", "ddb"):
        if period is None:
            raise ValueError(f"'period' is required for method='{method}'")
        if period < 1 or period > life:
            raise ValueError(f"period must be between 1 and {life}")
        if method == "syd":
            syd_sum = life * (life + 1) / 2
            result = (cost - salvage) * (life - period + 1) / syd_sum
        else:
            ddb_rate = 2.0 / life
            accumulated = 0.0
            result = 0.0
            for p in range(1, period + 1):
                dep = ddb_rate * (cost - accumulated)
                if accumulated + dep > cost - salvage:
                    dep = max(cost - salvage - accumulated, 0.0)
                accumulated += dep
                if p == period:
                    result = dep
                    break
        logger.info("depreciation %s period %d: %.2f", method.upper(), period, result)
        return {
            "method": method,
            "cost": cost,
            "salvage": salvage,
            "life": life,
            "period": period,
            "depreciation": round(result, 2),
        }
    else:
        raise ValueError(f"Unknown method '{method}'. Valid values: 'sln', 'syd', 'ddb'")


def calculate_irr(cash_flows: list[float]) -> dict:
    """Calculate Internal Rate of Return for a series of cash flows.

    cash_flows: list with at least one negative value (initial investment)
    and subsequent returns. E.g. [-1000, 300, 400, 500].
    Returns the periodic IRR as a decimal (e.g. 0.15 for 15%).
    """
    if len(cash_flows) < 2:
        raise ValueError("cash_flows must contain at least 2 values.")
    if not any(v < 0 for v in cash_flows):
        raise ValueError("cash_flows must contain at least one negative value (initial investment).")
    if not any(v > 0 for v in cash_flows):
        raise ValueError("cash_flows must contain at least one positive value (return).")

    result = npf.irr(cash_flows)
    if result is None or (isinstance(result, float) and (result != result)):  # NaN check
        raise ValueError("IRR could not converge for the given cash flows.")

    logger.info("IRR calculated: %.6f for %d cash flows", result, len(cash_flows))
    return {
        "irr": round(float(result), 6),
        "irr_percent": round(float(result) * 100, 4),
        "cash_flows": cash_flows,
    }


def create_sensitivity_table(
    file_path: str,
    sheet_name: str,
    output_cell: str,
    expression: str,
    var1_name: str,
    var1_values: list[float],
    var2_name: str | None = None,
    var2_values: list[float] | None = None,
) -> dict:
    """Create a one- or two-variable sensitivity table in the sheet.

    Evaluates `expression` over a grid of variable values and writes the
    results starting at `output_cell`.

    expression: arithmetic expression using var1_name (and var2_name).
                Only arithmetic ops and math functions are allowed.
    var1_values: list of values for the row variable.
    var2_values: optional list of values for the column variable (2-var table).
    output_cell: top-left cell where the table will be written.
    """
    from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

    # Validate expression
    all_vars = {var1_name}
    if var2_name:
        all_vars.add(var2_name)
    validate_expression(expression, allowed_names=frozenset(all_vars))

    # Compute results
    table: list[list[float | str]] = []
    if var2_name and var2_values:
        for v2 in var2_values:
            row: list[float | str] = []
            for v1 in var1_values:
                val = _eval_expression_vars(expression, {var1_name: v1, var2_name: v2})
                row.append(round(val, 4))
            table.append(row)
    else:
        for v1 in var1_values:
            val = _eval_expression_vars(expression, {var1_name: v1})
            single_row: list[float | str] = [round(val, 4)]
            table.append(single_row)

    # Write to workbook
    col_str, start_row = coordinate_from_string(output_cell)
    start_col = column_index_from_string(col_str)

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        # Header row: label in corner, then var1 values across columns
        ws.cell(row=start_row, column=start_col, value=f"{var1_name} →")
        for ci, v1 in enumerate(var1_values):
            ws.cell(row=start_row, column=start_col + 1 + ci, value=v1)
        # Data rows
        for ri, result_row in enumerate(table):
            label: float | str = var2_values[ri] if (var2_values and var2_name) else ""
            ws.cell(row=start_row + 1 + ri, column=start_col, value=label)
            for ci, cell_val in enumerate(result_row):
                ws.cell(row=start_row + 1 + ri, column=start_col + 1 + ci, value=cell_val)
        save_workbook_safe(wb, file_path)
    finally:
        wb.close()

    logger.info(
        "create_sensitivity_table: %s (%d×%d) written at %s in %s",
        expression,
        len(table),
        len(var1_values),
        output_cell,
        file_path,
    )
    return {
        "expression": expression,
        "var1": var1_name,
        "var1_values": var1_values,
        "var2": var2_name,
        "var2_values": var2_values,
        "table": table,
        "output_cell": output_cell,
        "file_path": file_path,
    }
