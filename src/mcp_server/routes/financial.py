from __future__ import annotations

from typing import Literal

from mcp.types import ToolAnnotations

import mcp_server.tools.financial as _financial
from mcp_server.models.solver import SolverResult

__all__ = [
    "goal_seek",
    "loan_amortization",
    "dcf_analysis",
    "budget_variance_analysis",
    "financial_ratio_analysis",
    "break_even_analysis",
    "create_sensitivity_table",
    "time_value_calc",
]


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
    """Find the variable_cell value that makes expression equal target_value, then write it to the workbook."""
    return _financial.goal_seek(
        file_path,
        sheet_name,
        variable_cell,
        expression,
        target_value,
        initial_value,
        tolerance,
        max_iterations,
    )


def loan_amortization(
    principal: float,
    annual_rate: float,
    years: int,
    payments_per_year: int = 12,
) -> dict:
    """Generate a loan amortization schedule with payment breakdown."""
    return _financial.loan_amortization(principal, annual_rate, years, payments_per_year)


def dcf_analysis(
    cash_flows: list[float],
    discount_rate: float,
    terminal_growth_rate: float = 0.02,
    initial_investment: float = 0.0,
) -> dict:
    """Discounted Cash Flow valuation with Gordon Growth Model terminal value."""
    return _financial.dcf_analysis(cash_flows, discount_rate, terminal_growth_rate, initial_investment)


def budget_variance_analysis(
    file_path: str,
    sheet_name: str = "Sheet1",
    category_column: str = "A",
    budget_column: str = "B",
    actual_column: str = "C",
    header_row: int = 1,
    output_file: str | None = None,
) -> dict:
    """Analyze budget vs actual spending. Returns variance per category with status."""
    return _financial.budget_variance_analysis(
        file_path,
        sheet_name,
        category_column,
        budget_column,
        actual_column,
        header_row,
        output_file,
    )


def financial_ratio_analysis(financial_data: dict, industry_benchmarks: dict | None = None) -> dict:
    """Compute financial ratios from raw financial metric values with optional benchmark comparison.

    ``financial_data`` is a dict of raw financial metric values (NOT computed ratio names).
    Valid keys: current_assets, current_liabilities, total_debt, total_equity, net_income,
    total_assets, revenue, gross_profit, ebitda, interest_expense.

    Example: {"current_assets": 500000, "current_liabilities": 250000, "net_income": 100000,
              "revenue": 1000000, "total_equity": 400000, "gross_profit": 600000}
    """
    return _financial.financial_ratio_analysis(financial_data, industry_benchmarks)


def break_even_analysis(fixed_costs: float, price_per_unit: float, variable_cost_per_unit: float) -> dict:
    """Calculate break-even point in units and revenue."""
    return _financial.break_even_analysis(fixed_costs, price_per_unit, variable_cost_per_unit)


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
    """Create a one- or two-variable sensitivity/what-if table in the workbook."""
    return _financial.create_sensitivity_table(
        file_path,
        sheet_name,
        output_cell,
        expression,
        var1_name,
        var1_values,
        var2_name,
        var2_values,
    )


def time_value_calc(
    operation: Literal["fv", "pv", "nper", "rate", "depreciation", "irr"],
    rate: float | None = None,
    nper: int | None = None,
    pmt: float | None = None,
    pv: float = 0.0,
    fv: float = 0.0,
    when: str = "end",
    guess: float = 0.1,
    cost: float | None = None,
    salvage: float | None = None,
    life: int | None = None,
    method: str = "sln",
    period: int | None = None,
    cash_flows: list[float] | None = None,
) -> dict:
    """Time value of money and depreciation calculations.

    operation="fv": Future value. Requires: rate, nper, pmt. Optional: pv, when.
    operation="pv": Present value. Requires: rate, nper, pmt. Optional: fv, when.
    operation="nper": Number of periods. Requires: rate, pmt, pv. Optional: fv, when.
    operation="rate": Interest rate. Requires: nper, pmt, pv. Optional: fv, when, guess.
    operation="depreciation": Asset depreciation. Requires: cost, salvage, life. Optional: method, period.
      method values: "sln" / "straight_line", "syd" / "sum_of_years" / "sum_of_years_digits",
                     "ddb" / "double_declining" / "double_declining_balance". Default: "sln".
    operation="irr": Internal Rate of Return. Requires: cash_flows (list of floats,
      first value typically negative as initial investment). Returns irr and irr_percent.
    """
    if operation == "fv":
        if rate is None:
            raise ValueError("rate is required for operation='fv'.")
        if nper is None:
            raise ValueError("nper is required for operation='fv'.")
        if pmt is None:
            raise ValueError("pmt is required for operation='fv'.")
        return _financial.calculate_fv(rate, nper, pmt, pv, when)
    if operation == "pv":
        if rate is None:
            raise ValueError("rate is required for operation='pv'.")
        if nper is None:
            raise ValueError("nper is required for operation='pv'.")
        if pmt is None:
            raise ValueError("pmt is required for operation='pv'.")
        return _financial.calculate_pv(rate, nper, pmt, fv, when)
    if operation == "nper":
        if rate is None:
            raise ValueError("rate is required for operation='nper'.")
        if pmt is None:
            raise ValueError("pmt is required for operation='nper'.")
        return _financial.calculate_nper(rate, pmt, pv, fv, when)
    if operation == "rate":
        if nper is None:
            raise ValueError("nper is required for operation='rate'.")
        if pmt is None:
            raise ValueError("pmt is required for operation='rate'.")
        return _financial.calculate_rate(nper, pmt, pv, fv, when, guess)
    if operation == "depreciation":
        if cost is None:
            raise ValueError("cost is required for operation='depreciation'.")
        if salvage is None:
            raise ValueError("salvage is required for operation='depreciation'.")
        if life is None:
            raise ValueError("life is required for operation='depreciation'.")
        return _financial.calculate_depreciation(cost, salvage, life, method, period)
    if operation == "irr":
        if cash_flows is None:
            raise ValueError("cash_flows is required for operation='irr'.")
        return _financial.calculate_irr(cash_flows)
    raise ValueError(f"Unknown operation: {operation}")


def register(mcp) -> None:
    """Register financial tools on *mcp*."""
    mcp.tool()(goal_seek)
    mcp.tool()(loan_amortization)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(dcf_analysis)
    mcp.tool()(budget_variance_analysis)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(financial_ratio_analysis)
    mcp.tool()(break_even_analysis)
    mcp.tool()(create_sensitivity_table)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(time_value_calc)
