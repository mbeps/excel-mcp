from __future__ import annotations

from typing import Literal

from mcp.types import ToolAnnotations

import mcp_server.tools.financial as _financial

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
    """Find a variable cell value that makes an expression evaluate to a target and write the result.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet containing the expression.
        variable_cell: Cell reference to adjust (e.g. "B2").
        expression: Arithmetic expression referencing worksheet cells (string).
        target_value: Numeric target value for the expression.
        initial_value: Starting guess for the solver.
        tolerance: Convergence tolerance.
        max_iterations: Maximum solver iterations.

    Returns:
        dict: Result with solved value, status, and iterations used.

    Notes:
        - Destructive: writes the solved value back to the workbook.
        - Recommend adding a short example expression in docs.
    """
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
    """Generate a loan amortization schedule for given principal, rate and term.

    Args:
        principal: Loan principal amount.
        annual_rate: Annual interest rate (fractional, e.g. 0.05 for 5%).
        years: Term in years.
        payments_per_year: Payment frequency (default 12).

    Returns:
        dict: Schedule rows and totals including payment amount, interest, principal breakdown.
    """
    return _financial.loan_amortization(principal, annual_rate, years, payments_per_year)


def dcf_analysis(
    cash_flows: list[float],
    discount_rate: float,
    terminal_growth_rate: float = 0.02,
    initial_investment: float = 0.0,
) -> dict:
    """Compute Discounted Cash Flow valuation with a Gordon Growth Model terminal value.

    Args:
        cash_flows: List of cash flows (period-ordered), first item normally year 0 investment (negative).
        discount_rate: Discount rate as decimal.
        terminal_growth_rate: Perpetuity growth for terminal value.
        initial_investment: Optional initial outlay to include in NPV.

    Returns:
        dict: NPV, terminal value, IRR and breakdowns.
    """
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
    """Compare budget vs actual values in a sheet and return variances per category.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet name.
        category_column, budget_column, actual_column: Column identifiers for the analysis.
        header_row: 1-based header row index.
        output_file: Optional path to write results.

    Returns:
        dict: Per-category variance and status.

    Notes:
        - Mutates workbook only if `output_file` provided.
    """
    return _financial.budget_variance_analysis(
        file_path,
        sheet_name,
        category_column,
        budget_column,
        actual_column,
        header_row,
        output_file,
    )


def financial_ratio_analysis(
    financial_data: dict[str, float],
    industry_benchmarks: dict[str, float] | None = None,
) -> dict:
    """Compute common financial ratios from raw financial metric inputs and compare them to benchmarks.

    Args:
        financial_data: Dict of raw metric values keyed by component name. Valid keys:
            current_assets, current_liabilities, inventory, total_debt, total_equity,
            net_income, total_assets, revenue, gross_profit, operating_income, ebitda,
            interest_expense. E.g. {"current_assets": 500000, "current_liabilities": 250000}.
        industry_benchmarks: Optional dict of benchmark ratio values to compare against,
            e.g. {"current_ratio": 2.0, "roe": 0.15}.

    Returns:
        dict: Computed ratios and optional benchmark comparisons.

    Notes:
        - This function is pure math and does not touch files.
    """
    return _financial.financial_ratio_analysis(financial_data, industry_benchmarks)


def break_even_analysis(fixed_costs: float, price_per_unit: float, variable_cost_per_unit: float) -> dict:
    """Calculate break-even units and revenue given fixed and variable costs.

    Args:
        fixed_costs: Total fixed costs.
        price_per_unit: Selling price per unit.
        variable_cost_per_unit: Variable cost per unit.

    Returns:
        dict: break_even_units and break_even_revenue.
    """
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
    """Create a 1- or 2-variable sensitivity table in the workbook.

    Evaluates `expression` over supplied value grids.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet containing the base formula or output cell.
        output_cell: Cell that contains or references the expression to evaluate.
        expression: Expression that will be evaluated relative to variable names.
        var1_name, var1_values: Name and values for variable 1.
        var2_name, var2_values: Optional name/values for a second variable (two-way table).

    Returns:
        dict: Summary including output range and written values.

    Notes:
        - Mutates workbook by inserting the table; confirm overwrite semantics
          when the target output area overlaps data.
    """
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
    """Perform a variety of time-value-of-money calculations and depreciation methods.

    Args:
        operation: One of "fv", "pv", "nper", "rate", "depreciation", "irr".
        rate, nper, pmt, pv, fv, when, guess: Parameters depending on operation.
        cost, salvage, life, method, period: Parameters for depreciation operations.
        cash_flows: For IRR, list of floats.

    Returns:
        dict: Operation-specific outputs (e.g. numeric answer, schedule, irr value).

    Raises:
        ValueError: If required args for the selected operation are missing.

    Notes:
        - Pure calculations except for methods that may write results when integrated into workbook workflows.
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
