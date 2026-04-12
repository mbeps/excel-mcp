"""TypedDict schemas for financial analysis results (amortization, budget variance, ratios, goal seek)."""

from __future__ import annotations

from typing import TypedDict


class AmortizationPeriod(TypedDict):
    """One period row in a loan amortization schedule.

    Keys:
        period (int): Period number (1-based).
        payment (float): Total payment amount for the period.
        principal (float): Principal portion of the payment.
        interest (float): Interest portion of the payment.
        balance (float): Remaining balance after the payment.
    """

    period: int
    payment: float
    principal: float
    interest: float
    balance: float


class BudgetVarianceItem(TypedDict):
    """Variance analysis row for a single budget category.

    Keys:
        category (str): Budget category name.
        budget (float): Budgeted amount.
        actual (float): Actual amount.
        variance (float): Absolute variance (actual - budget).
        variance_pct (float): Percent variance (variance / budget) as a float (can be negative).
        status (str): Human-readable status (e.g., 'over', 'under', 'on target').
    """

    category: str
    budget: float
    actual: float
    variance: float
    variance_pct: float
    status: str


class BudgetVarianceSummary(TypedDict):
    """Totals row for a budget variance analysis.

    Keys:
        total_budget, total_actual, total_variance, total_variance_pct (float): Aggregated totals.
    """

    total_budget: float
    total_actual: float
    total_variance: float
    total_variance_pct: float


class RatioEntry(TypedDict, total=False):
    """A financial ratio entry, with optional benchmark comparison.

    Keys (all optional unless present):
        value (float | None): Calculated ratio value or None if unavailable.
        error (str): Error message if computation failed.
        benchmark (float): Optional benchmark value to compare against.
        comparison (str): Human-readable comparison result (e.g., 'above', 'below').
    """

    value: float | None
    error: str
    benchmark: float
    comparison: str


class GoalSeekResult(TypedDict):
    """Result returned by goal_seek.

    Keys:
        variable_cell (str): Cell reference whose value was adjusted.
        found_value (float): Value placed in the variable cell to reach the target.
        achieved_result (float): Achieved value of the target cell after substitution.
        target_value (float): Requested target value.
        converged (bool): Whether solver convergence was achieved.
        iterations (int): Number of iterations performed.
    """

    variable_cell: str
    found_value: float
    achieved_result: float
    target_value: float
    converged: bool
    iterations: int
