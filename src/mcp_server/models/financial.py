from __future__ import annotations

from typing import TypedDict


class AmortizationPeriod(TypedDict):
    """One period row in a loan amortization schedule."""

    period: int
    payment: float
    principal: float
    interest: float
    balance: float


class BudgetVarianceItem(TypedDict):
    """Variance analysis row for a single budget category."""

    category: str
    budget: float
    actual: float
    variance: float
    variance_pct: float
    status: str


class BudgetVarianceSummary(TypedDict):
    """Totals row for a budget variance analysis."""

    total_budget: float
    total_actual: float
    total_variance: float
    total_variance_pct: float


class RatioEntry(TypedDict, total=False):
    """A financial ratio entry, with optional benchmark comparison."""

    value: float | None
    error: str
    benchmark: float
    comparison: str


class GoalSeekResult(TypedDict):
    """Result returned by goal_seek."""

    variable_cell: str
    found_value: float
    achieved_result: float
    target_value: float
    converged: bool
    iterations: int
