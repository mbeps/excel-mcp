from __future__ import annotations

from typing import TypedDict


class ColumnStatResult(TypedDict, total=False):
    """Descriptive statistics result for a column from column_statistics."""

    column: str
    count: int
    mean: float
    median: float
    min_val: float
    max_val: float
    std: float
    sum_val: float
    skewness: float | None
    kurtosis: float | None
    message: str


class RegressionResult(TypedDict, total=False):
    """Result returned by run_regression."""

    r_squared: float
    coefficients: dict[str, float]
    n_observations: int
    ss_residual: float
    ss_total: float
    output_sheet: str
    # Extended stats (present when statsmodels is available)
    adjusted_r_squared: float
    std_errors: list[float]
    t_values: list[float]
    p_values: list[float]
    f_statistic: float
    f_pvalue: float
    confidence_intervals: list[list[float]]
    intercept: float
    equation: str
    predictions: list[float]
