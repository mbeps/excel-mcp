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


class RegressionResult(TypedDict):
    """Result returned by run_regression."""

    r_squared: float
    coefficients: dict[str, float]
    n_observations: int
    ss_residual: float
    ss_total: float
    output_sheet: str
