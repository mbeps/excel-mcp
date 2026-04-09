from __future__ import annotations

"""Statistical result shapes: descriptive column stats and regression outputs.
"""

from typing import TypedDict


class ColumnStatResult(TypedDict, total=False):
    """Descriptive statistics result for a column returned by column_statistics.

    Keys (total=False): fields are optional and present when calculable.
        column (str): Column name.
        count, mean, median, min_val, max_val, std, sum_val: numeric stats.
        skewness, kurtosis (float | None): Higher-order moments; may be None.
        message (str): Informational message when stats are unavailable.
    """

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
    """Result returned by run_regression.

    Keys (total=False): present fields depend on whether extended stats are available.
        r_squared (float): R-squared value for the model.
        coefficients (dict[str, float]): Predictor coefficients.
        n_observations (int): Number of observations used.
        ss_residual, ss_total: Sum-of-squares values.
        output_sheet (str): If regression output is written to a sheet, its name.
        Extended stats: adjusted_r_squared, std_errors, t_values, p_values, f_statistic, f_pvalue,
        confidence_intervals, intercept, equation, predictions.
    """

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
