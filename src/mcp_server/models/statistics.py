"""Statistical result shapes: descriptive column stats and regression outputs."""

from __future__ import annotations

from typing import TypedDict


class RegressionExtras(TypedDict, total=False):
    """Extended regression diagnostics returned by the model-fitting helpers.

    The numpy fallback returns only ``r_squared``, ``adjusted_r_squared``,
    ``ss_residual``, ``ss_total`` and ``predictions``; statsmodels adds the
    remaining inference fields.
    """

    r_squared: float
    adjusted_r_squared: float
    ss_residual: float
    ss_total: float
    std_errors: list[float]
    t_values: list[float]
    p_values: list[float]
    f_statistic: float
    f_pvalue: float
    confidence_intervals: list[list[float]]
    predictions: list[float]


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
