from __future__ import annotations

from mcp.types import ToolAnnotations

import mcp_server.tools.statistical as _statistical
import mcp_server.tools.solver as _solver
from mcp_server.models.statistics import RegressionResult
from mcp_server.models.solver import SolverResult

__all__ = [
    "run_regression",
    "run_exponential_smoothing",
    "run_solver",
    "correlation_matrix",
]


def run_regression(
    file_path: str,
    sheet_name: str,
    y_column: str,
    x_columns: list[str],
    header_row: int = 1,
    output_sheet: str = "Regression Output",
    output_file: str | None = None,
) -> RegressionResult:
    """Run OLS linear regression and return coefficients, R-squared, and residuals.

    If output_file is provided, results are written to that file instead of file_path.
    """
    return _statistical.run_regression(
        file_path,
        sheet_name,
        y_column,
        x_columns,
        output_sheet=output_sheet,
        output_file=output_file,
        header_row=header_row,
    )


def run_exponential_smoothing(
    file_path: str,
    sheet_name: str,
    column: str,
    alpha: float = 0.3,
    new_column_name: str | None = None,
    header_row: int = 1,
    output_file: str | None = None,
    method: str = "simple",
    seasonal_periods: int | None = None,
    forecast_steps: int = 0,
    smoothing_trend: float | None = None,
    smoothing_seasonal: float | None = None,
) -> dict:
    """Apply exponential smoothing to a time series and write the result to a new column.

    method: "simple" (pandas EWM), "holt" (Holt linear trend), "holt_winters" (Holt-Winters seasonal).
    seasonal_periods: required for holt_winters (e.g. 12 for monthly data).
    forecast_steps: number of out-of-sample steps to forecast.
    smoothing_trend: trend smoothing factor for holt/holt_winters (0 < value <= 1). If omitted, statsmodels optimizes it.
    smoothing_seasonal: seasonal smoothing factor for holt_winters (0 < value <= 1). If omitted, statsmodels optimizes it.
    """
    return _statistical.run_exponential_smoothing(
        file_path,
        sheet_name,
        column,
        alpha,
        new_column_name,
        header_row,
        method=method,
        seasonal_periods=seasonal_periods,
        forecast_steps=forecast_steps,
        output_file=output_file,
        smoothing_trend=smoothing_trend,
        smoothing_seasonal=smoothing_seasonal,
    )


def run_solver(
    file_path: str,
    sheet_name: str,
    objective_expression: str,
    variable_cells: dict[str, list[float]],
    constraints: list[dict] | None = None,
    maximize: bool = False,
    tolerance: float = 1e-6,
    max_iterations: int = 1000,
) -> SolverResult:
    """Multi-variable constrained optimization using scipy.

    objective_expression: arithmetic expression using cell refs (e.g. "B2 * B3 - B4").
    variable_cells: {cell_ref: [lower_bound, upper_bound]} dict.
    constraints: list of {"expression": str, "type": "ineq"|"eq"} dicts.
    maximize: True to maximise instead of minimise.
    """
    normalised_cells: dict[str, tuple[float, float]] = {k: (v[0], v[1]) for k, v in variable_cells.items()}
    return _solver.run_solver(
        file_path,
        sheet_name,
        objective_expression,
        normalised_cells,
        constraints,
        maximize,
        tolerance,
        max_iterations,
    )


def correlation_matrix(
    file_path: str,
    sheet_name: str,
    columns: list[str] | None = None,
    output_sheet: str | None = None,
    output_file: str | None = None,
    header_row: int = 1,
) -> dict:
    """Compute a Pearson correlation matrix for numeric columns.

    columns: list of column names to include. If None, all numeric columns are used.
    output_sheet: if given, writes the matrix to this sheet (created if absent).
    output_file: target file for output; defaults to file_path.
    Returns: {"columns": [...], "matrix": [[float, ...], ...]}.
    """
    return _statistical.correlation_matrix(file_path, sheet_name, columns, output_sheet, output_file, header_row)


def register(mcp) -> None:
    """Register tools on *mcp*."""
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(run_regression)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(run_exponential_smoothing)
    mcp.tool()(run_solver)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(correlation_matrix)
