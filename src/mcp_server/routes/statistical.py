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
    """Run an OLS linear regression and optionally write results to a sheet/file.

    Args:
        file_path: Input workbook path.
        sheet_name: Worksheet containing data.
        y_column: Dependent variable column name.
        x_columns: List of independent variable column names.
        header_row: 1-based header row index.
        output_sheet: Optional sheet name for regression output.
        output_file: Optional path to write results to a separate file.

    Returns:
        RegressionResult: Contains coefficients, R-squared, residuals and diagnostics.

    Notes:
        - Read-only unless `output_file`/`output_sheet` is provided (then mutates workbook/creates file).
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
    """Apply exponential smoothing (simple/Holt/Holt-Winters) to a time series column.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet name.
        column: Column to smooth.
        alpha: Smoothing factor for simple smoothing.
        new_column_name: Optional name for output column; if omitted, a generated name is used.
        header_row: 1-based header index.
        output_file: Optional file to write output.
        method: One of "simple", "holt", "holt_winters".
        seasonal_periods: Required for Holt-Winters.
        forecast_steps: Number of out-of-sample forecast steps to produce.
        smoothing_trend, smoothing_seasonal: Optional fixed smoothing parameters.

    Returns:
        dict: Summary and references to output column/sheet.

    Notes:
        - May modify workbook if `output_file`/`new_column_name` provided.
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
    """Run constrained optimisation using scipy to minimise (or maximise) an objective built from cell references.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet providing objective or referenced cells.
        objective_expression: Arithmetic expression using cell refs (e.g. "B2 * B3 - B4").
        variable_cells: Mapping {cell_ref: [lower_bound, upper_bound]} for optimisation variables.
        constraints: Optional list of {"expression": str, "type": "ineq"|"eq"} constraints.
        maximize: If True, the objective is maximised instead of minimised.
        tolerance: Convergence tolerance.
        max_iterations: Maximum solver iterations.

    Returns:
        SolverResult: Contains solution, status, and diagnostics.

    Notes:
        - May write back solution values into the workbook depending on implementation — document write semantics.
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

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet name.
        columns: Optional list of column names to include. If None, all numeric columns are used.
        output_sheet: Optional sheet name to write the matrix.
        output_file: Optional file path to write results.
        header_row: 1-based header index.

    Returns:
        dict: {"columns": [...], "matrix": [[float, ...], ...]}.

    Notes:
        - Read-only unless `output_sheet`/`output_file` is set.
    """
    return _statistical.correlation_matrix(file_path, sheet_name, columns, output_sheet, output_file, header_row)


def register(mcp) -> None:
    """Register tools on *mcp*."""
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(run_regression)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(run_exponential_smoothing)
    mcp.tool()(run_solver)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(correlation_matrix)
