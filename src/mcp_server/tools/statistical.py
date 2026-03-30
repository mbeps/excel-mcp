"""Statistical analysis: regression, moving average, exponential smoothing."""

from __future__ import annotations

from logging import Logger

import numpy as np
import pandas as pd
from openpyxl import Workbook

from mcp_server.models.statistics import RegressionResult
from mcp_server.utils.excel_helpers import get_sheet, load_workbook_safe, read_sheet_df, save_workbook_safe
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def run_regression(
    file_path: str,
    sheet_name: str,
    y_column: str,
    x_columns: list[str],
    output_sheet: str = "Regression Output",
    output_file: str | None = None,
    header_row: int = 1,
) -> RegressionResult:
    """Run OLS linear regression and write results to a new sheet.

    Returns coefficients, R-squared, and observation count.
    """
    df = read_sheet_df(file_path, sheet_name, header_row)

    missing = [c for c in [y_column, *x_columns] if c not in df.columns]
    if missing:
        raise ValueError(f"Columns not found: {missing}. Available: {list(df.columns)}")

    y = pd.to_numeric(df[y_column], errors="coerce")
    X = df[x_columns].apply(pd.to_numeric, errors="coerce")

    valid_idx = X.dropna().index.intersection(y.dropna().index)
    if len(valid_idx) < len(x_columns) + 2:
        raise ValueError(
            f"Not enough valid observations ({len(valid_idx)}) for regression with {len(x_columns)} predictor(s)."
        )

    y_clean = y[valid_idx].values
    X_clean = X.loc[valid_idx].values
    X_with_const = np.column_stack([np.ones(len(X_clean)), X_clean])

    coeffs, _residuals, _rank, _sv = np.linalg.lstsq(X_with_const, y_clean, rcond=None)

    y_pred = X_with_const @ coeffs
    ss_res = float(np.sum((y_clean - y_pred) ** 2))
    ss_tot = float(np.sum((y_clean - np.mean(y_clean)) ** 2))
    r_squared = 1.0 - ss_res / ss_tot if ss_tot != 0 else 0.0

    coefficients: dict[str, float] = {
        "intercept": round(float(coeffs[0]), 6),
        **{col: round(float(c), 6) for col, c in zip(x_columns, coeffs[1:])},
    }
    results: RegressionResult = {
        "r_squared": round(float(r_squared), 6),
        "coefficients": coefficients,
        "n_observations": int(len(y_clean)),
        "ss_residual": round(ss_res, 6),
        "ss_total": round(ss_tot, 6),
        "output_sheet": output_sheet,
    }

    target_file = output_file or file_path
    if output_file:
        wb: Workbook = Workbook()
        ws = wb.active
        ws.title = output_sheet
    else:
        wb = load_workbook_safe(file_path)
        if output_sheet in wb.sheetnames:
            del wb[output_sheet]
        ws = wb.create_sheet(title=output_sheet)
    try:
        ws["A1"] = "Regression Results"
        ws["A3"] = "R-Squared"
        ws["B3"] = results["r_squared"]
        ws["A4"] = "N Observations"
        ws["B4"] = results["n_observations"]
        ws["A5"] = "SS Residual"
        ws["B5"] = results["ss_residual"]
        ws["A6"] = "SS Total"
        ws["B6"] = results["ss_total"]
        ws["A8"] = "Variable"
        ws["B8"] = "Coefficient"

        coeffs_dict = results["coefficients"]

        ws.cell(row=9, column=1, value="Intercept")
        ws.cell(row=9, column=2, value=coeffs_dict["intercept"])
        for i, col in enumerate(x_columns, start=1):
            ws.cell(row=9 + i, column=1, value=col)
            ws.cell(row=9 + i, column=2, value=coeffs_dict[col])

        save_workbook_safe(wb, target_file)
    finally:
        wb.close()

    logger.info("Regression: R²=%.4f, n=%d, output_sheet=%s", r_squared, len(y_clean), output_sheet)
    return results


def run_exponential_smoothing(
    file_path: str,
    sheet_name: str,
    value_column: str,
    alpha: float = 0.3,
    output_column: str | None = None,
    header_row: int = 1,
) -> dict[str, object]:
    """Compute and write exponentially smoothed (EWM) series to the sheet.

    alpha: smoothing factor (0 < alpha <= 1)
    """
    if not (0 < alpha <= 1):
        raise ValueError("alpha must be between 0 (exclusive) and 1 (inclusive).")

    df = read_sheet_df(file_path, sheet_name, header_row)
    if value_column not in df.columns:
        raise ValueError(f"Column '{value_column}' not found. Available: {list(df.columns)}")

    out_col = (output_column or f"EWM{alpha:.2f}_{value_column}")[:31]
    series = pd.to_numeric(df[value_column], errors="coerce").ewm(alpha=alpha, adjust=False).mean()
    values = series.tolist()

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        col_idx: int | None = None
        for cell in ws[header_row]:
            if cell.value == out_col:
                col_idx = cell.column
                break
        if col_idx is None:
            col_idx = ws.max_column + 1
            ws.cell(row=header_row, column=col_idx, value=out_col)

        for i, val in enumerate(values):
            ws.cell(row=header_row + 1 + i, column=col_idx, value=None if pd.isna(val) else float(val))

        save_workbook_safe(wb, file_path)
    finally:
        wb.close()

    logger.info("Exponential smoothing (alpha=%.2f) written to column '%s'", alpha, out_col)
    return {"alpha": alpha, "output_column": out_col, "rows_written": len(values)}
