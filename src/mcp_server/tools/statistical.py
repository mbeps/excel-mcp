"""Statistical analysis: regression, moving average, exponential smoothing."""

from __future__ import annotations

import os
from logging import Logger

import numpy as np
import pandas as pd
from openpyxl import Workbook

from mcp_server.models.statistics import RegressionResult
from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    read_sheet_df,
    save_workbook_safe,
    validate_file_path,
)
from mcp_server.utils.logger import configure_logging

try:
    import statsmodels.api as sm

    _HAS_STATSMODELS = True
except ImportError:  # pragma: no cover
    _HAS_STATSMODELS = False

logger: Logger = configure_logging(__name__)


def _build_equation(x_columns: list[str], coeffs_dict: dict[str, float]) -> str:
    """Build a human-readable equation string."""
    parts = [f"{coeffs_dict['intercept']:.6f}"]
    for col in x_columns:
        c = coeffs_dict[col]
        parts.append(f"+ {c:.6f}*{col}" if c >= 0 else f"- {abs(c):.6f}*{col}")
    return "Y = " + " ".join(parts)


def _regression_statsmodels(
    y_clean: np.ndarray,
    X_clean: np.ndarray,
    x_columns: list[str],
) -> tuple[dict[str, float], dict[str, object]]:
    """Run regression via statsmodels OLS and return (coefficients, extras)."""
    X_with_const = sm.add_constant(X_clean, has_constant="add")
    model = sm.OLS(y_clean, X_with_const)
    fit = model.fit()

    coeffs_array = fit.params
    coefficients: dict[str, float] = {
        "intercept": round(float(coeffs_array[0]), 6),
        **{col: round(float(c), 6) for col, c in zip(x_columns, coeffs_array[1:])},
    }

    ci = fit.conf_int(alpha=0.05)
    confidence_intervals: list[list[float]] = [
        [round(float(ci[i, 0]), 6), round(float(ci[i, 1]), 6)] for i in range(len(ci))
    ]

    y_pred = fit.predict(X_with_const)
    extras: dict[str, object] = {
        "r_squared": round(float(fit.rsquared), 6),
        "adjusted_r_squared": round(float(fit.rsquared_adj), 6),
        "ss_residual": round(float(fit.ssr), 6),
        "ss_total": round(float(fit.centered_tss), 6),
        "std_errors": [round(float(v), 6) for v in fit.bse],
        "t_values": [round(float(v), 6) for v in fit.tvalues],
        "p_values": [round(float(v), 6) for v in fit.pvalues],
        "f_statistic": round(float(fit.fvalue), 6),
        "f_pvalue": round(float(fit.f_pvalue), 6),
        "confidence_intervals": confidence_intervals,
        "predictions": [round(float(v), 6) for v in y_pred],
    }
    return coefficients, extras


def _regression_numpy(
    y_clean: np.ndarray,
    X_clean: np.ndarray,
    x_columns: list[str],
) -> tuple[dict[str, float], dict[str, object]]:
    """Fallback regression using numpy lstsq."""
    X_with_const = np.column_stack([np.ones(len(X_clean)), X_clean])
    coeffs, _residuals, _rank, _sv = np.linalg.lstsq(X_with_const, y_clean, rcond=None)

    y_pred = X_with_const @ coeffs
    ss_res = float(np.sum((y_clean - y_pred) ** 2))
    ss_tot = float(np.sum((y_clean - np.mean(y_clean)) ** 2))
    r_sq = 1.0 - ss_res / ss_tot if ss_tot != 0 else (1.0 if ss_res == 0 else 0.0)

    n = len(y_clean)
    p = X_clean.shape[1]
    adj_r_sq = 1.0 - (1.0 - r_sq) * (n - 1) / (n - p - 1) if n > p + 1 else r_sq

    coefficients: dict[str, float] = {
        "intercept": round(float(coeffs[0]), 6),
        **{col: round(float(c), 6) for col, c in zip(x_columns, coeffs[1:])},
    }
    extras: dict[str, object] = {
        "r_squared": round(r_sq, 6),
        "adjusted_r_squared": round(adj_r_sq, 6),
        "ss_residual": round(ss_res, 6),
        "ss_total": round(ss_tot, 6),
        "predictions": [round(float(v), 6) for v in y_pred],
    }
    return coefficients, extras


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

    Uses statsmodels for full inference (p-values, t-stats, confidence
    intervals); falls back to numpy if statsmodels is unavailable.
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

    if _HAS_STATSMODELS:
        coefficients, extras = _regression_statsmodels(y_clean, X_clean, x_columns)
    else:
        coefficients, extras = _regression_numpy(y_clean, X_clean, x_columns)

    equation = _build_equation(x_columns, coefficients)

    results: RegressionResult = {
        "r_squared": extras["r_squared"],  # type: ignore[typeddict-item]
        "coefficients": coefficients,
        "n_observations": int(len(y_clean)),
        "ss_residual": extras["ss_residual"],  # type: ignore[typeddict-item]
        "ss_total": extras["ss_total"],  # type: ignore[typeddict-item]
        "output_sheet": output_sheet,
        "intercept": coefficients["intercept"],
        "equation": equation,
        "predictions": extras["predictions"],  # type: ignore[typeddict-item]
        "adjusted_r_squared": extras["adjusted_r_squared"],  # type: ignore[typeddict-item]
    }
    if "std_errors" in extras:
        results["std_errors"] = extras["std_errors"]  # type: ignore[typeddict-item]
        results["t_values"] = extras["t_values"]  # type: ignore[typeddict-item]
        results["p_values"] = extras["p_values"]  # type: ignore[typeddict-item]
        results["f_statistic"] = extras["f_statistic"]  # type: ignore[typeddict-item]
        results["f_pvalue"] = extras["f_pvalue"]  # type: ignore[typeddict-item]
        results["confidence_intervals"] = extras["confidence_intervals"]  # type: ignore[typeddict-item]

    # ── Write results to sheet ──
    target_file = output_file or file_path
    if output_file and os.path.isfile(output_file) and os.path.realpath(output_file) == os.path.realpath(file_path):
        wb: Workbook = load_workbook_safe(file_path)
        if output_sheet in wb.sheetnames:
            del wb[output_sheet]
        ws = wb.create_sheet(title=output_sheet)
    elif output_file:
        wb = Workbook()
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
        ws["A4"] = "Adjusted R-Squared"
        ws["B4"] = results.get("adjusted_r_squared")
        ws["A5"] = "N Observations"
        ws["B5"] = results["n_observations"]
        ws["A6"] = "SS Residual"
        ws["B6"] = results["ss_residual"]
        ws["A7"] = "SS Total"
        ws["B7"] = results["ss_total"]
        if "f_statistic" in results:
            ws["A8"] = "F-Statistic"
            ws["B8"] = results["f_statistic"]
            ws["A9"] = "F p-value"
            ws["B9"] = results["f_pvalue"]

        header_start = 11
        headers = ["Variable", "Coefficient"]
        if "std_errors" in results:
            headers.extend(["Std Error", "t-value", "p-value", "CI Lower", "CI Upper"])
        for j, h in enumerate(headers, start=1):
            ws.cell(row=header_start, column=j, value=h)

        coeffs_dict = results["coefficients"]
        var_names = ["Intercept", *x_columns]
        for i, var in enumerate(var_names):
            row = header_start + 1 + i
            key = "intercept" if var == "Intercept" else var
            ws.cell(row=row, column=1, value=var)
            ws.cell(row=row, column=2, value=coeffs_dict[key])
            if "std_errors" in results:
                ws.cell(row=row, column=3, value=results["std_errors"][i])
                ws.cell(row=row, column=4, value=results["t_values"][i])
                ws.cell(row=row, column=5, value=results["p_values"][i])
                ws.cell(row=row, column=6, value=results["confidence_intervals"][i][0])
                ws.cell(row=row, column=7, value=results["confidence_intervals"][i][1])

        save_workbook_safe(wb, target_file)
    finally:
        wb.close()

    logger.info(
        "Regression: R²=%.4f, adj_R²=%.4f, n=%d, output_sheet=%s",
        results["r_squared"],
        results.get("adjusted_r_squared", 0),
        len(y_clean),
        output_sheet,
    )
    return results


def run_exponential_smoothing(
    file_path: str,
    sheet_name: str,
    value_column: str,
    alpha: float = 0.3,
    output_column: str | None = None,
    header_row: int = 1,
    method: str = "simple",
    seasonal_periods: int | None = None,
    forecast_steps: int = 0,
    output_file: str | None = None,
) -> dict[str, object]:
    """Compute and write exponentially smoothed series to the sheet.

    Args:
        alpha: smoothing factor (0 < alpha <= 1).
        method: "simple" for pandas EWM, "holt" for Holt linear trend,
                "holt_winters" for Holt-Winters seasonal (requires statsmodels).
        seasonal_periods: period length for Holt-Winters (e.g. 12 for monthly).
        forecast_steps: number of out-of-sample forecast steps to append.
    """
    if not (0 < alpha <= 1):
        raise ValueError("alpha must be between 0 (exclusive) and 1 (inclusive).")

    valid_methods = {"simple", "holt", "holt_winters"}
    if method not in valid_methods:
        raise ValueError(f"method must be one of {valid_methods}, got '{method}'.")

    df = read_sheet_df(file_path, sheet_name, header_row)
    if value_column not in df.columns:
        raise ValueError(f"Column '{value_column}' not found. Available: {list(df.columns)}")

    numeric_series = pd.to_numeric(df[value_column], errors="coerce").dropna()
    if len(numeric_series) < 3:
        raise ValueError("Not enough valid data points for smoothing.")

    out_col = (output_column or f"EWM{alpha:.2f}_{value_column}")[:31]
    forecast_values: list[float] = []
    model_info: dict[str, object] = {}

    if method in ("holt", "holt_winters") and _HAS_STATSMODELS:
        from statsmodels.tsa.holtwinters import ExponentialSmoothing as HW

        seasonal = "add" if method == "holt_winters" else None
        sp = seasonal_periods or (12 if method == "holt_winters" else None)
        if method == "holt_winters" and sp is not None and len(numeric_series) < 2 * sp:
            raise ValueError(
                f"Need at least {2 * sp} observations for Holt-Winters with seasonal_periods={sp}, "
                f"got {len(numeric_series)}."
            )

        hw_model = HW(
            numeric_series.values,
            trend="add",
            seasonal=seasonal,
            seasonal_periods=sp,
        )
        fit = hw_model.fit(smoothing_level=alpha, optimized=False)
        values = fit.fittedvalues.tolist()

        if forecast_steps > 0:
            forecast_values = fit.forecast(forecast_steps).tolist()

        model_info = {
            "method": method,
            "trend": "add",
            "seasonal": seasonal or "none",
            "seasonal_periods": sp,
            "aic": round(float(fit.aic), 4) if hasattr(fit, "aic") else None,
            "bic": round(float(fit.bic), 4) if hasattr(fit, "bic") else None,
        }
    elif method in ("holt", "holt_winters") and not _HAS_STATSMODELS:
        logger.warning("statsmodels not available; falling back to simple EWM.")
        full_series = pd.to_numeric(df[value_column], errors="coerce")
        values = full_series.ewm(alpha=alpha, adjust=False).mean().tolist()
        model_info = {"method": "simple", "note": "statsmodels not installed; fell back to simple EWM"}
    else:
        full_series = pd.to_numeric(df[value_column], errors="coerce")
        values = full_series.ewm(alpha=alpha, adjust=False).mean().tolist()
        model_info = {"method": "simple"}

        # SES forecast: last smoothed value repeated
        if forecast_steps > 0 and values:
            last_smoothed = values[-1]
            if not pd.isna(last_smoothed):
                forecast_values = [round(float(last_smoothed), 6)] * forecast_steps

    target_file = output_file or file_path
    if output_file and os.path.isfile(output_file) and os.path.realpath(output_file) == os.path.realpath(file_path):
        wb = load_workbook_safe(file_path)
    elif output_file:
        import openpyxl as _opx

        wb = _opx.Workbook()
        ws_src = wb.active
        ws_src.title = sheet_name
        src_wb = load_workbook_safe(file_path, data_only=True)
        src_ws = get_sheet(src_wb, sheet_name)
        for row in src_ws.iter_rows(values_only=False):
            for cell in row:
                ws_src.cell(row=cell.row, column=cell.column, value=cell.value)
        src_wb.close()
    else:
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

        if forecast_values:
            fc_col = ws.max_column + 1
            ws.cell(row=header_row, column=fc_col, value=f"Forecast_{out_col}"[:31])
            for i, val in enumerate(forecast_values):
                ws.cell(row=header_row + 1 + len(values) + i, column=fc_col, value=round(float(val), 6))

        save_workbook_safe(wb, target_file)
    finally:
        wb.close()

    logger.info("Exponential smoothing (%s, alpha=%.2f) written to column '%s'", method, alpha, out_col)
    return {
        "alpha": alpha,
        "output_column": out_col,
        "rows_written": len(values),
        "forecast_steps": len(forecast_values),
        **model_info,
    }


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
    output_file: if given, writes to a separate file; defaults to file_path.
    Returns a dict with "columns" list and "matrix" (list of rows, each a list of floats).
    """
    validate_file_path(file_path, must_exist=True)
    dest_path = output_file or file_path
    if output_file:
        validate_file_path(output_file, must_exist=False)

    df = read_sheet_df(file_path, sheet_name, header_row=header_row)
    if df.empty:
        raise ValueError(f"Sheet '{sheet_name}' is empty.")

    if columns:
        missing = [c for c in columns if c not in df.columns]
        if missing:
            raise ValueError(f"Columns not found: {missing}")
        df = df[columns]

    numeric_df = df.select_dtypes(include="number")
    if numeric_df.empty:
        raise ValueError("No numeric columns found for correlation matrix.")

    corr = numeric_df.corr()
    col_names = list(corr.columns)
    matrix = [[round(v, 6) if pd.notna(v) else None for v in row] for row in corr.values.tolist()]

    if output_sheet:
        wb = load_workbook_safe(dest_path)
        try:
            if output_sheet in wb.sheetnames:
                ws = wb[output_sheet]
                ws.delete_rows(1, ws.max_row)
            else:
                ws = wb.create_sheet(output_sheet)
            # Header row
            ws.cell(row=1, column=1, value="")
            for ci, name in enumerate(col_names, start=2):
                ws.cell(row=1, column=ci, value=name)
            # Data rows
            for ri, (name, row_vals) in enumerate(zip(col_names, matrix), start=2):
                ws.cell(row=ri, column=1, value=name)
                for ci, val in enumerate(row_vals, start=2):
                    ws.cell(row=ri, column=ci, value=val)
            save_workbook_safe(wb, dest_path)
        finally:
            wb.close()

    logger.info("Computed correlation matrix for %d columns in '%s'", len(col_names), sheet_name)
    result: dict = {
        "columns": col_names,
        "matrix": matrix,
        "file_path": file_path,
    }
    if output_sheet:
        result["output_sheet"] = output_sheet
        result["output_file"] = dest_path
    return result
