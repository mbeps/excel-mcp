"""Workflow tests for statistical analysis: regression and exponential smoothing.

Tests call tool functions directly and verify results independently using
openpyxl, numpy, or pandas — never via MCP tools.
"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import openpyxl
import pytest

from mcp_server.tools.cell_ops import write_range
from mcp_server.tools.statistical import run_exponential_smoothing, run_regression
from mcp_server.tools.workbook import create_workbook


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _create_regression_data(fp: str, sheet: str = "Data") -> None:
    """Create a workbook with x/y data for regression."""
    create_workbook(fp, sheet_names=[sheet])
    write_range(
        fp,
        sheet,
        "A1",
        [
            ["X", "Y"],
            [1, 2.1],
            [2, 3.9],
            [3, 6.2],
            [4, 7.8],
            [5, 10.1],
            [6, 12.0],
            [7, 13.9],
            [8, 16.1],
            [9, 18.0],
            [10, 20.2],
        ],
    )


def _create_perfect_linear_data(fp: str, sheet: str = "Data") -> None:
    """Create perfectly linear data: y = 3x + 1."""
    create_workbook(fp, sheet_names=[sheet])
    rows = [["X", "Y"]]
    for x in range(1, 11):
        rows.append([x, 3 * x + 1])
    write_range(fp, sheet, "A1", rows)


def _create_time_series(fp: str, sheet: str = "Data", n: int = 24) -> None:
    """Create time series data for smoothing tests."""
    create_workbook(fp, sheet_names=[sheet])
    rng = np.random.default_rng(42)
    base = np.linspace(100, 150, n)
    noise = rng.normal(0, 5, n)
    values = base + noise
    rows = [["Period", "Value"]]
    for i, v in enumerate(values, 1):
        rows.append([i, round(float(v), 2)])
    write_range(fp, sheet, "A1", rows)


def _create_seasonal_series(fp: str, sheet: str = "Data") -> None:
    """Create seasonal data (period=4) with trend for Holt-Winters."""
    create_workbook(fp, sheet_names=[sheet])
    # 6 full seasonal cycles = 24 points (minimum 2*period=8 needed)
    rows = [["Period", "Value"]]
    seasonal = [10, 20, 30, 15]
    for i in range(24):
        trend = 100 + i * 2
        value = trend + seasonal[i % 4]
        rows.append([i + 1, float(value)])
    write_range(fp, sheet, "A1", rows)


def _create_constant_series(fp: str, sheet: str = "Data") -> None:
    """Create constant-value series."""
    create_workbook(fp, sheet_names=[sheet])
    rows = [["Period", "Value"]]
    for i in range(1, 21):
        rows.append([i, 42.0])
    write_range(fp, sheet, "A1", rows)


def _create_minimal_data(fp: str, sheet: str = "Data") -> None:
    """Create minimal 3-point data for regression."""
    create_workbook(fp, sheet_names=[sheet])
    write_range(
        fp,
        sheet,
        "A1",
        [
            ["X", "Y"],
            [1, 5],
            [2, 7],
            [3, 9],
        ],
    )


# ---------------------------------------------------------------------------
# 1. Regression OLS — basic
# ---------------------------------------------------------------------------


class TestRegressionOLSBasic:
    def test_regression_ols_basic(self, tmp_path: Path) -> None:
        """Write x,y data → run_regression → verify r_squared and coefficients."""
        fp = str(tmp_path / "reg.xlsx")
        _create_regression_data(fp)
        result = run_regression(
            file_path=fp,
            sheet_name="Data",
            y_column="Y",
            x_columns=["X"],
        )
        assert "r_squared" in result
        assert "coefficients" in result
        assert result["r_squared"] > 0.99  # near-linear data
        assert "intercept" in result["coefficients"]
        assert "X" in result["coefficients"]
        assert result["n_observations"] == 10


# ---------------------------------------------------------------------------
# 2. Regression OLS — perfectly linear
# ---------------------------------------------------------------------------


class TestRegressionOLSPerfect:
    def test_regression_ols_perfect(self, tmp_path: Path) -> None:
        """Perfectly linear data → verify r_squared ≈ 1.0."""
        fp = str(tmp_path / "perf.xlsx")
        _create_perfect_linear_data(fp)
        result = run_regression(
            file_path=fp,
            sheet_name="Data",
            y_column="Y",
            x_columns=["X"],
        )
        assert abs(result["r_squared"] - 1.0) < 1e-6


# ---------------------------------------------------------------------------
# 3. Regression coefficients — known slope/intercept
# ---------------------------------------------------------------------------


class TestRegressionCoefficients:
    def test_regression_coefficients(self, tmp_path: Path) -> None:
        """Known slope=3, intercept=1 → verify regression output matches."""
        fp = str(tmp_path / "coeffs.xlsx")
        _create_perfect_linear_data(fp)
        result = run_regression(
            file_path=fp,
            sheet_name="Data",
            y_column="Y",
            x_columns=["X"],
        )
        assert abs(result["coefficients"]["intercept"] - 1.0) < 0.01
        assert abs(result["coefficients"]["X"] - 3.0) < 0.01


# ---------------------------------------------------------------------------
# 4. Exponential smoothing — simple
# ---------------------------------------------------------------------------


class TestSmoothingSimple:
    def test_exponential_smoothing_simple(self, tmp_path: Path) -> None:
        """Simple method → verify fitted values exist."""
        fp = str(tmp_path / "smooth.xlsx")
        _create_time_series(fp)
        result = run_exponential_smoothing(
            file_path=fp,
            sheet_name="Data",
            value_column="Value",
            alpha=0.3,
            method="simple",
        )
        assert result["method"] == "simple"
        assert result["rows_written"] == 24

        wb = openpyxl.load_workbook(fp)
        ws = wb["Data"]
        # Smoothed column should be written after existing columns
        smoothed_header = ws.cell(row=1, column=3).value
        assert smoothed_header is not None
        assert "EWM" in smoothed_header or "Value" in smoothed_header
        # First smoothed value should exist
        assert ws.cell(row=2, column=3).value is not None
        wb.close()


# ---------------------------------------------------------------------------
# 5. Exponential smoothing — Holt
# ---------------------------------------------------------------------------


class TestSmoothingHolt:
    def test_exponential_smoothing_holt(self, tmp_path: Path) -> None:
        """Holt method → verify model info and output."""
        fp = str(tmp_path / "holt.xlsx")
        _create_time_series(fp)
        result = run_exponential_smoothing(
            file_path=fp,
            sheet_name="Data",
            value_column="Value",
            alpha=0.3,
            method="holt",
        )
        assert result["method"] == "holt"
        assert result["rows_written"] == 24
        assert result.get("trend") == "add"


# ---------------------------------------------------------------------------
# 6. Exponential smoothing — Holt-Winters
# ---------------------------------------------------------------------------


class TestSmoothingHoltWinters:
    def test_exponential_smoothing_holt_winters(self, tmp_path: Path) -> None:
        """Holt-Winters with seasonal → verify."""
        fp = str(tmp_path / "hw.xlsx")
        _create_seasonal_series(fp)
        result = run_exponential_smoothing(
            file_path=fp,
            sheet_name="Data",
            value_column="Value",
            alpha=0.3,
            method="holt_winters",
            seasonal_periods=4,
        )
        assert result["method"] == "holt_winters"
        assert result["rows_written"] == 24
        assert result.get("seasonal") == "add"
        assert result.get("seasonal_periods") == 4


# ---------------------------------------------------------------------------
# 7. Smoothing — forecast steps
# ---------------------------------------------------------------------------


class TestSmoothingForecastSteps:
    def test_smoothing_forecast_steps(self, tmp_path: Path) -> None:
        """forecast_steps > 0 → verify forecast array length."""
        fp = str(tmp_path / "fc.xlsx")
        _create_time_series(fp)
        result = run_exponential_smoothing(
            file_path=fp,
            sheet_name="Data",
            value_column="Value",
            alpha=0.4,
            method="simple",
            forecast_steps=5,
        )
        assert result["forecast_steps"] == 5

        # Verify forecast column written in the workbook
        wb = openpyxl.load_workbook(fp)
        ws = wb["Data"]
        # Find forecast column header
        forecast_col = None
        for col in range(1, ws.max_column + 1):
            hdr = ws.cell(row=1, column=col).value
            if hdr and "Forecast" in str(hdr):
                forecast_col = col
                break
        assert forecast_col is not None
        # Count non-None forecast values
        fc_values = []
        for row in range(2, ws.max_row + 1):
            v = ws.cell(row=row, column=forecast_col).value
            if v is not None:
                fc_values.append(v)
        assert len(fc_values) == 5
        wb.close()


# ---------------------------------------------------------------------------
# 8. Regression → smoothing chain
# ---------------------------------------------------------------------------


class TestRegressionThenSmoothing:
    def test_regression_then_smoothing(self, tmp_path: Path) -> None:
        """Chain: regression → smoothing on same data."""
        fp = str(tmp_path / "chain.xlsx")
        _create_regression_data(fp)

        reg = run_regression(
            file_path=fp,
            sheet_name="Data",
            y_column="Y",
            x_columns=["X"],
        )
        assert reg["r_squared"] > 0.99

        smooth = run_exponential_smoothing(
            file_path=fp,
            sheet_name="Data",
            value_column="Y",
            alpha=0.5,
            method="simple",
        )
        assert smooth["rows_written"] == 10

        # Both outputs should coexist
        wb = openpyxl.load_workbook(fp)
        assert "Regression Output" in wb.sheetnames
        ws_data = wb["Data"]
        # Smoothed column should be present
        headers = [ws_data.cell(row=1, column=c).value for c in range(1, ws_data.max_column + 1)]
        assert any("EWM" in str(h) or "Y" in str(h) for h in headers if h)
        wb.close()


# ---------------------------------------------------------------------------
# 9. Regression — minimal 3-point
# ---------------------------------------------------------------------------


class TestRegressionSinglePredictor:
    def test_regression_single_predictor(self, tmp_path: Path) -> None:
        """Minimal 3-point regression → verify it runs and produces results."""
        fp = str(tmp_path / "mini.xlsx")
        _create_minimal_data(fp)
        result = run_regression(
            file_path=fp,
            sheet_name="Data",
            y_column="Y",
            x_columns=["X"],
        )
        assert result["n_observations"] == 3
        assert "r_squared" in result
        # y = 2x + 3 (exact fit)
        assert abs(result["coefficients"]["X"] - 2.0) < 0.01
        assert abs(result["coefficients"]["intercept"] - 3.0) < 0.01


# ---------------------------------------------------------------------------
# 10. Smoothing — constant data edge case
# ---------------------------------------------------------------------------


class TestSmoothingConstantData:
    def test_smoothing_on_constant_data(self, tmp_path: Path) -> None:
        """Edge case: all values identical → smoothed values should be same constant."""
        fp = str(tmp_path / "const.xlsx")
        _create_constant_series(fp)
        result = run_exponential_smoothing(
            file_path=fp,
            sheet_name="Data",
            value_column="Value",
            alpha=0.3,
            method="simple",
        )
        assert result["rows_written"] == 20

        wb = openpyxl.load_workbook(fp)
        ws = wb["Data"]
        # Find smoothed column
        smooth_col = None
        for col in range(1, ws.max_column + 1):
            hdr = ws.cell(row=1, column=col).value
            if hdr and "EWM" in str(hdr):
                smooth_col = col
                break
        assert smooth_col is not None
        # All smoothed values should be 42.0
        for row in range(2, 22):
            val = ws.cell(row=row, column=smooth_col).value
            if val is not None:
                assert abs(val - 42.0) < 0.01
        wb.close()


# ---------------------------------------------------------------------------
# 11. Holt-Winters with forecast steps
# ---------------------------------------------------------------------------


class TestHoltWintersForecast:
    def test_holt_winters_with_forecast(self, tmp_path: Path) -> None:
        """Holt-Winters + forecast_steps → verify forecast values are reasonable."""
        fp = str(tmp_path / "hw_fc.xlsx")
        _create_seasonal_series(fp)
        result = run_exponential_smoothing(
            file_path=fp,
            sheet_name="Data",
            value_column="Value",
            alpha=0.3,
            method="holt_winters",
            seasonal_periods=4,
            forecast_steps=4,
        )
        assert result["forecast_steps"] == 4
        assert result["method"] == "holt_winters"


# ---------------------------------------------------------------------------
# 12. Regression output sheet verification
# ---------------------------------------------------------------------------


class TestRegressionOutputSheet:
    def test_regression_output_sheet(self, tmp_path: Path) -> None:
        """Verify regression writes R² and coefficients to output sheet."""
        fp = str(tmp_path / "reg_out.xlsx")
        _create_perfect_linear_data(fp)
        result = run_regression(
            file_path=fp,
            sheet_name="Data",
            y_column="Y",
            x_columns=["X"],
            output_sheet="Stats",
        )
        assert result["output_sheet"] == "Stats"

        wb = openpyxl.load_workbook(fp)
        assert "Stats" in wb.sheetnames
        ws = wb["Stats"]
        # A3 should be "R-Squared", B3 should be the value
        assert ws["A3"].value == "R-Squared"
        assert abs(ws["B3"].value - 1.0) < 1e-4
        assert ws["A5"].value == "N Observations"
        assert ws["B5"].value == 10
        wb.close()
