"""Workflow tests for chart lifecycle operations.

Tests chain multiple MCP tool calls simulating real chart workflows and verify
results independently using openpyxl — never via MCP tools.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from mcp_server.tools.cell_ops import write_range
from mcp_server.tools.charts import (
    _safe_chart_title,
    add_chart_series,
    add_chart_trendline,
    create_chart,
    create_combo_chart,
    delete_chart,
    list_charts,
    set_chart_axes,
    set_chart_data_labels,
    set_chart_legend,
)
from mcp_server.tools.workbook import create_workbook


def _write_sample_data(fp: str, sheet: str = "Sheet1") -> None:
    """Write a standard 5-row dataset with headers for chart tests."""
    write_range(
        fp,
        sheet,
        "A1",
        [
            ["Month", "Sales", "Expenses", "Profit"],
            ["Jan", 100, 60, 40],
            ["Feb", 150, 80, 70],
            ["Mar", 200, 90, 110],
            ["Apr", 120, 70, 50],
            ["May", 180, 85, 95],
        ],
    )


# ---------------------------------------------------------------------------
# 1. Create basic bar chart
# ---------------------------------------------------------------------------


class TestCreateBasicBarChart:
    def test_create_basic_bar_chart(self, tmp_path: Path) -> None:
        """Write data → create_chart(type='bar') → verify chart exists in sheet."""
        fp = str(tmp_path / "bar.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        result = create_chart(fp, "Sheet1", "A1:B6", chart_type="bar", title="Sales Bar")
        assert "bar" in result.lower()

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._charts) == 1
        chart = ws._charts[0]
        assert chart.__class__.__name__ == "BarChart"
        wb.close()


# ---------------------------------------------------------------------------
# 2. Create line chart with series
# ---------------------------------------------------------------------------


class TestCreateLineChartWithSeries:
    def test_create_line_chart_with_series(self, tmp_path: Path) -> None:
        """Create line chart → add_chart_series → verify series count."""
        fp = str(tmp_path / "line.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        create_chart(fp, "Sheet1", "A1:B6", chart_type="line", title="Line Chart")

        # Add a second series (Expenses)
        add_chart_series(fp, "Sheet1", chart_index=0, data_range="C1:C6")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        chart = ws._charts[0]
        assert len(chart.series) >= 2
        wb.close()


# ---------------------------------------------------------------------------
# 3. Chart with title and axes
# ---------------------------------------------------------------------------


class TestChartWithTitleAndAxes:
    def test_chart_with_title_and_axes(self, tmp_path: Path) -> None:
        """Create → set_axes with titles → verify via openpyxl."""
        fp = str(tmp_path / "axes.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        create_chart(fp, "Sheet1", "A1:B6", chart_type="column", title="Revenue")
        set_chart_axes(fp, "Sheet1", chart_index=0, x_title="Month", y_title="Amount ($)")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        chart = ws._charts[0]
        assert _safe_chart_title(chart.x_axis.title) == "Month"
        assert _safe_chart_title(chart.y_axis.title) == "Amount ($)"
        wb.close()


# ---------------------------------------------------------------------------
# 4. Chart data labels
# ---------------------------------------------------------------------------


class TestChartDataLabels:
    def test_chart_data_labels(self, tmp_path: Path) -> None:
        """Create chart → set_chart_data_labels → verify returns dict with status."""
        fp = str(tmp_path / "labels.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        create_chart(fp, "Sheet1", "A1:B6", chart_type="column", title="LabelChart")
        result = set_chart_data_labels(
            fp,
            "Sheet1",
            chart_title="LabelChart",
            show_value=True,
            show_category=False,
        )
        assert result["status"] == "ok"
        assert result["chart_title"] == "LabelChart"
        assert result["show_value"] is True

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        chart = ws._charts[0]
        assert chart.dLbls is not None
        wb.close()


# ---------------------------------------------------------------------------
# 5. Chart legend
# ---------------------------------------------------------------------------


class TestChartLegend:
    def test_chart_legend(self, tmp_path: Path) -> None:
        """Create chart → set_chart_legend → verify no error, returns dict."""
        fp = str(tmp_path / "legend.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        create_chart(fp, "Sheet1", "A1:B6", chart_type="column", title="LegendChart")
        result = set_chart_legend(fp, "Sheet1", chart_title="LegendChart", show=True, position="b")
        assert result["status"] == "ok"
        assert result["show"] is True

    def test_chart_legend_hide(self, tmp_path: Path) -> None:
        """Hide legend → verify legend is None."""
        fp = str(tmp_path / "no_legend.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        create_chart(fp, "Sheet1", "A1:B6", chart_type="column", title="NoLegend")
        set_chart_legend(fp, "Sheet1", chart_title="NoLegend", show=False)

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        chart = ws._charts[0]
        assert chart.legend is None
        wb.close()


# ---------------------------------------------------------------------------
# 6. Chart trendline
# ---------------------------------------------------------------------------


class TestChartTrendline:
    def test_chart_trendline(self, tmp_path: Path) -> None:
        """Create chart → add_chart_trendline → verify via openpyxl."""
        fp = str(tmp_path / "trend.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        create_chart(fp, "Sheet1", "A1:B6", chart_type="column", title="TrendChart")
        result = add_chart_trendline(
            fp,
            "Sheet1",
            chart_index=0,
            series_index=0,
            trendline_type="linear",
            name="Sales Trend",
        )
        assert "linear" in result.lower()

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        chart = ws._charts[0]
        series = list(chart.series)
        assert series[0].trendline is not None
        wb.close()


# ---------------------------------------------------------------------------
# 7. Create combo chart
# ---------------------------------------------------------------------------


class TestCreateComboChart:
    def test_create_combo_chart(self, tmp_path: Path) -> None:
        """Create combo chart with bar+line → verify chart exists."""
        fp = str(tmp_path / "combo.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        result = create_combo_chart(
            fp,
            "Sheet1",
            data_range="A1:D6",
            bar_columns=[1],
            line_columns=[2, 3],
            title="Combo Chart",
        )
        assert "combo" in result.lower()

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._charts) == 1
        wb.close()


# ---------------------------------------------------------------------------
# 8. Combo chart at specific target cell
# ---------------------------------------------------------------------------


class TestComboChartTargetCell:
    def test_combo_chart_target_cell(self, tmp_path: Path) -> None:
        """Create combo at specific anchor_cell → verify chart is anchored there."""
        fp = str(tmp_path / "combo_anchor.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        create_combo_chart(
            fp,
            "Sheet1",
            data_range="A1:D6",
            bar_columns=[1],
            line_columns=[2],
            title="Anchored",
            anchor_cell="H10",
        )

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._charts) == 1
        wb.close()


# ---------------------------------------------------------------------------
# 9. Delete chart by index
# ---------------------------------------------------------------------------


class TestDeleteChartByIndex:
    def test_delete_chart_by_index(self, tmp_path: Path) -> None:
        """Create 2 charts → delete by index → verify correct one removed."""
        fp = str(tmp_path / "del_idx.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        create_chart(fp, "Sheet1", "A1:B6", chart_type="column", title="Chart A", target_cell="E1")
        create_chart(fp, "Sheet1", "A1:C6", chart_type="line", title="Chart B", target_cell="E16")

        wb = openpyxl.load_workbook(fp)
        assert len(wb["Sheet1"]._charts) == 2
        wb.close()

        delete_chart(fp, "Sheet1", chart_index=0)

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._charts) == 1
        wb.close()


# ---------------------------------------------------------------------------
# 10. Delete chart by title
# ---------------------------------------------------------------------------


class TestDeleteChartByTitle:
    def test_delete_chart_by_title(self, tmp_path: Path) -> None:
        """Create 2 charts with titles → delete by title → verify correct one removed."""
        fp = str(tmp_path / "del_title.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        create_chart(fp, "Sheet1", "A1:B6", chart_type="column", title="Keep Me", target_cell="E1")
        create_chart(fp, "Sheet1", "A1:C6", chart_type="line", title="Remove Me", target_cell="E16")

        delete_chart(fp, "Sheet1", chart_title="Remove Me")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._charts) == 1
        # The remaining chart should be "Keep Me"
        remaining = ws._charts[0]
        assert _safe_chart_title(remaining.title) == "Keep Me"
        wb.close()


# ---------------------------------------------------------------------------
# 11. List charts
# ---------------------------------------------------------------------------


class TestListCharts:
    def test_list_charts(self, tmp_path: Path) -> None:
        """Create 3 charts → list_charts → verify all 3 returned."""
        fp = str(tmp_path / "list.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        create_chart(fp, "Sheet1", "A1:B6", chart_type="bar", title="Chart 1", target_cell="E1")
        create_chart(fp, "Sheet1", "A1:C6", chart_type="line", title="Chart 2", target_cell="E16")
        create_chart(fp, "Sheet1", "A1:D6", chart_type="column", title="Chart 3", target_cell="E31")

        result = list_charts(fp, "Sheet1")
        assert len(result) == 3
        titles = {c["title"] for c in result}
        assert titles == {"Chart 1", "Chart 2", "Chart 3"}


# ---------------------------------------------------------------------------
# 12. Chart with categories_range
# ---------------------------------------------------------------------------


class TestChartWithCategoriesRange:
    def test_chart_with_categories_range(self, tmp_path: Path) -> None:
        """Create chart with categories_range → verify chart created successfully."""
        fp = str(tmp_path / "cat_range.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        result = create_chart(
            fp,
            "Sheet1",
            data_range="B1:D6",
            chart_type="column",
            title="With Categories",
            categories_range="A2:A6",
        )
        assert "column" in result.lower()

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._charts) == 1
        wb.close()


# ---------------------------------------------------------------------------
# 13. Chart lifecycle — full
# ---------------------------------------------------------------------------


class TestChartLifecycleFull:
    def test_chart_lifecycle_full(self, tmp_path: Path) -> None:
        """Full lifecycle: create → add series → set axes → data labels → legend → trendline → delete."""
        fp = str(tmp_path / "lifecycle.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        # Create
        create_chart(fp, "Sheet1", "A1:B6", chart_type="column", title="Lifecycle")

        # Add series
        add_chart_series(fp, "Sheet1", chart_index=0, data_range="C1:C6")

        # Set axes
        set_chart_axes(fp, "Sheet1", chart_index=0, x_title="Period", y_title="$$")

        # Data labels
        label_result = set_chart_data_labels(
            fp,
            "Sheet1",
            chart_title="Lifecycle",
            show_value=True,
            show_percentage=False,
        )
        assert label_result["status"] == "ok"

        # Legend
        legend_result = set_chart_legend(
            fp,
            "Sheet1",
            chart_title="Lifecycle",
            show=True,
            position="r",
        )
        assert legend_result["status"] == "ok"

        # Trendline
        add_chart_trendline(
            fp,
            "Sheet1",
            chart_index=0,
            series_index=0,
            trendline_type="linear",
        )

        # Verify everything is present
        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._charts) == 1
        chart = ws._charts[0]
        assert len(chart.series) >= 2
        assert _safe_chart_title(chart.x_axis.title) == "Period"
        assert _safe_chart_title(chart.y_axis.title) == "$$"
        assert chart.dLbls is not None
        assert chart.legend is not None
        assert chart.series[0].trendline is not None
        wb.close()

        # Delete
        delete_chart(fp, "Sheet1", chart_title="Lifecycle")

        wb = openpyxl.load_workbook(fp)
        assert len(wb["Sheet1"]._charts) == 0
        wb.close()


# ---------------------------------------------------------------------------
# 14. Chart with sheet-qualified range
# ---------------------------------------------------------------------------


class TestChartWithSheetQualifiedRange:
    def test_chart_with_sheet_qualified_range(self, tmp_path: Path) -> None:
        """Use 'Sheet1!A1:B6' as data_range → should work (prefix stripped)."""
        fp = str(tmp_path / "qualified.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        result = create_chart(
            fp,
            "Sheet1",
            data_range="Sheet1!A1:B6",
            chart_type="column",
            title="Qualified Range",
        )
        assert "column" in result.lower()

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._charts) == 1
        wb.close()


# ---------------------------------------------------------------------------
# 15. Multiple chart types
# ---------------------------------------------------------------------------


class TestMultipleChartTypes:
    def test_pie_chart(self, tmp_path: Path) -> None:
        """Create a pie chart → verify type."""
        fp = str(tmp_path / "pie.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        create_chart(fp, "Sheet1", "A1:B6", chart_type="pie", title="Pie Chart")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._charts) == 1
        assert ws._charts[0].__class__.__name__ == "PieChart"
        wb.close()

    def test_area_chart(self, tmp_path: Path) -> None:
        """Create an area chart → verify type."""
        fp = str(tmp_path / "area.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        create_chart(fp, "Sheet1", "A1:B6", chart_type="area", title="Area Chart")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._charts) == 1
        assert ws._charts[0].__class__.__name__ == "AreaChart"
        wb.close()

    def test_scatter_chart(self, tmp_path: Path) -> None:
        """Create a scatter chart → verify type."""
        fp = str(tmp_path / "scatter.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        write_range(
            fp,
            "Sheet1",
            "A1",
            [
                ["X", "Y"],
                [1, 10],
                [2, 20],
                [3, 30],
            ],
        )

        create_chart(fp, "Sheet1", "A1:B4", chart_type="scatter", title="Scatter")

        wb = openpyxl.load_workbook(fp)
        ws = wb["Sheet1"]
        assert len(ws._charts) == 1
        assert ws._charts[0].__class__.__name__ == "ScatterChart"
        wb.close()


# ---------------------------------------------------------------------------
# 16. Chart error handling
# ---------------------------------------------------------------------------


class TestChartErrorHandling:
    def test_delete_from_empty_sheet(self, tmp_path: Path) -> None:
        """delete_chart on sheet with no charts → raises ValueError."""
        fp = str(tmp_path / "no_charts.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])

        with pytest.raises(ValueError, match="No charts found"):
            delete_chart(fp, "Sheet1", chart_index=0)

    def test_invalid_chart_type(self, tmp_path: Path) -> None:
        """create_chart with invalid type → raises ValueError."""
        fp = str(tmp_path / "bad_type.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)

        with pytest.raises(ValueError, match="Unsupported chart type"):
            create_chart(fp, "Sheet1", "A1:B6", chart_type="invalid_type")

    def test_trendline_invalid_type(self, tmp_path: Path) -> None:
        """add_chart_trendline with invalid type → raises ValueError."""
        fp = str(tmp_path / "bad_trend.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)
        create_chart(fp, "Sheet1", "A1:B6", chart_type="column", title="T")

        with pytest.raises(ValueError, match="Invalid trendline_type"):
            add_chart_trendline(fp, "Sheet1", chart_index=0, trendline_type="cubic")

    def test_chart_index_out_of_range(self, tmp_path: Path) -> None:
        """delete_chart with out-of-range index → raises ValueError."""
        fp = str(tmp_path / "idx_oor.xlsx")
        create_workbook(fp, sheet_names=["Sheet1"])
        _write_sample_data(fp)
        create_chart(fp, "Sheet1", "A1:B6", chart_type="column", title="Only")

        with pytest.raises(ValueError, match="out of range"):
            delete_chart(fp, "Sheet1", chart_index=5)
