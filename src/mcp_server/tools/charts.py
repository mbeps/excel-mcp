"""Chart operations: create, delete, and list Excel charts."""

from __future__ import annotations

from logging import Logger

from openpyxl.chart import (
    AreaChart,
    BarChart,
    BubbleChart,
    DoughnutChart,
    LineChart,
    PieChart,
    RadarChart,
    Reference,
    ScatterChart,
    StockChart,
)
from openpyxl.chart.legend import Legend
from openpyxl.utils import range_boundaries

from mcp_server.utils.excel_helpers import get_sheet, load_workbook_safe, save_workbook_safe
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

CHART_TYPES = {"bar", "column", "line", "pie", "scatter", "area", "radar", "doughnut", "bubble", "stock"}


def _make_chart(
    chart_type: str,
) -> BarChart | LineChart | PieChart | ScatterChart | AreaChart | RadarChart | DoughnutChart | BubbleChart | StockChart:
    if chart_type == "bar":
        return BarChart(type="bar")
    if chart_type == "column":
        return BarChart(type="col")
    if chart_type == "line":
        return LineChart()
    if chart_type == "pie":
        return PieChart()
    if chart_type == "scatter":
        return ScatterChart()
    if chart_type == "area":
        return AreaChart()
    if chart_type == "radar":
        return RadarChart()
    if chart_type == "doughnut":
        return DoughnutChart()
    if chart_type == "bubble":
        return BubbleChart()
    if chart_type == "stock":
        return StockChart()
    raise ValueError(f"Unsupported chart type '{chart_type}'. Allowed: {CHART_TYPES}")


def create_chart(
    file_path: str,
    sheet_name: str,
    data_range: str,
    chart_type: str = "column",
    target_cell: str = "E1",
    title: str = "",
    x_axis_title: str = "",
    y_axis_title: str = "",
    style: int = 10,
    width: float = 15,
    height: float = 10,
) -> str:
    """Create a native Excel chart from a data range."""
    if chart_type not in CHART_TYPES:
        raise ValueError(f"Unsupported chart type '{chart_type}'. Allowed: {CHART_TYPES}")

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        min_col, min_row, max_col, max_row = range_boundaries(data_range)

        chart = _make_chart(chart_type)
        chart.title = title or None
        chart.style = style
        chart.width = width
        chart.height = height

        if chart_type == "scatter":
            x_values = Reference(ws, min_col=min_col, min_row=min_row + 1, max_row=max_row)
            for col_idx in range(min_col + 1, max_col + 1):
                y_values = Reference(ws, min_col=col_idx, min_row=min_row, max_row=max_row)
                chart.add_data(y_values, titles_from_data=True)
                chart.series[-1].xvalues = x_values
        elif chart_type == "bubble":
            # Bubble chart needs x, y, size columns (at least 3 data columns)
            x_values = Reference(ws, min_col=min_col, min_row=min_row + 1, max_row=max_row)
            y_values = Reference(ws, min_col=min_col + 1, min_row=min_row, max_row=max_row)
            size_values = Reference(ws, min_col=min_col + 2, min_row=min_row + 1, max_row=max_row)
            chart.add_data(y_values, titles_from_data=True)
            chart.series[0].xvalues = x_values
            chart.series[0].zvalues = size_values
        elif chart_type == "stock":
            # Stock chart expects Open/High/Low/Close data in columns
            data = Reference(ws, min_col=min_col + 1, min_row=min_row, max_row=max_row, max_col=max_col)
            categories = Reference(ws, min_col=min_col, min_row=min_row + 1, max_row=max_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(categories)
        else:
            data = Reference(ws, min_col=min_col + 1, min_row=min_row, max_row=max_row, max_col=max_col)
            categories = Reference(ws, min_col=min_col, min_row=min_row + 1, max_row=max_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(categories)

        no_axis_types = {"pie", "doughnut"}
        if chart_type not in no_axis_types and x_axis_title:
            chart.x_axis.title = x_axis_title
        if chart_type not in no_axis_types and y_axis_title:
            chart.y_axis.title = y_axis_title

        ws.add_chart(chart, target_cell)
        save_workbook_safe(wb, file_path)
        logger.info("Created %s chart at %s!%s", chart_type, sheet_name, target_cell)
        return f"Created {chart_type} chart at {target_cell} in '{sheet_name}'."
    finally:
        wb.close()


def delete_chart(file_path: str, sheet_name: str, chart_index: int = 0) -> str:
    """Delete a chart from a sheet by index."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        charts = ws._charts
        if not charts:
            raise ValueError(f"No charts found in sheet '{sheet_name}'.")
        if chart_index < 0 or chart_index >= len(charts):
            raise ValueError(f"Chart index {chart_index} out of range (0-{len(charts) - 1}).")

        removed = charts.pop(chart_index)
        save_workbook_safe(wb, file_path)
        title = removed.title or "(untitled)"
        logger.info("Deleted chart %d ('%s') from %s", chart_index, title, sheet_name)
        return f"Deleted chart {chart_index} ('{title}') from '{sheet_name}'."
    finally:
        wb.close()


def list_charts(file_path: str, sheet_name: str) -> list[dict]:
    """List all charts on a sheet with title, type, and position."""
    wb = load_workbook_safe(file_path, read_only=False)
    try:
        ws = get_sheet(wb, sheet_name)
        result = []
        for chart in ws._charts:
            result.append(
                {
                    "title": chart.title or "(untitled)",
                    "type": type(chart).__name__,
                    "position": (
                        getattr(chart, "anchor", None) and str(chart.anchor._from)
                        if hasattr(chart, "anchor")
                        else "unknown"
                    ),
                }
            )
        return result
    finally:
        wb.close()


def update_chart_properties(
    file_path: str,
    sheet_name: str,
    chart_index: int = 0,
    title: str | None = None,
    x_axis_title: str | None = None,
    y_axis_title: str | None = None,
    legend_position: str | None = None,
    show_legend: bool | None = None,
    style: int | None = None,
) -> str:
    """Update properties of an existing chart."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        charts = ws._charts
        if not charts:
            raise ValueError(f"No charts found in sheet '{sheet_name}'.")
        if chart_index < 0 or chart_index >= len(charts):
            raise ValueError(f"Chart index {chart_index} out of range (0-{len(charts) - 1}).")

        chart = charts[chart_index]
        updated: list[str] = []

        if title is not None:
            chart.title = title or None
            updated.append("title")
        if style is not None:
            chart.style = style
            updated.append("style")
        if x_axis_title is not None and hasattr(chart, "x_axis"):
            chart.x_axis.title = x_axis_title
            updated.append("x_axis_title")
        if y_axis_title is not None and hasattr(chart, "y_axis"):
            chart.y_axis.title = y_axis_title
            updated.append("y_axis_title")
        if show_legend is False:
            chart.legend = None
            updated.append("legend (hidden)")
        elif legend_position is not None:
            chart.legend = Legend()
            chart.legend.position = legend_position
            updated.append(f"legend_position={legend_position}")

        save_workbook_safe(wb, file_path)
        logger.info("Updated chart %d in %s: %s", chart_index, sheet_name, updated)
        return f"Updated chart {chart_index} in '{sheet_name}': {', '.join(updated)}."
    finally:
        wb.close()
