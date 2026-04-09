from __future__ import annotations

from typing import Literal

import mcp_server.tools.charts as _charts
from mcp_server.models.charts import ChartInfo

__all__ = [
    "chart",
]


def chart(
    action: Literal[
        "create", "delete", "list", "add_series", "set_axes", "trendline", "combo", "data_labels", "legend", "update"
    ],
    file_path: str,
    sheet_name: str,
    chart_index: int | None = None,
    data_range: str | None = None,
    chart_type: str = "column",
    target_cell: str = "E1",
    title: str | None = None,
    x_axis_title: str = "",
    y_axis_title: str = "",
    style: int = 10,
    width: float = 15,
    height: float = 10,
    title_from_data: bool = True,
    x_title: str | None = None,
    y_title: str | None = None,
    x_min: float | None = None,
    x_max: float | None = None,
    y_min: float | None = None,
    y_max: float | None = None,
    y_number_format: str | None = None,
    log_scale_y: bool = False,
    series_index: int = 0,
    trendline_type: str = "linear",
    name: str | None = None,
    periods_forward: int = 0,
    periods_backward: int = 0,
    bar_columns: list[int] | None = None,
    line_columns: list[int] | None = None,
    anchor_cell: str = "F1",
    x_axis_column: int = 0,
    chart_title: str | None = None,
    show_value: bool = True,
    show_category: bool = False,
    show_series_name: bool = False,
    show_percentage: bool = False,
    label_position: str | None = None,
    show_legend: bool = True,
    legend_position: str | None = None,
) -> str | list[ChartInfo]:
    """Perform chart lifecycle and series/configuration operations on a worksheet.

    Args:
        action: One of "create", "delete", "list", "add_series", "set_axes", "trendline", "combo", "data_labels", "legend", "update".
            - "create": create a chart from `data_range` at `target_cell`.
            - "delete": remove a chart (optionally `chart_index`). Destructive.
            - "list": return list of charts (read-only).
            - "add_series": add a series to an existing chart (requires `chart_index`, `data_range`).
            - "set_axes": configure axis titles/ranges/number format.
            - "trendline": add a trendline to a series.
            - "combo": create a combo chart; requires `bar_columns` and `line_columns`.
            - "data_labels": configure data labels by `chart_title`.
            - "legend": configure legend by `chart_title`.
            - "update": update basic title/size/anchor of an existing chart.
        file_path: Workbook path.
        sheet_name: Worksheet name containing chart or data.
        chart_index: Optional index of target chart (default 0 or required for some actions).
        data_range: Range string used for chart creation or series.
        chart_type: Chart kind (e.g. "column", "line").
        target_cell: Anchor cell for new chart.
        ... (other visual/series parameters)

    Returns:
        str or list[ChartInfo]: Created chart id/string or list of ChartInfo for "list".

    Raises:
        ValueError: If required parameters are missing for the selected `action`.

    Notes:
        - Dispatch mapping: see function source; common destructive actions include "delete".
        - Chart creation mutates the workbook; consider documenting expected anchor and sizing units.
    """
    if action == "create":
        if data_range is None:
            raise ValueError("data_range is required for action='create'.")
        return _charts.create_chart(
            file_path,
            sheet_name,
            data_range,
            chart_type,
            target_cell,
            name or title or "",
            x_axis_title,
            y_axis_title,
            style,
            width,
            height,
        )
    if action == "delete":
        return _charts.delete_chart(file_path, sheet_name, chart_index or 0)
    if action == "list":
        return _charts.list_charts(file_path, sheet_name)
    if action == "add_series":
        if chart_index is None:
            raise ValueError("chart_index is required for action='add_series'.")
        if data_range is None:
            raise ValueError("data_range is required for action='add_series'.")
        return _charts.add_chart_series(file_path, sheet_name, chart_index, data_range, title_from_data)
    if action == "set_axes":
        return _charts.set_chart_axes(
            file_path,
            sheet_name,
            chart_index or 0,
            x_title,
            y_title,
            x_min,
            x_max,
            y_min,
            y_max,
            y_number_format,
            log_scale_y,
        )
    if action == "trendline":
        return _charts.add_chart_trendline(
            file_path,
            sheet_name,
            chart_index or 0,
            series_index,
            trendline_type,
            name,
            periods_forward,
            periods_backward,
        )
    if action == "combo":
        if not data_range:
            raise ValueError("data_range is required for action='combo'.")
        if not bar_columns:
            raise ValueError("bar_columns is required for action='combo'.")
        if not line_columns:
            raise ValueError("line_columns is required for action='combo'.")
        return _charts.create_combo_chart(
            file_path,
            sheet_name,
            data_range,
            bar_columns,
            line_columns,
            title,
            anchor_cell,
            x_axis_column,
            width,
            height,
        )
    if action == "data_labels":
        if not chart_title:
            raise ValueError("chart_title is required for action='data_labels'.")
        return _charts.set_chart_data_labels(
            file_path,
            sheet_name,
            chart_title,
            show_value,
            show_category,
            show_series_name,
            show_percentage,
            label_position,
        )
    if action == "legend":
        if not chart_title:
            raise ValueError("chart_title is required for action='legend'.")
        return _charts.set_chart_legend(file_path, sheet_name, chart_title, show_legend, legend_position)
    if action == "update":
        return _charts.update_chart(
            file_path,
            sheet_name,
            chart_index or 0,
            title,
            width if width != 15 else None,
            height if height != 10 else None,
            anchor_cell if anchor_cell != "F1" else None,
        )
    raise ValueError(f"Unknown action: {action}")


def register(mcp) -> None:
    """Register tools on *mcp*."""
    mcp.tool()(chart)
