"""Chart configuration model and metadata types used by chart tools."""

from __future__ import annotations

from typing import Literal, TypedDict

from pydantic import BaseModel, Field


class ChartConfig(BaseModel):
    """Configuration for inserting a chart into a worksheet.

    Attributes:
        chart_type (Literal): One of 'bar', 'column', 'line', 'pie', 'scatter', 'area'. Required.
        title (str): Chart title.
        x_axis_title, y_axis_title (str): Axis labels.
        style (int): Built-in Excel chart style number.
        width (float): Chart width in cm.
        height (float): Chart height in cm.
    """

    chart_type: Literal["bar", "column", "line", "pie", "scatter", "area"] = Field(
        ..., description="Type of chart to create."
    )
    title: str = Field("", description="Chart title text.")
    x_axis_title: str = Field("", description="Label for the X axis.")
    y_axis_title: str = Field("", description="Label for the Y axis.")
    style: int = Field(10, description="Built-in Excel chart style number.")
    width: float = Field(15, description="Chart width in cm.")
    height: float = Field(10, description="Chart height in cm.")


class ChartInfo(TypedDict):
    """Info about a single chart on a worksheet returned by list operations.

    Keys:
        title (str): Chart title.
        type (str): Chart type string.
        position (str | None): Anchor or position description, if available.
    """

    title: str
    type: str
    position: str | None
