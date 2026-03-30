from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field

from .common import BorderStyle, HorizontalAlignment, VerticalAlignment


class FormatOptions(BaseModel):
    """Cell formatting options for styling spreadsheet cells."""

    bold: bool = Field(False, description="Apply bold font weight.")
    italic: bool = Field(False, description="Apply italic font style.")
    font_size: int | None = Field(None, description="Font size in points.")
    font_color: str | None = Field(None, description="Font color as a hex string, e.g. 'FF0000'.")
    bg_color: str | None = Field(None, description="Background fill color as a hex string.")
    number_format: str | None = Field(None, description="Excel number format code, e.g. '#,##0.00'.")
    horizontal_alignment: HorizontalAlignment | None = Field(
        None, description="Horizontal alignment: left, center, right, justify."
    )
    vertical_alignment: VerticalAlignment | None = Field(None, description="Vertical alignment: top, center, bottom.")
    wrap_text: bool = Field(False, description="Enable text wrapping within the cell.")
    border_style: BorderStyle | None = Field(None, description="Border style: thin, medium, thick, double, etc.")
    border_color: str | None = Field(None, description="Border color as a hex string.")


class ConditionalFormatRule(BaseModel):
    """Rule definition for conditional formatting."""

    format_type: Literal["color_scale", "data_bar", "icon_set"] = Field(
        ..., description="Type of conditional format to apply."
    )
    start_color: str = Field("FF0000", description="Start color for gradients (hex).")
    mid_color: str = Field("FFFF00", description="Midpoint color for three-color scales (hex).")
    end_color: str = Field("00FF00", description="End color for gradients (hex).")
    icon_style: str = Field("3Arrows", description="Icon set style name for icon_set format type.")


class SortCriteria(BaseModel):
    """Sorting specification for a single column."""

    column: str = Field(..., description="Column letter or header name to sort by.")
    ascending: bool = Field(True, description="Sort in ascending order when True, descending when False.")
