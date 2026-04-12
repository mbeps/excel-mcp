"""Formatting and conditional-formatting option models.

Used by formatting tools to accept styling requests and describe conditional formatting rules.
"""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field

from .common import BorderStyle, HorizontalAlignment, VerticalAlignment


class FormatOptions(BaseModel):
    """Cell formatting options for styling spreadsheet cells.

    Attributes:
        bold (bool): Apply bold font weight. Default False.
        italic (bool): Apply italic font style. Default False.
        font_size (int | None): Font size in points.
        font_color (str | None): Font color as hex string, e.g. 'FF0000'.
        bg_color (str | None): Background fill color as hex string.
        number_format (str | None): Excel number format code.
        horizontal_alignment (HorizontalAlignment | None): Optional horizontal alignment.
        vertical_alignment (VerticalAlignment | None): Optional vertical alignment.
        wrap_text (bool): Enable text wrapping. Default False.
        border_style (BorderStyle | None): Optional border style.
        border_color (str | None): Border color as hex string.

    Notes:
        - Color strings should be hex without a leading '#', matching openpyxl conventions.
    """

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
    """Rule definition for conditional formatting.

    Attributes:
        format_type (Literal): One of 'color_scale', 'data_bar', 'icon_set'. Required.
        start_color, mid_color, end_color (str): Hex color strings used for gradients.
        icon_style (str): Icon set style name used when `format_type == 'icon_set'`.
    """

    format_type: Literal["color_scale", "data_bar", "icon_set"] = Field(
        ..., description="Type of conditional format to apply."
    )
    start_color: str = Field("FF0000", description="Start color for gradients (hex).")
    mid_color: str = Field("FFFF00", description="Midpoint color for three-color scales (hex).")
    end_color: str = Field("00FF00", description="End color for gradients (hex).")
    icon_style: str = Field("3Arrows", description="Icon set style name for icon_set format type.")


class SortCriteria(BaseModel):
    """Sorting specification for a single column.

    Attributes:
        column (str): Column letter or header name to sort by. Required.
        ascending (bool): True = ascending, False = descending. Default True.
    """

    column: str = Field(..., description="Column letter or header name to sort by.")
    ascending: bool = Field(True, description="Sort in ascending order when True, descending when False.")
