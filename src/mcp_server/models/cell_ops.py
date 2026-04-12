"""Schemas for individual cell values and range reads/writes.

Includes models used by read_cell/read_range and helpers that carry style/alignment metadata.
"""

from __future__ import annotations

from typing import TypedDict

from pydantic import BaseModel, Field

from .common import CellScalar


class CellValue(BaseModel):
    """Value and metadata for a single cell.

    Attributes:
        cell_ref (str): Cell reference in A1 notation, e.g. 'B3'. Required.
        value (CellScalar): Current value of the cell. May be None for empty cells.
        data_type (str): Data type code reported by openpyxl (e.g. 's', 'n', 'd', 'b'). Required.

    Notes:
        - This model is used as a response schema for single-cell read operations.
    """

    cell_ref: str = Field(..., description="Cell reference in A1 notation, e.g. 'B3'.")
    value: CellScalar = Field(..., description="Current value of the cell.")
    data_type: str = Field(..., description="Data type code reported by openpyxl (s, n, d, b, etc.).")


class RangeData(BaseModel):
    """Raw grid data read from a rectangular range.

    Attributes:
        rows (list[list[CellScalar]]): Row-major list of cell values.
        row_count (int): Number of rows returned. Required.
        col_count (int): Number of columns returned. Required.

    Notes:
        - row_count and col_count must match the dimensions of `rows` when present.
    """

    rows: list[list[CellScalar]] = Field(..., description="Row-major list of cell values.")
    row_count: int = Field(..., description="Number of rows returned.")
    col_count: int = Field(..., description="Number of columns returned.")


class CellAlignmentInfo(TypedDict):
    """Alignment settings extracted from a cell's style.

    Keys:
        horizontal (str | None): Horizontal alignment name or None.
        vertical (str | None): Vertical alignment name or None.
        wrap_text (bool | None): Whether text wrapping is enabled.
    """

    horizontal: str | None
    vertical: str | None
    wrap_text: bool | None


class CellStyleInfo(TypedDict):
    """Cell style information returned when `show_style=True` in read_range.

    Keys:
        font_name (str | None): Font family name.
        font_size (float | None): Font size in points.
        bold (bool | None): Whether bold is set.
        italic (bool | None): Whether italic is set.
        font_color (str | None): Font color as hex string.
        fill_color (str | None): Background fill color as hex string.
        number_format (str): Excel number format code.
        alignment (CellAlignmentInfo): Nested alignment dict.
    """

    font_name: str | None
    font_size: float | None
    bold: bool | None
    italic: bool | None
    font_color: str | None
    fill_color: str | None
    number_format: str
    alignment: CellAlignmentInfo
