"""Schemas for individual cell values and range reads/writes.

Includes models used by read_cell/read_range and helpers that carry style/alignment metadata.
"""

from __future__ import annotations

from typing import TypedDict

from pydantic import BaseModel, Field

from .common import CellScalar


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


class ChunkReadResult(BaseModel):
    """A paginated chunk of rows from a large dataset.

    Attributes:
        rows (list[dict[str, CellScalar]]): Rows in this chunk (required).
        chunk_start (int): 0-based starting row index (required).
        chunk_size (int): Number of rows returned (required).
        has_more (bool): True if additional rows remain (required).
        next_start_row (int | None): Next start row to request, or None if has_more is False.
    """

    rows: list[dict[str, CellScalar]] = Field(..., description="Rows in this chunk as column-keyed dicts.")
    chunk_start: int = Field(..., description="0-based starting row index of this chunk.")
    chunk_size: int = Field(..., description="Number of rows in this chunk.")
    has_more: bool = Field(..., description="Whether more rows remain after this chunk.")
    next_start_row: int | None = Field(
        None,
        description="Row number to pass as start_row to get the next chunk; None if has_more=False.",
    )
