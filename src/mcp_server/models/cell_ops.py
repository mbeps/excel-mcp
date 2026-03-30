from __future__ import annotations

from typing import TypedDict

from pydantic import BaseModel, Field

from .common import CellScalar


class CellValue(BaseModel):
    """Value and metadata for a single cell."""

    cell_ref: str = Field(..., description="Cell reference in A1 notation, e.g. 'B3'.")
    value: CellScalar = Field(..., description="Current value of the cell.")
    data_type: str = Field(..., description="Data type code reported by openpyxl (s, n, d, b, etc.).")


class RangeData(BaseModel):
    """Raw grid data read from a rectangular range."""

    rows: list[list[CellScalar]] = Field(..., description="Row-major list of cell values.")
    row_count: int = Field(..., description="Number of rows returned.")
    col_count: int = Field(..., description="Number of columns returned.")


class CellAlignmentInfo(TypedDict):
    """Alignment settings extracted from a cell's style."""

    horizontal: str | None
    vertical: str | None
    wrap_text: bool | None


class CellStyleInfo(TypedDict):
    """Full style/metadata for a cell, returned by read_range with show_style=True."""

    font_name: str | None
    font_size: float | None
    bold: bool | None
    italic: bool | None
    font_color: str | None
    fill_color: str | None
    number_format: str
    alignment: CellAlignmentInfo
