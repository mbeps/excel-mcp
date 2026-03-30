from __future__ import annotations

from typing import TypedDict

from pydantic import BaseModel, Field


class SheetInfo(BaseModel):
    """Basic dimensional info for a single worksheet."""

    name: str = Field(..., description="Name of the worksheet.")
    min_row: int | None = Field(None, description="First used row (1-based), or None if empty.")
    max_row: int | None = Field(None, description="Last used row (1-based), or None if empty.")
    min_col: int | None = Field(None, description="First used column (1-based), or None if empty.")
    max_col: int | None = Field(None, description="Last used column (1-based), or None if empty.")


class SheetSummary(BaseModel):
    """High-level summary of a worksheet's contents."""

    name: str = Field(..., description="Name of the worksheet.")
    row_count: int = Field(..., description="Total number of used rows.")
    col_count: int = Field(..., description="Total number of used columns.")
    headers: list[str] = Field(..., description="Column header values from the first row.")
    used_range: str = Field(..., description="Used cell range in A1 notation, e.g. 'A1:D10'.")


class WorkbookCreatedResult(BaseModel):
    """Confirmation returned after creating a new workbook."""

    file_path: str = Field(..., description="Absolute path of the created workbook file.")
    sheets: list[str] = Field(..., description="Names of sheets in the new workbook.")


class WorkbookMetadata(BaseModel):
    """Full metadata for an Excel workbook."""

    file_path: str = Field(..., description="Absolute path to the workbook file.")
    sheets: list[SheetInfo] = Field(..., description="Dimensional info for each sheet.")
    active_sheet: str | None = Field(..., description="Name of the currently active sheet.")
    named_ranges: list[dict[str, str]] = Field(..., description="Named ranges defined in the workbook.")


class SheetCreatedInfo(TypedDict):
    """Info about a sheet created by write_multi_sheet."""

    name: str
    header_count: int
    row_count: int
    column_widths_set: list[str]


class WriteMultiSheetResult(TypedDict):
    """Result returned by write_multi_sheet."""

    file_path: str
    sheets_created: list[SheetCreatedInfo]


class ValidationRangeResult(TypedDict, total=False):
    """Result dict from validate_excel_range / _validate_single_cell."""

    valid: bool
    message: str
    start_cell: str
    end_cell: str | None
    is_range: bool
    column: str
    row: int
    col_index: int
