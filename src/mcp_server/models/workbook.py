"""Workbook- and sheet-level metadata models returned by workbook and worksheet tools."""

from __future__ import annotations

from typing import TypedDict

from pydantic import BaseModel, Field


class SheetInfo(BaseModel):
    """Basic dimensional info for a single worksheet.

    Attributes:
        name (str): Worksheet name. Required.
        min_row (int | None): First used row (1-based), or None if empty.
        max_row (int | None): Last used row (1-based), or None if empty.
        min_col (int | None): First used column (1-based), or None if empty.
        max_col (int | None): Last used column (1-based), or None if empty.
    """

    name: str = Field(..., description="Name of the worksheet.")
    min_row: int | None = Field(None, description="First used row (1-based), or None if empty.")
    max_row: int | None = Field(None, description="Last used row (1-based), or None if empty.")
    min_col: int | None = Field(None, description="First used column (1-based), or None if empty.")
    max_col: int | None = Field(None, description="Last used column (1-based), or None if empty.")


class SheetSummary(BaseModel):
    """High-level summary of a worksheet's contents.

    Attributes:
        name (str): Worksheet name. Required.
        row_count (int): Number of used rows. Required.
        col_count (int): Number of used columns. Required.
        headers (list[str]): Column headers from the first row. Required.
        used_range (str): Used range in A1 notation (e.g. 'A1:D10'). Required.
    """

    name: str = Field(..., description="Name of the worksheet.")
    row_count: int = Field(..., description="Total number of used rows.")
    col_count: int = Field(..., description="Total number of used columns.")
    headers: list[str] = Field(..., description="Column header values from the first row.")
    used_range: str = Field(..., description="Used cell range in A1 notation, e.g. 'A1:D10'.")


class WorkbookCreatedResult(BaseModel):
    """Confirmation returned after creating a new workbook.

    Attributes:
        file_path (str): Absolute path of the created workbook. Required.
        sheets (list[str]): Names of sheets in the new workbook. Required.
    """

    file_path: str = Field(..., description="Absolute path of the created workbook file.")
    sheets: list[str] = Field(..., description="Names of sheets in the new workbook.")


class WorkbookMetadata(BaseModel):
    """Full metadata for an Excel workbook.

    Attributes:
        file_path (str): Absolute path to the workbook file. Required.
        sheets (list[SheetInfo]): Dimensional info for each sheet. Required.
        active_sheet (str | None): Name of the active sheet.
        named_ranges (list[dict[str, str]]): Named ranges present in the workbook.
    """

    file_path: str = Field(..., description="Absolute path to the workbook file.")
    sheets: list[SheetInfo] = Field(..., description="Dimensional info for each sheet.")
    active_sheet: str | None = Field(..., description="Name of the currently active sheet.")
    named_ranges: list[dict[str, str]] = Field(..., description="Named ranges defined in the workbook.")


class SheetCreatedInfo(TypedDict):
    """Info about a sheet created by `write_multi_sheet`.

    Keys:
        name (str): Sheet name. Required.
        header_count (int): Number of header rows written.
        row_count (int): Number of data rows written.
        column_widths_set (list[str]): Columns for which widths were set.
    """

    name: str
    header_count: int
    row_count: int
    column_widths_set: list[str]


class WriteMultiSheetResult(TypedDict):
    """Result returned by write_multi_sheet.

    Keys:
        file_path (str): Path to the workbook. Required.
        sheets_created (list[SheetCreatedInfo]): Metadata for sheets written.
    """

    file_path: str
    sheets_created: list[SheetCreatedInfo]


class _SheetDefinitionBase(TypedDict):
    name: str


class SheetDefinition(_SheetDefinitionBase, total=False):
    """Input definition for a single sheet used by write_multi_sheet.

    Keys (total=False):
        name (str): Required base key in _SheetDefinitionBase.
        headers (list[str]): Optional header row values.
        data, values (list[list[object]]): Row data; choose one of data or values depending on API.
        column_widths (dict[str, float]): Optional column width map.
    """

    headers: list[str]
    data: list[list[object]]
    values: list[list[object]]
    column_widths: dict[str, float]


class ValidationRangeResult(TypedDict, total=False):
    """Result dictionary produced when validating an Excel range.

    Keys (total=False):
        valid (bool): Whether the range is valid.
        message (str): Informational message when invalid.
        start_cell (str): Start cell reference.
        end_cell (str | None): End cell if range.
        is_range (bool): True when a multi-cell range.
        column (str): Column letter for a single-cell validation.
        row (int): Row number for a single-cell validation.
        col_index (int): 1-based numeric column index.
    """

    valid: bool
    message: str
    start_cell: str
    end_cell: str | None
    is_range: bool
    column: str
    row: int
    col_index: int
