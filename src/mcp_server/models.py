"""Pydantic models for the Excel MCP server."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

# ── Response Models ──────────────────────────────────────────────────────────


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


class CellValue(BaseModel):
    """Value and metadata for a single cell."""

    cell_ref: str = Field(..., description="Cell reference in A1 notation, e.g. 'B3'.")
    value: Any = Field(..., description="Current value of the cell.")
    data_type: str = Field(..., description="Data type code reported by openpyxl (s, n, d, b, etc.).")


class RangeData(BaseModel):
    """Raw grid data read from a rectangular range."""

    rows: list[list[Any]] = Field(..., description="Row-major list of cell values.")
    row_count: int = Field(..., description="Number of rows returned.")
    col_count: int = Field(..., description="Number of columns returned.")


class FilterResult(BaseModel):
    """Rows that matched a filter condition."""

    matched_rows: list[list[Any]] = Field(..., description="Rows matching the filter criteria.")
    total_rows: int = Field(..., description="Total number of rows evaluated.")
    matched_count: int = Field(..., description="Number of rows that matched.")
    headers: list[str] = Field(..., description="Column headers for the returned rows.")


class ColumnStats(BaseModel):
    """Descriptive statistics for a single numeric column."""

    column: str = Field(..., description="Column letter or header name.")
    count: int = Field(..., description="Number of non-empty numeric values.")
    mean: float | None = Field(None, description="Arithmetic mean, or None if unavailable.")
    median: float | None = Field(None, description="Median value, or None if unavailable.")
    min_val: float | None = Field(None, description="Minimum value, or None if unavailable.")
    max_val: float | None = Field(None, description="Maximum value, or None if unavailable.")
    std: float | None = Field(None, description="Standard deviation, or None if unavailable.")
    sum_val: float | None = Field(None, description="Sum of values, or None if unavailable.")


class AggregateResult(BaseModel):
    """Result of a group-by aggregation operation."""

    groups: list[dict[str, Any]] = Field(..., description="List of group dictionaries with keys and aggregated values.")
    group_by: str = Field(..., description="Column used for grouping.")
    operation: str = Field(..., description="Aggregation operation applied (sum, mean, count, etc.).")


class DuplicateResult(BaseModel):
    """Rows identified as duplicates."""

    duplicates: list[list[Any]] = Field(..., description="Duplicate rows found.")
    count: int = Field(..., description="Number of duplicate rows.")
    headers: list[str] = Field(..., description="Column headers for the returned rows.")


class CsvPreview(BaseModel):
    """Preview of a CSV file's contents."""

    headers: list[str] = Field(..., description="Column headers from the first row.")
    rows: list[list[Any]] = Field(..., description="Preview rows of data.")
    total_rows: int = Field(..., description="Total number of rows in the CSV file.")


class WorkbookMetadata(BaseModel):
    """Full metadata for an Excel workbook."""

    file_path: str = Field(..., description="Absolute path to the workbook file.")
    sheets: list[SheetInfo] = Field(..., description="Dimensional info for each sheet.")
    active_sheet: str = Field(..., description="Name of the currently active sheet.")
    named_ranges: list[dict[str, str]] = Field(..., description="Named ranges defined in the workbook.")


class DataProfile(BaseModel):
    """Statistical profile of a dataset."""

    columns: list[str] = Field(..., description="Column names in the dataset.")
    row_count: int = Field(..., description="Total number of data rows.")
    missing_values: dict[str, int] = Field(..., description="Count of missing values per column.")
    duplicates: int = Field(..., description="Number of fully duplicate rows.")
    summary_statistics: dict[str, dict[str, Any]] = Field(
        ..., description="Per-column summary statistics (count, mean, std, min, max, etc.)."
    )


class PivotResult(BaseModel):
    """Result of a pivot table operation."""

    data: list[dict[str, Any]] = Field(..., description="Pivoted data rows as dictionaries.")
    index_columns: list[str] = Field(..., description="Columns used as the pivot index.")
    value_columns: list[str] = Field(..., description="Columns aggregated in the pivot.")
    operation: str = Field(..., description="Aggregation function applied (sum, mean, count, etc.).")


class ValidationResult(BaseModel):
    """Result of a formula or expression validation."""

    valid: bool = Field(..., description="Whether the expression is valid.")
    tokens: list[dict[str, str]] = Field(..., description="Parsed tokens from the expression.")
    error: str | None = Field(None, description="Error message if validation failed, or None.")


class ChunkReadResult(BaseModel):
    """A paginated chunk of rows from a large dataset."""

    rows: list[dict[str, Any]] = Field(..., description="Rows in this chunk as column-keyed dicts.")
    chunk_start: int = Field(..., description="0-based starting row index of this chunk.")
    chunk_size: int = Field(..., description="Number of rows in this chunk.")
    has_more: bool = Field(..., description="Whether more rows remain after this chunk.")


# ── Input Models ─────────────────────────────────────────────────────────────


class FormatOptions(BaseModel):
    """Cell formatting options for styling spreadsheet cells."""

    bold: bool = Field(False, description="Apply bold font weight.")
    italic: bool = Field(False, description="Apply italic font style.")
    font_size: int | None = Field(None, description="Font size in points.")
    font_color: str | None = Field(None, description="Font color as a hex string, e.g. 'FF0000'.")
    bg_color: str | None = Field(None, description="Background fill color as a hex string.")
    number_format: str | None = Field(None, description="Excel number format code, e.g. '#,##0.00'.")
    horizontal_alignment: str | None = Field(None, description="Horizontal alignment: left, center, right, justify.")
    vertical_alignment: str | None = Field(None, description="Vertical alignment: top, center, bottom.")
    wrap_text: bool = Field(False, description="Enable text wrapping within the cell.")
    border_style: str | None = Field(None, description="Border style: thin, medium, thick, double, etc.")
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


class ChartConfig(BaseModel):
    """Configuration for inserting a chart into a worksheet."""

    chart_type: Literal["bar", "column", "line", "pie", "scatter", "area"] = Field(
        ..., description="Type of chart to create."
    )
    title: str = Field("", description="Chart title text.")
    x_axis_title: str = Field("", description="Label for the X axis.")
    y_axis_title: str = Field("", description="Label for the Y axis.")
    style: int = Field(10, description="Built-in Excel chart style number.")
    width: float = Field(15, description="Chart width in cm.")
    height: float = Field(10, description="Chart height in cm.")


class SortCriteria(BaseModel):
    """Sorting specification for a single column."""

    column: str = Field(..., description="Column letter or header name to sort by.")
    ascending: bool = Field(True, description="Sort in ascending order when True, descending when False.")
