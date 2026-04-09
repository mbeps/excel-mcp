from __future__ import annotations

from mcp.types import ToolAnnotations

import mcp_server.tools.analysis as _analysis
from mcp_server.models.analysis import ColumnStats

__all__ = [
    "sort_data",
    "column_statistics",
    "aggregate_data",
    "find_duplicates",
    "vlookup_helper",
    "filter_data_advanced",
    "insert_subtotals",
    "profile_data",
    "value_counts",
]


def sort_data(
    file_path: str,
    sheet_name: str,
    sort_by: list[dict] | None = None,
    column: str | None = None,
    ascending: bool = True,
    has_header: bool = True,
) -> str:
    """Sort sheet data by one or more columns and write back."""
    return _analysis.sort_data(file_path, sheet_name, sort_by, column, ascending, has_header)


def column_statistics(file_path: str, sheet_name: str, column: str, has_header: bool = True) -> ColumnStats:
    """Compute descriptive statistics (mean, median, std, min, max, sum) for a numeric column."""
    return _analysis.column_statistics(file_path, sheet_name, column, has_header)


def aggregate_data(
    file_path: str,
    sheet_name: str,
    group_by: str | list[str],
    value_column: str,
    operation: str = "sum",
    has_header: bool = True,
    aggfunc: str | dict | None = None,
) -> dict:
    """Group by one or more columns and aggregate (sum, mean, count, min, max, median, std)."""
    return _analysis.aggregate_data(file_path, sheet_name, group_by, value_column, operation, has_header, aggfunc)


def find_duplicates(file_path: str, sheet_name: str, columns: list[str], has_header: bool = True) -> dict:
    """Find duplicate rows based on specified columns."""
    return _analysis.find_duplicates(file_path, sheet_name, columns, has_header)


def vlookup_helper(
    lookup_file: str,
    data_file: str,
    lookup_column: str,
    data_key_column: str,
    data_return_columns: list[str],
    lookup_sheet: str = "Sheet1",
    data_sheet: str = "Sheet1",
    fuzzy: bool = False,
    fuzzy_threshold: float = 0.8,
    output_file: str | None = None,
    header_row: int = 1,
) -> dict:
    """Cross-file VLOOKUP with optional fuzzy string matching."""
    return _analysis.vlookup_helper(
        lookup_file,
        data_file,
        lookup_column,
        data_key_column,
        data_return_columns,
        lookup_sheet,
        data_sheet,
        fuzzy,
        fuzzy_threshold,
        output_file,
        header_row,
    )


def filter_data_advanced(
    file_path: str,
    sheet_name: str,
    conditions: list[dict],
    logic: str = "AND",
    output_sheet: str | None = None,
    header_row: int = 1,
) -> dict:
    """Multi-condition AND/OR filtering. Each condition: {column, operator, value}."""
    return _analysis.filter_data_advanced(file_path, sheet_name, conditions, logic, output_sheet, header_row)


def insert_subtotals(
    file_path: str,
    sheet_name: str,
    group_col: str,
    value_col: str,
    subtotal_func: int = 9,
    include_grand_total: bool = True,
) -> dict:
    """Insert SUBTOTAL formula rows after each group in a sorted sheet.

    subtotal_func: 1=AVERAGE, 2=COUNT, 3=COUNTA, 4=MAX, 5=MIN, 9=SUM.
    """
    return _analysis.insert_subtotals(file_path, sheet_name, group_col, value_col, subtotal_func, include_grand_total)


def profile_data(
    file_path: str,
    sheet: str | None = None,
    data_range: str | None = None,
) -> dict:
    """Profile data in a worksheet, returning column statistics, types, null counts, and sample values."""
    return _analysis.profile_data(file_path, sheet, data_range)


def value_counts(
    file_path: str,
    sheet_name: str,
    column: str,
    normalize: bool = False,
    top_n: int | None = None,
    dropna: bool = True,
    has_header: bool = True,
) -> dict:
    """Return a full value-frequency table for a column.

    column: header name of the column to count.
    normalize: if True, return proportions (0-1) instead of raw counts.
    top_n: if given, return only the top N most-frequent values.
    dropna: if True (default), exclude null values from counts.
    Returns: {"column", "total_rows", "normalize", "counts": [{"value", "count"}, ...]}.
    """
    return _analysis.value_counts(file_path, sheet_name, column, normalize, top_n, dropna, has_header)


def register(mcp) -> None:
    """Register analysis tools on *mcp*."""
    mcp.tool()(sort_data)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(column_statistics)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(aggregate_data)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(find_duplicates)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(vlookup_helper)
    mcp.tool()(filter_data_advanced)
    mcp.tool()(insert_subtotals)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(profile_data)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(value_counts)
