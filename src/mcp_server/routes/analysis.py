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
    """Sort worksheet rows by one or more columns and write back the result.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet to sort.
        sort_by: Optional explicit sort descriptor list (e.g. [{"column": "A", "ascending": True}]).
        column: Convenience single-column sort (deprecated in favour of `sort_by`).
        ascending: Boolean default sort order when `column` is used.
        has_header: Whether the sheet has a header row.

    Returns:
        str: Result message.

    Notes:
        - Destructive: overwrites sheet rows.
    """
    return _analysis.sort_data(file_path, sheet_name, sort_by, column, ascending, has_header)


def column_statistics(file_path: str, sheet_name: str, column: str, has_header: bool = True) -> ColumnStats:
    """Compute descriptive statistics for a numeric column (mean, median, std, min, max, sum).

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet name.
        column: Column name or letter to analyse.
        has_header: Whether the sheet has a header row.

    Returns:
        ColumnStats: Pydantic model with statistical measures.

    Notes:
        - Read-only.
    """
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
    """Group rows by column(s) and aggregate values using the specified operation.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet name.
        group_by: Column or list of columns to group by.
        value_column: Column to aggregate.
        operation: Aggregation operation (e.g. 'sum', 'mean', 'count').
        has_header: Whether the sheet has a header row.
        aggfunc: Optional pandas-style aggfunc or mapping.

    Returns:
        dict: Aggregated results (may be written to sheet if underlying tool provides an option).
    """
    return _analysis.aggregate_data(file_path, sheet_name, group_by, value_column, operation, has_header, aggfunc)


def find_duplicates(file_path: str, sheet_name: str, columns: list[str], has_header: bool = True) -> dict:
    """Identify duplicate rows based on a list of columns.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet name.
        columns: Columns used to determine duplicates.
        has_header: Whether the sheet has a header row.

    Returns:
        dict: Duplicate groups and row indices.
    """
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
    """Perform cross-file lookup akin to VLOOKUP with optional fuzzy matching.

    Args:
        lookup_file: Workbook containing keys to look up.
        data_file: Workbook containing reference data.
        lookup_column: Column in `lookup_file` to match.
        data_key_column: Column in `data_file` to join on.
        data_return_columns: Columns from `data_file` to return.
        lookup_sheet, data_sheet: Sheet names.
        fuzzy: If True, perform fuzzy matching.
        fuzzy_threshold: Threshold for fuzzy confidence.
        output_file: Optional path to write augmented lookup results.
        header_row: 1-based header row index.

    Returns:
        dict: Mapping rows to matched results and match scores.

    Notes:
        - Read-only on inputs unless `output_file` is provided.
    """
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
    """Filter rows using multiple conditions combined with AND/OR logic.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet name.
        conditions: List of condition dicts (each: {column, operator, value}).
        logic: "AND" or "OR" to combine conditions.
        output_sheet: Optional sheet to write filtered output.
        header_row: 1-based header row index.

    Returns:
        dict: Filtered rows or summary.
    """
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

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet name.
        group_col: Column used to group rows.
        value_col: Column to subtotal.
        subtotal_func: Excel subtotal function code (9=SUM by default).
        include_grand_total: Whether to append a grand total row.

    Returns:
        dict: Summary including ranges where subtotals were inserted.

    Notes:
        - Destructive: modifies the sheet structure and inserts new rows.
    """
    return _analysis.insert_subtotals(file_path, sheet_name, group_col, value_col, subtotal_func, include_grand_total)


def profile_data(
    file_path: str,
    sheet: str | None = None,
    data_range: str | None = None,
) -> dict:
    """Produce a data profile for a sheet or range listing types, null counts, unique counts and samples.

    Args:
        file_path: Workbook path.
        sheet: Optional sheet name.
        data_range: Optional range to restrict profiling.

    Returns:
        dict: Per-column profile metadata.

    Notes:
        - Read-only.
    """
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
    """Return frequency counts for a column as counts or normalized proportions.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet name.
        column: Column name to analyse.
        normalize: If True return proportions instead of raw counts.
        top_n: If provided, return only the top N values.
        dropna: Exclude nulls when True.
        has_header: Whether the sheet has a header row.

    Returns:
        dict: {"column", "total_rows", "normalize", "counts": [{"value", "count"}, ...]}.
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
