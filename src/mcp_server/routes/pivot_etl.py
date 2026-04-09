from __future__ import annotations

from mcp.types import ToolAnnotations

import mcp_server.tools.pivot_etl as _pivot_etl

__all__ = [
    "create_pivot_table",
    "refresh_pivot_table",
    "unpivot_data",
    "merge_datasets",
    "add_computed_column",
    "deduplicate_data",
]


def create_pivot_table(
    file_path: str,
    sheet_name: str,
    index_cols: list[str],
    value_cols: list[str],
    aggfunc: str | dict = "sum",
    output_sheet: str | None = None,
    output_file: str | None = None,
    column_field: str | None = None,
    date_freq: str | None = None,
) -> dict:
    """Create a pivot table and optionally write results to a sheet or file.

    date_freq: if set, groups datetime index columns by this period before pivoting.
      Common values: 'ME' (month-end), 'QE' (quarter-end), 'YE' (year-end), 'W' (weekly).
      Any valid pandas DateOffset alias is accepted.
    """
    return _pivot_etl.create_pivot_table(
        file_path,
        sheet_name,
        index_cols,
        value_cols,
        aggfunc,
        output_sheet,
        output_file,
        False,
        column_field,
        date_freq,
    )


def refresh_pivot_table(
    file_path: str,
    output_sheet: str,
    source_file_path: str | None = None,
    source_sheet: str | None = None,
) -> dict:
    """Refresh a previously-created pivot table by re-running its stored definition."""
    return _pivot_etl.refresh_pivot_table(file_path, output_sheet, source_file_path, source_sheet)


def unpivot_data(
    file_path: str,
    sheet_name: str,
    id_vars: list[str],
    value_vars: list[str],
    var_name: str = "Variable",
    value_name: str = "Value",
) -> dict:
    """Unpivot (melt) data from wide to long format."""
    return _pivot_etl.unpivot_data(file_path, sheet_name, id_vars, value_vars, var_name, value_name)


def merge_datasets(
    file_path: str,
    sheet1: str,
    sheet2: str,
    join_key: str | list[str] | None = None,
    how: str = "left",
    output_sheet: str | None = None,
    left_on: str | list[str] | None = None,
    right_on: str | list[str] | None = None,
) -> dict:
    """Merge two sheets like a SQL join (left, right, inner, outer).

    Use ``join_key`` when both sheets share the same column name(s).
    Use ``left_on`` / ``right_on`` to join on differently-named columns.
    """
    return _pivot_etl.merge_datasets(
        file_path,
        sheet1,
        sheet2,
        join_key,
        how,
        output_sheet,
        left_on=left_on,
        right_on=right_on,
    )


def add_computed_column(
    file_path: str,
    sheet_name: str,
    new_column_name: str,
    expression: str,
    has_header: bool = True,
    column_type: str = "formula",
    source_col: str | None = None,
    window: int | None = None,
    rolling_func: str = "mean",
) -> str:
    """Add a computed column using a pandas-eval expression (e.g. 'Revenue - Cost').

    column_type='formula' (default): evaluate expression via pandas eval.
    column_type='cumsum': compute a running total of source_col. Requires: source_col.
    column_type='rolling': compute a rolling window aggregation of source_col.
      Requires: source_col, window (int). Optional: rolling_func ('mean' or 'sum', default 'mean').
    """
    return _pivot_etl.add_computed_column(
        file_path,
        sheet_name,
        new_column_name,
        expression,
        has_header,
        column_type,
        source_col,
        window,
        rolling_func,
    )


def deduplicate_data(
    file_path: str,
    sheet_name: str,
    columns: list[str] | None = None,
    keep: str = "first",
) -> str:
    """Remove duplicate rows from a sheet. keep: 'first', 'last', or False."""
    return _pivot_etl.deduplicate_data(file_path, sheet_name, columns, keep)


def register(mcp) -> None:
    """Register pivot/ETL tools on *mcp*."""
    mcp.tool()(create_pivot_table)
    mcp.tool()(refresh_pivot_table)
    mcp.tool()(unpivot_data)
    mcp.tool()(merge_datasets)
    mcp.tool()(add_computed_column)
    mcp.tool(annotations=ToolAnnotations(destructiveHint=True))(deduplicate_data)
