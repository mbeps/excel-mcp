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
    """Build a pivot table from a source sheet and optionally write it to `output_sheet`/`output_file`.

    Args:
        file_path: Workbook path.
        sheet_name: Source data sheet.
        index_cols: List of column names to use as index (rows).
        value_cols: Columns to aggregate.
        aggfunc: Aggregation function or dict (e.g. "sum", "mean" or {col: "sum"}).
        output_sheet: Optional destination sheet for pivot output.
        output_file: Optional file to write the pivot output.
        column_field: Optional field used for pivot columns.
        date_freq: Optional date grouping alias (e.g. 'ME', 'YE', 'W').

    Returns:
        dict: Details about output including created sheet and saved pivot metadata.

    Notes:
        - Writes to workbook when `output_sheet`/`output_file` is provided. Stores pivot definitions in `_mcp_pivots` for refresh.
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
    """Refresh a previously created pivot table by re-running its stored definition.

    Args:
        file_path: Workbook path containing stored pivot definitions.
        output_sheet: Name of the pivot output sheet to refresh.
        source_file_path, source_sheet: Optional explicit sources to override stored sources.

    Returns:
        dict: Summary of refresh results.

    Notes:
        - Mutates the workbook by overwriting the pivot output area.
    """
    return _pivot_etl.refresh_pivot_table(file_path, output_sheet, source_file_path, source_sheet)


def unpivot_data(
    file_path: str,
    sheet_name: str,
    id_vars: list[str],
    value_vars: list[str],
    var_name: str = "Variable",
    value_name: str = "Value",
) -> dict:
    """Melt (unpivot) wide-form data to long-form using id_vars and value_vars.

    Args:
        file_path: Workbook path.
        sheet_name: Source sheet.
        id_vars: Columns to keep as identifiers.
        value_vars: Columns to melt into variable/value pairs.
        var_name: Name for the variable column.
        value_name: Name for the value column.

    Returns:
        dict: Result summary and destination range if written.
    """
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
    """Join two sheets within the workbook similar to SQL join semantics.

    Args:
        file_path: Path to workbook.
        sheet1, sheet2: Names of the two sheets to join.
        join_key: Column name(s) common to both sheets (shorthand for left_on/right_on).
        how: One of "left", "right", "inner", "outer".
        output_sheet: Optional sheet name to write merged results.
        left_on, right_on: Optional explicit join keys for differently named columns.

    Returns:
        dict: Key counts and output information.
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
) -> str | dict:
    """Add a computed column either via pandas-eval formula or as a cumsum/rolling operation.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet name.
        new_column_name: Column name to add.
        expression: Expression string for pandas.eval when column_type=='formula'.
        has_header: Whether the sheet has a header row.
        column_type: One of 'formula', 'cumsum', 'rolling'.
        source_col: Required for 'cumsum' and 'rolling'.
        window: Integer window for rolling operations.
        rolling_func: Aggregation for rolling (default 'mean').

    Returns:
        str: Message indicating success and destination column.

    Notes:
        - Accepts user-provided expressions — underlying code performs AST checks; docstring should link to safety doc.
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
    """Remove duplicate rows from a sheet, optionally using a subset of columns.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet name.
        columns: Optional list of columns to consider for duplicates.
        keep: Which duplicate to keep: 'first', 'last', or False (drop all duplicates).

    Returns:
        str: Summary message and number of rows removed.

    Notes:
        - Destructive: modifies the workbook unless an `output_file` variant is implemented upstream.
    """
    return _pivot_etl.deduplicate_data(file_path, sheet_name, columns, keep)


def register(mcp) -> None:
    """Register pivot/ETL tools on *mcp*."""
    mcp.tool()(create_pivot_table)
    mcp.tool()(refresh_pivot_table)
    mcp.tool()(unpivot_data)
    mcp.tool()(merge_datasets)
    mcp.tool()(add_computed_column)
    mcp.tool(annotations=ToolAnnotations(destructiveHint=True))(deduplicate_data)
