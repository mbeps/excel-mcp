from __future__ import annotations

from typing import Literal

import mcp_server.tools.multi_file as _multi_file

__all__ = [
    "multi_file",
]


def multi_file(
    action: Literal["aggregate", "filter", "validate", "compare"],
    file_paths: list[str] | None = None,
    file_a: str | None = None,
    file_b: str | None = None,
    column: str | None = None,
    operation: str = "sum",
    operator: str | None = None,
    value: str | float | None = None,
    key_column: str | None = None,
    check_columns: list[str] | None = None,
    sheet_name: str = "Sheet1",
    header_row: int = 1,
    output_file: str | None = None,
    compare_values: bool = True,
    compare_formulas: bool = False,
    sheet_name_b: str | None = None,
) -> dict:
    """Cross-file operations.

    action="aggregate": Aggregate a column across files. Requires: file_paths, column. Optional: operation.
    action="filter": Filter rows across files. Requires: file_paths, column, operator, value.
    action="validate": Cross-file consistency check. Requires: file_paths, key_column.
    action="compare": Compare two workbooks. Requires: file_a, file_b. Optional: sheet_name, sheet_name_b.
    """
    if action == "aggregate":
        if file_paths is None:
            raise ValueError("file_paths is required for action='aggregate'.")
        if column is None:
            raise ValueError("column is required for action='aggregate'.")
        return _multi_file.bulk_aggregate_multi_files(
            file_paths,
            column,
            operation,
            sheet_name,
            header_row,
            output_file,
        )
    if action == "filter":
        if file_paths is None:
            raise ValueError("file_paths is required for action='filter'.")
        if column is None:
            raise ValueError("column is required for action='filter'.")
        if operator is None:
            raise ValueError("operator is required for action='filter'.")
        if value is None:
            raise ValueError("value is required for action='filter'.")
        return _multi_file.bulk_filter_multi_files(
            file_paths,
            column,
            operator,
            value,
            sheet_name,
            header_row,
            output_file,
        )
    if action == "validate":
        if file_paths is None:
            raise ValueError("file_paths is required for action='validate'.")
        if key_column is None:
            raise ValueError("key_column is required for action='validate'.")
        return _multi_file.validate_data_consistency(file_paths, key_column, check_columns, sheet_name, header_row)
    if action == "compare":
        if not file_a:
            raise ValueError("file_a is required for action='compare'.")
        if not file_b:
            raise ValueError("file_b is required for action='compare'.")
        return _multi_file.compare_workbooks(file_a, file_b, sheet_name, output_file, sheet_name_b)
    raise ValueError(f"Unknown action: {action}")


def register(mcp) -> None:
    """Register tools on *mcp*."""
    mcp.tool()(multi_file)
