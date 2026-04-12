from __future__ import annotations

from typing import Any, Literal

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
    """Perform cross-workbook operations: aggregate, filter, validate schema consistency, compare two workbooks.

    Args:
        action: "aggregate", "filter", "validate", or "compare".
            - "aggregate": requires `file_paths` and `column`; returns aggregated metric.
            - "filter": requires `file_paths`, `column`, `operator`, `value`; returns filtered rows or writes to `output_file`.
            - "validate": requires `file_paths` and `key_column`; checks schema/consistency across files.
            - "compare": requires `file_a` and `file_b`; returns diff summary and optionally writes a report.
        file_paths, file_a, file_b: File list / pair for relevant actions.
        column, operation, operator, value: Parameters for aggregation/filtering.
        output_file: Optional path to write the result.

    Returns:
        dict: Operation-specific result (e.g. aggregation numbers, diffs, validation errors).

    Notes:
        - When `output_file` is provided operations may write new files — document overwrite policy.
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


def register(mcp: Any) -> None:
    """Register tools on *mcp*."""
    mcp.tool()(multi_file)
