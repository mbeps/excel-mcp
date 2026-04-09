from __future__ import annotations

from typing import Literal

from mcp.types import ToolAnnotations

import mcp_server.tools.cleaning as _cleaning
import mcp_server.tools.csv_ops as _csv_ops

__all__ = [
    "split_column",
    "data_cleaner",
    "parse_date_column",
    "csv_ops",
]


def split_column(
    file_path: str,
    sheet_name: str,
    column: str,
    delimiter: str = ",",
    new_columns: list[str] | None = None,
    drop_original: bool = True,
    output_file: str | None = None,
    header_row: int = 1,
) -> dict:
    """Split a text column into multiple columns by delimiter."""
    return _cleaning.split_column(
        file_path,
        sheet_name,
        column,
        delimiter,
        new_columns,
        drop_original,
        output_file,
        header_row,
    )


def data_cleaner(
    file_path: str,
    sheet_name: str = "Sheet1",
    operations: list[str] | None = None,
    columns: list[str] | None = None,
    preview: bool = False,
    output_file: str | None = None,
    header_row: int = 1,
    fill_missing_strategy: str = "value",
    fill_value: str | None = None,
) -> dict:
    """Batch data cleaning pipeline.

    Operations: trim_whitespace, remove_empty_rows, remove_empty_columns,
    normalize_text, fix_numbers, remove_duplicates, fill_missing.
    """
    return _cleaning.data_cleaner(
        file_path,
        sheet_name,
        operations,
        columns,
        preview,
        output_file,
        header_row,
        fill_missing_strategy,
        fill_value,
    )


def parse_date_column(
    file_path: str,
    sheet_name: str,
    column: str,
    output_column: str | None = None,
    output_format: str = "%Y-%m-%d",
    dayfirst: bool = False,
    header_row: int = 1,
) -> dict:
    """Parse mixed date formats in a column and normalise to a standard output format."""
    return _cleaning.parse_date_column(
        file_path,
        sheet_name,
        column,
        output_column,
        output_format,
        dayfirst,
        header_row,
    )


def csv_ops(
    action: Literal["preview", "to_xlsx", "to_csv"],
    file_path: str | None = None,
    csv_path: str | None = None,
    xlsx_path: str | None = None,
    sheet_name: str = "Sheet1",
    output_path: str | None = None,
    rows: int = 10,
    delimiter: str = ",",
    encoding: str = "utf-8",
) -> dict | str:
    """CSV operations: preview, convert CSV to XLSX, export XLSX to CSV.

    action="preview": Preview CSV. Requires: file_path. Optional: rows, delimiter, encoding.
    action="to_xlsx": Convert CSV to XLSX. Requires: csv_path, xlsx_path. Optional: sheet_name, delimiter, encoding.
    action="to_csv": Export sheet to CSV. Requires: file_path, sheet_name, output_path. Optional: delimiter, encoding.
    """
    if action == "preview":
        if file_path is None:
            raise ValueError("file_path is required for action='preview'.")
        return _csv_ops.read_csv_preview(file_path, rows, delimiter, encoding)
    if action == "to_xlsx":
        effective_csv = csv_path or file_path
        effective_xlsx = xlsx_path or output_path
        if not effective_csv:
            raise ValueError("csv_path (or file_path) is required for action='to_xlsx'.")
        if not effective_xlsx:
            raise ValueError("xlsx_path (or output_path) is required for action='to_xlsx'.")
        return _csv_ops.csv_to_xlsx(effective_csv, effective_xlsx, sheet_name, delimiter, encoding)
    if action == "to_csv":
        if file_path is None:
            raise ValueError("file_path is required for action='to_csv'.")
        if output_path is None:
            raise ValueError("output_path is required for action='to_csv'.")
        return _csv_ops.xlsx_to_csv(file_path, sheet_name, output_path, delimiter, encoding)
    raise ValueError(f"Unknown action: {action}")


def register(mcp) -> None:
    """Register cleaning/CSV tools on *mcp*."""
    mcp.tool()(split_column)
    mcp.tool()(data_cleaner)
    mcp.tool(annotations=ToolAnnotations(destructiveHint=True))(parse_date_column)
    mcp.tool()(csv_ops)
