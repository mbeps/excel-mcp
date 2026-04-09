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
    """Split a single text column into multiple columns using a delimiter.

    Args:
        file_path: Path to workbook.
        sheet_name: Worksheet name.
        column: Column letter or header name to split.
        delimiter: Delimiter string (default ",").
        new_columns: Optional list of new column names.
        drop_original: If True, remove the original column after split.
        output_file: Optional path to write results instead of overwriting input.
        header_row: 1-based header row index.

    Returns:
        dict: Summary of created columns and row counts.

    Notes:
        - Destructive unless `output_file` is provided.
    """
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
    """Run a pipeline of cleaning operations (trim, dedupe, fill missing, normalize, etc.) on a sheet.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet to operate on.
        operations: Ordered list of operations to run (see tool docs for allowed names).
        columns: Optional subset of columns to target.
        preview: If True, return a preview without persisting changes.
        output_file: Optional path to write cleaned results.
        header_row: 1-based header index.
        fill_missing_strategy: Strategy name for filling missing values.
        fill_value: Literal value to use if strategy is "value".

    Returns:
        dict: Summary including rows modified and operations applied.

    Notes:
        - Document allowed `operations` strings in the route or underlying tool docs.
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
    """Parse varying date formats in a column and write normalized results to an output column.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet name.
        column: Column to parse.
        output_column: Optional target column; if omitted, overwrites `column`.
        output_format: Strftime format for normalized output.
        dayfirst: Whether to parse day-first dates.
        header_row: 1-based header index.

    Returns:
        dict: Summary including number of parsed rows and parse errors.

    Notes:
        - Destructive by default when `output_column` targets existing column.
    """
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
    """CSV helper operations: preview CSV, convert CSV→XLSX, and export XLSX→CSV.

    Args:
        action: "preview", "to_xlsx", or "to_csv".
            - "preview": requires `file_path` (CSV path); returns a small preview as dict.
            - "to_xlsx": requires `csv_path` (or `file_path`) and `xlsx_path` (or `output_path`).
            - "to_csv": requires `file_path` (XLSX) and `output_path` (CSV destination).
        file_path, csv_path, xlsx_path, output_path: Path parameters as described above.
        sheet_name: Sheet to export when converting XLSX→CSV.
        rows: Number of preview rows to return.
        delimiter, encoding: CSV parameters.

    Returns:
        dict or str: Preview dict for "preview" or destination path for conversions.

    Raises:
        ValueError: If required path parameters are missing for the selected `action`.

    Notes:
        - "to_xlsx" and "to_csv" perform file writes; document whether they overwrite existing files.
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
