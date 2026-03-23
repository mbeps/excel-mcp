"""CSV import/export and preview operations."""

from __future__ import annotations

from logging import Logger

import pandas as pd

from mcp_server.utils.excel_helpers import validate_file_path
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def read_csv_preview(file_path: str, rows: int = 10, delimiter: str = ",") -> dict:
    """Preview the first N rows of a CSV file."""
    path = validate_file_path(file_path)
    df = pd.read_csv(path, delimiter=delimiter)
    preview = df.head(rows)
    return {
        "headers": list(df.columns),
        "rows": preview.values.tolist(),
        "total_rows": len(df),
    }


def csv_to_xlsx(csv_path: str, xlsx_path: str, sheet_name: str = "Sheet1", delimiter: str = ",") -> str:
    """Convert a CSV file to .xlsx format."""
    src = validate_file_path(csv_path)
    validate_file_path(xlsx_path, must_exist=False)
    df = pd.read_csv(src, delimiter=delimiter)
    df.to_excel(xlsx_path, sheet_name=sheet_name, index=False, engine="openpyxl")
    logger.info("Converted %s -> %s", csv_path, xlsx_path)
    return f"Converted '{csv_path}' to '{xlsx_path}' (sheet '{sheet_name}', {len(df)} rows)."


def xlsx_to_csv(file_path: str, sheet_name: str, output_path: str, delimiter: str = ",") -> str:
    """Export a worksheet to CSV."""
    src = validate_file_path(file_path)
    validate_file_path(output_path, must_exist=False)
    df = pd.read_excel(src, sheet_name=sheet_name, engine="openpyxl")
    df.to_csv(output_path, index=False, sep=delimiter)
    logger.info("Exported %s!%s -> %s", file_path, sheet_name, output_path)
    return f"Exported '{sheet_name}' from '{file_path}' to '{output_path}' ({len(df)} rows)."
