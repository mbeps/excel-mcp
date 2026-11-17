"""CSV import/export and preview operations."""

from __future__ import annotations

from logging import Logger
from pathlib import Path

import pandas as pd

from mcp_server.models.common import CellScalar
from mcp_server.utils.excel_helpers import sanitize_excel_value, validate_file_path
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def _read_csv_with_fallback(path: Path, delimiter: str, encoding: str) -> pd.DataFrame:
    """Read a CSV file with *encoding*, falling back to latin-1 on UnicodeDecodeError."""
    try:
        return pd.read_csv(path, delimiter=delimiter, encoding=encoding)
    except UnicodeDecodeError:
        logger.warning("Decode with encoding='%s' failed for %s, falling back to latin-1", encoding, path)
        return pd.read_csv(path, delimiter=delimiter, encoding="latin-1")


def read_csv_preview(
    file_path: str, rows: int = 10, delimiter: str = ",", encoding: str = "utf-8"
) -> dict[str, list[list[CellScalar]] | int | list[str]]:
    """Preview the first N rows of a CSV file."""
    path = validate_file_path(file_path)
    df = _read_csv_with_fallback(path, delimiter, encoding)
    preview = df.head(rows)
    return {
        "headers": list(df.columns),
        "rows": [[sanitize_excel_value(val) for val in row] for row in preview.values.tolist()],
        "total_rows": len(df),
    }


def csv_to_xlsx(
    csv_path: str, xlsx_path: str, sheet_name: str = "Sheet1", delimiter: str = ",", encoding: str = "utf-8"
) -> str:
    """Convert a CSV file to .xlsx format."""
    src = validate_file_path(csv_path)
    validate_file_path(xlsx_path, must_exist=False)
    df = _read_csv_with_fallback(src, delimiter, encoding)
    df.to_excel(xlsx_path, sheet_name=sheet_name, index=False, engine="openpyxl")
    logger.info("Converted %s -> %s", csv_path, xlsx_path)
    return f"Converted '{csv_path}' to '{xlsx_path}' (sheet '{sheet_name}', {len(df)} rows)."


def xlsx_to_csv(
    file_path: str, sheet_name: str, output_path: str, delimiter: str = ",", encoding: str = "utf-8"
) -> str:
    """Export a worksheet to CSV."""
    src = validate_file_path(file_path)
    validate_file_path(output_path, must_exist=False)
    df = pd.read_excel(src, sheet_name=sheet_name, engine="openpyxl")
    df.to_csv(output_path, index=False, sep=delimiter, encoding=encoding)
    logger.info("Exported %s!%s -> %s", file_path, sheet_name, output_path)
    return f"Exported '{sheet_name}' from '{file_path}' to '{output_path}' ({len(df)} rows)."
