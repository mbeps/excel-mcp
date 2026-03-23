"""CSV import/export and preview operations."""

from __future__ import annotations

import csv
from logging import Logger

import pandas as pd

from mcp_server.utils.excel_helpers import validate_file_path
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def read_csv_preview(file_path: str, rows: int = 10, delimiter: str = ",", encoding: str = "utf-8") -> dict:
    """Preview the first N rows of a CSV file."""
    path = validate_file_path(file_path)
    try:
        df = pd.read_csv(path, delimiter=delimiter, encoding=encoding)
    except UnicodeDecodeError:
        logger.warning("UTF-8 decode failed for %s, falling back to latin-1", file_path)
        df = pd.read_csv(path, delimiter=delimiter, encoding="latin-1")
    preview = df.head(rows)
    return {
        "headers": list(df.columns),
        "rows": preview.values.tolist(),
        "total_rows": len(df),
    }


def detect_csv_dialect(file_path: str, sample_bytes: int = 4096) -> dict:
    """Auto-detect CSV delimiter, quote character, and encoding."""
    path = validate_file_path(file_path)
    raw = path.read_bytes()[:sample_bytes]

    # Detect encoding
    if raw.startswith(b"\xef\xbb\xbf"):
        encoding = "utf-8-sig"
    else:
        try:
            raw.decode("utf-8")
            encoding = "utf-8"
        except UnicodeDecodeError:
            encoding = "latin-1"

    sample = raw.decode(encoding, errors="replace")

    try:
        sniffer = csv.Sniffer()
        dialect = sniffer.sniff(sample)
        delimiter = dialect.delimiter
        quotechar = dialect.quotechar or '"'
        has_header = sniffer.has_header(sample)
    except csv.Error:
        delimiter = ","
        quotechar = '"'
        has_header = True

    return {
        "delimiter": delimiter,
        "quotechar": quotechar,
        "encoding": encoding,
        "has_header": has_header,
    }


def validate_csv(
    file_path: str,
    expected_columns: list[str] | None = None,
    delimiter: str = ",",
    encoding: str = "utf-8",
) -> dict:
    """Validate a CSV file's structure and optionally check expected columns."""
    path = validate_file_path(file_path)
    try:
        df = pd.read_csv(path, delimiter=delimiter, encoding=encoding)
    except UnicodeDecodeError:
        logger.warning("UTF-8 decode failed for %s, falling back to latin-1", file_path)
        df = pd.read_csv(path, delimiter=delimiter, encoding="latin-1")

    actual_columns = list(df.columns)
    missing_columns: list[str] = []
    extra_columns: list[str] = []

    if expected_columns is not None:
        missing_columns = [c for c in expected_columns if c not in actual_columns]
        extra_columns = [c for c in actual_columns if c not in expected_columns]

    empty_rows = int(df.isna().all(axis=1).sum())
    empty_columns = [c for c in actual_columns if df[c].isna().all()]

    valid = len(missing_columns) == 0

    return {
        "valid": valid,
        "column_count": len(actual_columns),
        "row_count": len(df),
        "columns": actual_columns,
        "missing_columns": missing_columns,
        "extra_columns": extra_columns,
        "empty_rows": empty_rows,
        "empty_columns": empty_columns,
    }


def csv_to_xlsx(
    csv_path: str, xlsx_path: str, sheet_name: str = "Sheet1", delimiter: str = ",", encoding: str = "utf-8"
) -> str:
    """Convert a CSV file to .xlsx format."""
    src = validate_file_path(csv_path)
    validate_file_path(xlsx_path, must_exist=False)
    df = pd.read_csv(src, delimiter=delimiter, encoding=encoding)
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
