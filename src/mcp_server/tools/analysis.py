"""Data analysis operations: filter, sort, statistics, aggregation, and search/replace."""

from __future__ import annotations

from logging import Logger

import pandas as pd

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
    validate_file_path,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

FILTER_OPERATORS = {"==", "!=", ">", "<", ">=", "<=", "contains", "startswith", "endswith"}
AGGREGATE_OPERATIONS = {"sum", "mean", "count", "min", "max", "median", "std"}


def _read_sheet_df(file_path: str, sheet_name: str, has_header: bool = True) -> pd.DataFrame:
    path = validate_file_path(file_path)
    header = 0 if has_header else None
    return pd.read_excel(path, sheet_name=sheet_name, header=header, engine="openpyxl")


def filter_data(
    file_path: str,
    sheet_name: str,
    column: str,
    operator: str,
    value: str | int | float,
    has_header: bool = True,
) -> dict:
    """Filter rows by a column condition."""
    if operator not in FILTER_OPERATORS:
        raise ValueError(f"Unsupported operator '{operator}'. Allowed: {FILTER_OPERATORS}")

    df = _read_sheet_df(file_path, sheet_name, has_header)
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found. Available: {list(df.columns)}")

    col = df[column]
    if operator == "==":
        mask = col == value
    elif operator == "!=":
        mask = col != value
    elif operator == ">":
        mask = col > value
    elif operator == "<":
        mask = col < value
    elif operator == ">=":
        mask = col >= value
    elif operator == "<=":
        mask = col <= value
    elif operator == "contains":
        mask = col.astype(str).str.contains(str(value), case=False, na=False)
    elif operator == "startswith":
        mask = col.astype(str).str.startswith(str(value), na=False)
    else:  # endswith
        mask = col.astype(str).str.endswith(str(value), na=False)

    matched = df[mask]
    return {
        "matched_rows": matched.values.tolist(),
        "total_rows": len(df),
        "matched_count": len(matched),
        "headers": list(df.columns),
    }


def sort_data(
    file_path: str,
    sheet_name: str,
    sort_by: list[dict] | None = None,
    column: str | None = None,
    ascending: bool = True,
    has_header: bool = True,
) -> str:
    """Sort sheet data by one or more columns and write back."""
    df = _read_sheet_df(file_path, sheet_name, has_header)

    if sort_by:
        cols = [s["column"] for s in sort_by]
        asc = [s.get("ascending", True) for s in sort_by]
        for c in cols:
            if c not in df.columns:
                raise ValueError(f"Column '{c}' not found. Available: {list(df.columns)}")
        df = df.sort_values(by=cols, ascending=asc)
    elif column:
        if column not in df.columns:
            raise ValueError(f"Column '{column}' not found. Available: {list(df.columns)}")
        df = df.sort_values(by=column, ascending=ascending)
    else:
        raise ValueError("Provide either 'sort_by' or 'column' to sort.")

    wb = load_workbook_safe(file_path)
    ws = get_sheet(wb, sheet_name)

    # Clear existing data rows (keep header if present)
    start_row = 2 if has_header else 1
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.value = None

    for r_idx, row_data in enumerate(df.values.tolist(), start=start_row):
        for c_idx, val in enumerate(row_data, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val)

    save_workbook_safe(wb, file_path)
    logger.info("Sorted sheet '%s' in %s", sheet_name, file_path)
    return f"Sorted {len(df)} rows in '{sheet_name}'."


def column_statistics(file_path: str, sheet_name: str, column: str, has_header: bool = True) -> dict:
    """Compute descriptive statistics for a numeric column."""
    df = _read_sheet_df(file_path, sheet_name, has_header)
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found. Available: {list(df.columns)}")

    col = df[column]
    if not pd.api.types.is_numeric_dtype(col):
        return {
            "column": column,
            "message": f"Column '{column}' is not numeric (dtype: {col.dtype}). Cannot compute statistics.",
        }

    desc = col.describe()
    return {
        "column": column,
        "count": int(desc["count"]),
        "mean": float(desc["mean"]),
        "median": float(col.median()),
        "min_val": float(desc["min"]),
        "max_val": float(desc["max"]),
        "std": float(desc["std"]),
        "sum_val": float(col.sum()),
    }


def aggregate_data(
    file_path: str,
    sheet_name: str,
    group_by: str,
    value_column: str,
    operation: str,
    has_header: bool = True,
) -> dict:
    """Group by a column and apply an aggregation operation."""
    if operation not in AGGREGATE_OPERATIONS:
        raise ValueError(f"Unsupported operation '{operation}'. Allowed: {AGGREGATE_OPERATIONS}")

    df = _read_sheet_df(file_path, sheet_name, has_header)
    for col_name in (group_by, value_column):
        if col_name not in df.columns:
            raise ValueError(f"Column '{col_name}' not found. Available: {list(df.columns)}")

    grouped = df.groupby(group_by)[value_column].agg(operation).reset_index()
    grouped.columns = [group_by, f"{value_column}_{operation}"]

    return {
        "groups": grouped.to_dict(orient="records"),
        "group_by": group_by,
        "operation": operation,
    }


def find_duplicates(file_path: str, sheet_name: str, columns: list[str], has_header: bool = True) -> dict:
    """Find duplicate rows based on specified columns."""
    df = _read_sheet_df(file_path, sheet_name, has_header)
    for c in columns:
        if c not in df.columns:
            raise ValueError(f"Column '{c}' not found. Available: {list(df.columns)}")

    mask = df.duplicated(subset=columns, keep=False)
    dupes = df[mask]
    return {
        "duplicates": dupes.values.tolist(),
        "count": len(dupes),
        "headers": list(df.columns),
    }


def profile_data(file_path: str, sheet_name: str) -> dict:
    """Comprehensive data profiling: types, missing values, duplicates, summary stats."""
    df = _read_sheet_df(file_path, sheet_name, has_header=True)

    missing = {col: int(df[col].isna().sum()) for col in df.columns}

    numeric_cols = df.select_dtypes(include="number")
    summary = {}
    if not numeric_cols.empty:
        desc = numeric_cols.describe()
        summary = {col: desc[col].to_dict() for col in desc.columns}

    return {
        "columns": [{"name": col, "dtype": str(df[col].dtype)} for col in df.columns],
        "row_count": len(df),
        "missing_values": missing,
        "duplicates": int(df.duplicated().sum()),
        "summary_statistics": summary,
    }


def search_replace(
    file_path: str,
    sheet_name: str,
    search_value: str,
    replace_value: str,
    cell_range: str | None = None,
) -> str:
    """Find and replace values in a sheet or specific range."""
    wb = load_workbook_safe(file_path)
    ws = get_sheet(wb, sheet_name)

    count = 0
    if cell_range:
        from openpyxl.utils import range_boundaries

        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        cells = ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col)
    else:
        cells = ws.iter_rows(
            min_row=ws.min_row or 1,
            max_row=ws.max_row or 1,
            min_col=ws.min_column or 1,
            max_col=ws.max_column or 1,
        )

    for row in cells:
        for cell in row:
            if cell.value is not None and str(cell.value) == search_value:
                cell.value = replace_value
                count += 1

    save_workbook_safe(wb, file_path)
    logger.info("Replaced %d occurrences of '%s' in %s!%s", count, search_value, sheet_name, file_path)
    return f"Replaced {count} occurrence(s) of '{search_value}' with '{replace_value}' in '{sheet_name}'."


def calculate_correlation(
    file_path: str,
    sheet_name: str,
    columns: list[str],
    has_header: bool = True,
) -> dict:
    """Calculate correlation matrix between specified numeric columns."""
    df = _read_sheet_df(file_path, sheet_name, has_header)
    for c in columns:
        if c not in df.columns:
            raise ValueError(f"Column '{c}' not found. Available: {list(df.columns)}")

    corr = df[columns].corr()
    return {
        "columns": columns,
        "matrix": corr.values.tolist(),
    }


def rank_data(
    file_path: str,
    sheet_name: str,
    column: str,
    method: str = "dense",
    ascending: bool = True,
    has_header: bool = True,
) -> dict:
    """Rank values in a column."""
    df = _read_sheet_df(file_path, sheet_name, has_header)
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found. Available: {list(df.columns)}")

    ranks = df[column].rank(method=method, ascending=ascending)
    rankings = [{"value": v, "rank": r} for v, r in zip(df[column].tolist(), ranks.tolist())]
    return {
        "column": column,
        "method": method,
        "rankings": rankings,
    }


def calculate_percentiles(
    file_path: str,
    sheet_name: str,
    column: str,
    percentiles: list[float] | None = None,
    has_header: bool = True,
) -> dict:
    """Calculate percentile values for a column."""
    df = _read_sheet_df(file_path, sheet_name, has_header)
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found. Available: {list(df.columns)}")

    if percentiles is None:
        percentiles = [0.1, 0.25, 0.5, 0.75, 0.9]

    result = df[column].quantile(percentiles)
    return {
        "column": column,
        "percentile_values": {str(p): float(v) for p, v in result.items()},
    }


def sample_data(
    file_path: str,
    sheet_name: str,
    n: int = 10,
    fraction: float | None = None,
    random_state: int | None = None,
    has_header: bool = True,
) -> dict:
    """Random sample of rows from a sheet."""
    df = _read_sheet_df(file_path, sheet_name, has_header)

    if fraction is not None:
        sampled = df.sample(frac=fraction, random_state=random_state)
    else:
        sampled = df.sample(n=min(n, len(df)), random_state=random_state)

    return {
        "rows": sampled.values.tolist(),
        "sample_size": len(sampled),
        "total_rows": len(df),
        "headers": list(df.columns),
    }


def create_histogram(
    file_path: str,
    sheet_name: str,
    column: str,
    bins: int = 10,
    has_header: bool = True,
) -> dict:
    """Create histogram bins for a numeric column."""
    df = _read_sheet_df(file_path, sheet_name, has_header)
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found. Available: {list(df.columns)}")

    cut = pd.cut(df[column], bins=bins)
    counts = cut.value_counts(sort=False)
    bin_edges = [float(interval.left) for interval in counts.index] + [float(counts.index[-1].right)]
    frequencies = counts.tolist()

    return {
        "column": column,
        "bin_edges": bin_edges,
        "frequencies": frequencies,
    }


def transpose_data(
    file_path: str,
    sheet_name: str,
    output_sheet: str | None = None,
    has_header: bool = True,
) -> str:
    """Transpose data (rows <-> columns) and write back."""
    df = _read_sheet_df(file_path, sheet_name, has_header)
    transposed = df.T.reset_index()

    target_sheet = output_sheet or sheet_name
    wb = load_workbook_safe(file_path)

    if target_sheet in wb.sheetnames:
        ws = wb[target_sheet]
        for row in ws.iter_rows():
            for cell in row:
                cell.value = None
    else:
        ws = wb.create_sheet(target_sheet)

    for c_idx, col_val in enumerate(transposed.columns, start=1):
        ws.cell(row=1, column=c_idx, value=col_val)

    for r_idx, row_data in enumerate(transposed.values.tolist(), start=2):
        for c_idx, val in enumerate(row_data, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val)

    save_workbook_safe(wb, file_path)
    logger.info("Transposed data from '%s' to '%s' in %s", sheet_name, target_sheet, file_path)
    return f"Transposed data from '{sheet_name}' to '{target_sheet}' ({len(df)} rows -> {len(df.columns)} rows)."


def extract_unique_values(
    file_path: str,
    sheet_name: str,
    column: str,
    has_header: bool = True,
) -> dict:
    """Extract unique values from a column."""
    df = _read_sheet_df(file_path, sheet_name, has_header)
    if column not in df.columns:
        raise ValueError(f"Column '{column}' not found. Available: {list(df.columns)}")

    unique = df[column].dropna().unique().tolist()
    return {
        "column": column,
        "unique_values": unique,
        "count": len(unique),
    }
