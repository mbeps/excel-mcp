"""Data analysis operations: filter, sort, statistics, aggregation, and search/replace."""

from __future__ import annotations

import difflib
from logging import Logger

import openpyxl
import pandas as pd
from scipy import stats

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
    from mcp_server.utils.excel_helpers import read_sheet_df

    return read_sheet_df(file_path, sheet_name, header_row=1 if has_header else 0)


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
    result = {
        "column": column,
        "count": int(desc["count"]),
        "mean": float(desc["mean"]),
        "median": float(col.median()),
        "min_val": float(desc["min"]),
        "max_val": float(desc["max"]),
        "std": float(desc["std"]),
        "sum_val": float(col.sum()),
    }

    numeric_vals = col.dropna()
    if len(numeric_vals) >= 3:
        result["skewness"] = float(stats.skew(numeric_vals, bias=False))
        result["kurtosis"] = float(stats.kurtosis(numeric_vals, bias=False))
    else:
        result["skewness"] = None
        result["kurtosis"] = None

    return result


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


def vlookup_helper(
    lookup_file: str,
    data_file: str,
    lookup_column: str,
    data_key_column: str,
    data_return_columns: list[str],
    lookup_sheet: str = "Sheet1",
    data_sheet: str = "Sheet1",
    fuzzy: bool = False,
    fuzzy_threshold: float = 0.8,
    output_file: str | None = None,
    header_row: int = 1,
) -> dict:
    """Cross-file VLOOKUP-like join with optional fuzzy string matching."""
    from openpyxl.utils import column_index_from_string

    # Load lookup workbook
    lookup_wb = load_workbook_safe(lookup_file, read_only=True)
    lookup_ws = get_sheet(lookup_wb, lookup_sheet)
    lookup_col_idx = column_index_from_string(lookup_column)

    # Extract lookup values (skip header row)
    lookup_values: list[tuple[int, str | int | float | None]] = []
    for row in lookup_ws.iter_rows(min_row=header_row + 1, max_row=lookup_ws.max_row):
        cell = row[lookup_col_idx - 1] if lookup_col_idx - 1 < len(row) else None
        if cell is not None:
            lookup_values.append((cell.row, cell.value))
    lookup_wb.close()

    # Load data workbook
    data_wb = load_workbook_safe(data_file, read_only=True)
    data_ws = get_sheet(data_wb, data_sheet)
    data_key_idx = column_index_from_string(data_key_column)
    return_col_idxs = [column_index_from_string(c) for c in data_return_columns]

    # Extract header names from the data file
    header_cells = list(data_ws.iter_rows(min_row=header_row, max_row=header_row))[0]
    return_headers = []
    for idx in return_col_idxs:
        if idx - 1 < len(header_cells) and header_cells[idx - 1].value is not None:
            return_headers.append(str(header_cells[idx - 1].value))
        else:
            return_headers.append(f"Col_{idx}")

    # Build data lookup index: key_value -> list of return-column values
    data_rows: list[tuple[object, list[object]]] = []
    for row in data_ws.iter_rows(min_row=header_row + 1, max_row=data_ws.max_row):
        key_cell = row[data_key_idx - 1] if data_key_idx - 1 < len(row) else None
        key_val = key_cell.value if key_cell is not None else None
        ret_vals = []
        for idx in return_col_idxs:
            c = row[idx - 1] if idx - 1 < len(row) else None
            ret_vals.append(c.value if c is not None else None)
        data_rows.append((key_val, ret_vals))
    data_wb.close()

    # Perform matching
    results: list[dict] = []
    for _src_row, lv in lookup_values:
        if lv is None:
            continue

        best_match: dict | None = None

        if not fuzzy:
            # Exact match
            for dk, dvals in data_rows:
                if dk == lv:
                    best_match = {
                        "lookup_value": lv,
                        "matched_value": dk,
                        "return_values": dict(zip(return_headers, dvals)),
                    }
                    break
        else:
            # Fuzzy string matching
            lv_str = str(lv)
            best_score = 0.0
            for dk, dvals in data_rows:
                if dk is None:
                    continue
                score = difflib.SequenceMatcher(None, lv_str.lower(), str(dk).lower()).ratio()
                if score >= fuzzy_threshold and score > best_score:
                    best_score = score
                    best_match = {
                        "lookup_value": lv,
                        "matched_value": dk,
                        "similarity": round(best_score, 4),
                        "return_values": dict(zip(return_headers, dvals)),
                    }

        if best_match:
            results.append(best_match)
        else:
            results.append(
                {
                    "lookup_value": lv,
                    "matched_value": None,
                    "return_values": {h: None for h in return_headers},
                }
            )

    # Optionally write results to output file
    if output_file:
        validate_file_path(output_file, must_exist=False)
        out_wb = openpyxl.Workbook()
        out_ws = out_wb.active
        out_ws.title = "VLookup Results"

        # Write headers
        out_headers = ["Lookup Value", "Matched Value"]
        if fuzzy:
            out_headers.append("Similarity")
        out_headers.extend(return_headers)
        for c_idx, h in enumerate(out_headers, start=1):
            out_ws.cell(row=1, column=c_idx, value=h)

        # Write data rows
        for r_idx, row_data in enumerate(results, start=2):
            c = 1
            out_ws.cell(row=r_idx, column=c, value=row_data["lookup_value"])
            c += 1
            out_ws.cell(row=r_idx, column=c, value=row_data["matched_value"])
            c += 1
            if fuzzy:
                out_ws.cell(row=r_idx, column=c, value=row_data.get("similarity"))
                c += 1
            for h in return_headers:
                out_ws.cell(row=r_idx, column=c, value=row_data["return_values"].get(h))
                c += 1

        save_workbook_safe(out_wb, output_file)
        logger.info("VLookup results written to %s", output_file)

    return {
        "matched": sum(1 for r in results if r["matched_value"] is not None),
        "unmatched": sum(1 for r in results if r["matched_value"] is None),
        "total": len(results),
        "results": results,
        "output_file": output_file,
    }


def find_cells_by_format(
    file_path: str,
    sheet_name: str,
    bold: bool | None = None,
    italic: bool | None = None,
    fill_color: str | None = None,
    font_color: str | None = None,
    number_format: str | None = None,
) -> list[dict]:
    """Find cells matching specified formatting conditions.

    At least one condition must be provided.
    Returns list of {cell_ref, value, bold, italic, fill_color, font_color, number_format}.
    """
    if all(v is None for v in (bold, italic, fill_color, font_color, number_format)):
        raise ValueError("At least one formatting condition must be specified.")

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        results = []
        for row in ws.iter_rows():
            for cell in row:
                font = cell.font
                fill = cell.fill

                if bold is not None and font.bold != bold:
                    continue
                if italic is not None and font.italic != italic:
                    continue
                if fill_color is not None:
                    fg = fill.fgColor
                    cell_color = fg.rgb if fg.type == "rgb" else str(fg.value or "")
                    if fill_color.upper() not in cell_color.upper():
                        continue
                if font_color is not None:
                    fc = font.color
                    cell_font_color = (fc.rgb if fc and fc.type == "rgb" else "") or ""
                    if font_color.upper() not in cell_font_color.upper():
                        continue
                if number_format is not None and number_format not in (cell.number_format or ""):
                    continue

                fg = fill.fgColor
                cell_fill_rgb = fg.rgb if fg.type == "rgb" else str(fg.value or "")
                fc = font.color
                cell_font_rgb = (fc.rgb if fc and fc.type == "rgb" else "") or ""

                results.append(
                    {
                        "cell_ref": cell.coordinate,
                        "value": cell.value,
                        "bold": font.bold,
                        "italic": font.italic,
                        "fill_color": cell_fill_rgb,
                        "font_color": cell_font_rgb,
                        "number_format": cell.number_format,
                    }
                )

        logger.info("Found %d cells matching format criteria in %s!%s", len(results), sheet_name, file_path)
        return results
    finally:
        wb.close()


def normalize_data(
    file_path: str,
    sheet_name: str,
    columns: list[str],
    method: str = "min_max",
    output_sheet: str | None = None,
    has_header: bool = True,
) -> str:
    """Normalize numeric columns using min-max [0,1] or z-score normalization.

    Writes normalized values back to the sheet (in-place or to output_sheet).
    Returns a summary string.
    """
    if method not in ("min_max", "zscore"):
        raise ValueError(f"Unsupported method '{method}'. Allowed: 'min_max', 'zscore'.")

    df = _read_sheet_df(file_path, sheet_name, has_header)
    for col in columns:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found. Available: {list(df.columns)}")

    normalized_cols = []
    skipped_cols = []
    for col in columns:
        if not pd.api.types.is_numeric_dtype(df[col]):
            skipped_cols.append(col)
            continue
        if method == "min_max":
            col_min = df[col].min()
            col_max = df[col].max()
            if col_max == col_min:
                df[col] = 0.0
            else:
                df[col] = (df[col] - col_min) / (col_max - col_min)
        else:  # zscore
            col_mean = df[col].mean()
            col_std = df[col].std()
            if col_std == 0:
                df[col] = 0.0
            else:
                df[col] = (df[col] - col_mean) / col_std
        normalized_cols.append(col)

    target_sheet = output_sheet or sheet_name
    wb = load_workbook_safe(file_path)
    try:
        if target_sheet in wb.sheetnames:
            ws = wb[target_sheet]
            for row in ws.iter_rows():
                for cell in row:
                    cell.value = None
        else:
            ws = wb.create_sheet(target_sheet)

        if has_header:
            for c_idx, col_name in enumerate(df.columns, start=1):
                ws.cell(row=1, column=c_idx, value=col_name)
            for r_idx, row_data in enumerate(df.values.tolist(), start=2):
                for c_idx, val in enumerate(row_data, start=1):
                    ws.cell(row=r_idx, column=c_idx, value=val)
        else:
            for r_idx, row_data in enumerate(df.values.tolist(), start=1):
                for c_idx, val in enumerate(row_data, start=1):
                    ws.cell(row=r_idx, column=c_idx, value=val)

        save_workbook_safe(wb, file_path)
    finally:
        wb.close()
    logger.info("Normalized columns %s in '%s' using '%s'", normalized_cols, sheet_name, method)

    parts = [f"Normalized {len(normalized_cols)} column(s) using '{method}': {normalized_cols}."]
    if skipped_cols:
        parts.append(f"Skipped non-numeric: {skipped_cols}.")
    return " ".join(parts)


def export_analysis(
    data: list[dict] | list[list],
    output_file: str,
    sheet_name: str = "Analysis",
    headers: list[str] | None = None,
) -> str:
    """Write analysis results (list of dicts or list of lists) to a new Excel file."""
    validate_file_path(output_file, must_exist=False)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    if data and isinstance(data[0], dict):
        col_headers = headers or list(data[0].keys())
        ws.append(col_headers)
        for row in data:
            ws.append([row.get(h) for h in col_headers])
    else:
        if headers:
            ws.append(headers)
        for row in data:
            ws.append(list(row))

    save_workbook_safe(wb, output_file)
    logger.info("Exported %d rows to '%s'", len(data), output_file)
    return output_file
