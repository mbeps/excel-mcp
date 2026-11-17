"""Pivot table, ETL, and data transformation operations."""

from __future__ import annotations

import ast
import operator
from logging import Logger
from typing import Any, Literal

import numpy as np
import pandas as pd
from openpyxl import Workbook

from mcp_server.models.common import CellScalar
from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_hidden_json,
    load_workbook_safe,
    read_sheet_df,
    sanitize_excel_value,
    save_hidden_json,
    save_workbook_safe,
    write_dataframe_to_workbook,
)
from mcp_server.utils.expression_validator import SAFE_MATH_FUNCS, validate_expression
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

_PIVOTS_SHEET = "_mcp_pivots"


def _load_pivots(wb: Workbook) -> dict:
    """Load pivot definitions from the hidden pivots sheet."""
    return load_hidden_json(wb, _PIVOTS_SHEET)


def _save_pivots(wb: Workbook, pivots: dict) -> None:
    """Save pivot definitions dict to hidden sheet."""
    save_hidden_json(wb, _PIVOTS_SHEET, pivots)


_BIN_OPS: dict[type[ast.operator], Any] = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.FloorDiv: operator.floordiv,
    ast.Mod: operator.mod,
    ast.Pow: operator.pow,
}

_CMP_OPS: dict[type[ast.cmpop], Any] = {
    ast.Gt: operator.gt,
    ast.Lt: operator.lt,
    ast.GtE: operator.ge,
    ast.LtE: operator.le,
    ast.Eq: operator.eq,
    ast.NotEq: operator.ne,
}

_VECTOR_MATH_FUNCS: dict[str, Any] = {
    "sqrt": np.sqrt,
    "log": np.log,
    "log10": np.log10,
    "exp": np.exp,
    "sin": np.sin,
    "cos": np.cos,
    "tan": np.tan,
    "abs": np.abs,
    "pow": np.power,
}


def _eval_ast_node(node: ast.AST, df: pd.DataFrame) -> Any:
    """Safely evaluate a validated AST node against a DataFrame."""
    if isinstance(node, ast.Expression):
        return _eval_ast_node(node.body, df)
    if isinstance(node, ast.Constant):
        return node.value
    if isinstance(node, ast.Name):
        if node.id in df.columns:
            return df[node.id]
        raise ValueError(f"Unknown column '{node.id}' in expression.")
    if isinstance(node, ast.UnaryOp):
        operand = _eval_ast_node(node.operand, df)
        return -operand if isinstance(node.op, ast.USub) else operand
    if isinstance(node, ast.BinOp):
        left = _eval_ast_node(node.left, df)
        right = _eval_ast_node(node.right, df)
        return _BIN_OPS[type(node.op)](left, right)
    if isinstance(node, ast.Compare):
        left = _eval_ast_node(node.left, df)
        result: Any = None
        for op, comp in zip(node.ops, node.comparators):
            right = _eval_ast_node(comp, df)
            step = _CMP_OPS[type(op)](left, right)
            result = step if result is None else (result & step)
            left = right
        return result
    if isinstance(node, ast.BoolOp):
        acc = _eval_ast_node(node.values[0], df)
        for val_node in node.values[1:]:
            rhs = _eval_ast_node(val_node, df)
            acc = (acc & rhs) if isinstance(node.op, ast.And) else (acc | rhs)
        return acc
    if isinstance(node, ast.IfExp):
        cond = _eval_ast_node(node.test, df)
        return np.where(cond, _eval_ast_node(node.body, df), _eval_ast_node(node.orelse, df))
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in SAFE_MATH_FUNCS:
        fn = _VECTOR_MATH_FUNCS.get(node.func.id, SAFE_MATH_FUNCS[node.func.id])
        args = [_eval_ast_node(a, df) for a in node.args]
        return fn(*args)  # type: ignore[operator]
    raise ValueError(f"Unsupported AST construct in expression: {type(node).__name__}")


def _records_sanitized(df: pd.DataFrame) -> list[dict[str, CellScalar]]:
    """Convert DataFrame to records with NaN/Timestamp sanitized."""
    return [{str(k): sanitize_excel_value(v) for k, v in rec.items()} for rec in df.to_dict(orient="records")]


def _write_df_to_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame) -> None:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)

    # Write headers
    for c_idx, col_name in enumerate(df.columns, start=1):
        ws.cell(row=1, column=c_idx, value=sanitize_excel_value(col_name))

    # Write data
    for r_idx, row_data in enumerate(df.values.tolist(), start=2):
        for c_idx, val in enumerate(row_data, start=1):
            ws.cell(row=r_idx, column=c_idx, value=sanitize_excel_value(val))


def create_pivot_table(
    file_path: str,
    sheet_name: str,
    index_cols: list[str],
    value_cols: list[str],
    aggfunc: str | dict = "sum",
    output_sheet: str | None = None,
    output_file: str | None = None,
    include_margins: bool = False,
    column_field: str | None = None,
    date_freq: str | None = None,
) -> dict[str, list[dict[str, CellScalar]] | list[str] | str | dict | None]:
    """Create a static pivot table using pandas and write to a sheet or file."""
    df = read_sheet_df(file_path, sheet_name)

    all_required_cols = index_cols + value_cols + ([column_field] if column_field else [])
    for col in all_required_cols:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found. Available: {list(df.columns)}")

    # Apply date period grouping if requested
    if date_freq:
        new_index: list = []
        for col in index_cols:
            try:
                df[col] = pd.to_datetime(df[col], errors="coerce")
                if df[col].notna().any():
                    new_index.append(pd.Grouper(key=col, freq=date_freq))
                else:
                    new_index.append(col)
            except Exception:
                new_index.append(col)
        pivot_index = new_index
    else:
        pivot_index = index_cols

    pivot_kwargs: dict = {"margins": True, "margins_name": "Total"} if include_margins else {}
    if column_field:
        pivot_kwargs["columns"] = column_field
    pivot = pd.pivot_table(df, index=pivot_index, values=value_cols, aggfunc=aggfunc, **pivot_kwargs)
    pivot = pivot.reset_index()

    # Flatten multi-level column headers produced by column_field
    if column_field:
        pivot.columns = [
            "_".join(str(c) for c in col).strip("_") if isinstance(col, tuple) else str(col) for col in pivot.columns
        ]

    if output_file:
        write_dataframe_to_workbook(
            pivot,
            file_path,
            sheet_name=output_sheet or "Sheet1",
            output_file=output_file,
            start_row=1,
            include_header=True,
        )
        logger.info("Pivot table written to %s", output_file)
    elif output_sheet:
        wb = load_workbook_safe(file_path)
        try:
            _write_df_to_sheet(wb, output_sheet, pivot)
            pivots = _load_pivots(wb)
            pivots[output_sheet] = {
                "file_path": file_path,
                "sheet_name": sheet_name,
                "index_cols": index_cols,
                "value_cols": value_cols,
                "aggfunc": aggfunc,
                "output_sheet": output_sheet,
                "include_margins": include_margins,
                "column_field": column_field,
                "date_freq": date_freq,
            }
            _save_pivots(wb, pivots)
            save_workbook_safe(wb, file_path)
            logger.info("Pivot table written to sheet '%s' in %s", output_sheet, file_path)
        finally:
            wb.close()

    return {
        "data": _records_sanitized(pivot),
        "index_columns": index_cols,
        "value_columns": value_cols,
        "column_field": column_field,
        "operation": aggfunc,
    }


def refresh_pivot_table(
    file_path: str,
    output_sheet: str,
    source_file_path: str | None = None,
    source_sheet: str | None = None,
) -> dict:
    """Refresh a pivot table by re-running its stored definition.

    Reads the pivot parameters persisted in the hidden '_mcp_pivots' sheet
    and re-invokes create_pivot_table() to regenerate the snapshot.
    """
    wb = load_workbook_safe(file_path)
    try:
        pivots = _load_pivots(wb)
        if output_sheet not in pivots:
            raise ValueError(f"No pivot definition found for sheet '{output_sheet}'. Available: {list(pivots.keys())}")
        params = dict(pivots[output_sheet])
    finally:
        wb.close()

    if source_file_path is not None:
        params["file_path"] = source_file_path
    if source_sheet is not None:
        params["sheet_name"] = source_sheet

    result = create_pivot_table(
        file_path=params["file_path"],
        sheet_name=params["sheet_name"],
        index_cols=params["index_cols"],
        value_cols=params["value_cols"],
        aggfunc=params.get("aggfunc", "sum"),
        output_sheet=params["output_sheet"],
        output_file=None,
        include_margins=params.get("include_margins", False),
        column_field=params.get("column_field"),
        date_freq=params.get("date_freq"),
    )

    data = result.get("data")
    rows = len(data) if isinstance(data, list) else 0
    logger.info("Refreshed pivot '%s' in %s (%d rows)", output_sheet, file_path, rows)
    return {
        "refreshed": output_sheet,
        "source_file": params["file_path"],
        "source_sheet": params["sheet_name"],
        "rows": rows,
    }


def unpivot_data(
    file_path: str,
    sheet_name: str,
    id_vars: list[str],
    value_vars: list[str],
    var_name: str = "Variable",
    value_name: str = "Value",
) -> dict[str, list[dict[str, CellScalar]] | int]:
    """Unpivot (melt) data from wide to long format."""
    df = read_sheet_df(file_path, sheet_name)

    for col in id_vars + value_vars:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found. Available: {list(df.columns)}")

    melted = pd.melt(df, id_vars=id_vars, value_vars=value_vars, var_name=var_name, value_name=value_name)
    return {
        "data": _records_sanitized(melted),
        "row_count": len(melted),
    }


def merge_datasets(
    file_path: str,
    sheet1: str,
    sheet2: str,
    join_key: str | list[str] | None = None,
    how: str = "left",
    output_sheet: str | None = None,
    suffixes: tuple[str, str] = ("_x", "_y"),
    left_on: str | list[str] | None = None,
    right_on: str | list[str] | None = None,
) -> dict[str, list[dict[str, CellScalar]] | int]:
    """Merge two sheets like a SQL join."""
    valid_how = {"left", "right", "inner", "outer"}
    if how not in valid_how:
        raise ValueError(f"Unsupported join type '{how}'. Allowed: {valid_how}")

    if left_on is not None and right_on is None:
        raise ValueError("'right_on' is required when 'left_on' is provided.")
    if right_on is not None and left_on is None:
        raise ValueError("'left_on' is required when 'right_on' is provided.")
    if join_key is None and left_on is None:
        raise ValueError("Either 'join_key' or both 'left_on'/'right_on' must be provided.")

    df1 = read_sheet_df(file_path, sheet1)
    df2 = read_sheet_df(file_path, sheet2)

    if left_on is not None and right_on is not None:
        lkeys = [left_on] if isinstance(left_on, str) else left_on
        rkeys = [right_on] if isinstance(right_on, str) else right_on
        for k in lkeys:
            if k not in df1.columns:
                raise ValueError(f"Key '{k}' not found in sheet '{sheet1}'. Available: {list(df1.columns)}")
        for k in rkeys:
            if k not in df2.columns:
                raise ValueError(f"Key '{k}' not found in sheet '{sheet2}'. Available: {list(df2.columns)}")
        if df1[lkeys].duplicated().any() or df2[rkeys].duplicated().any():
            logger.warning("Merge keys are not unique in one or both datasets. Result may contain unexpected rows.")
        merged = pd.merge(df1, df2, left_on=lkeys, right_on=rkeys, how=how, suffixes=suffixes)
    else:
        # join_key is guaranteed non-None here (checked above)
        keys: list[str] = [join_key] if isinstance(join_key, str) else list(join_key or [])
        for k in keys:
            if k not in df1.columns:
                raise ValueError(f"Key '{k}' not found in sheet '{sheet1}'. Available: {list(df1.columns)}")
            if k not in df2.columns:
                raise ValueError(f"Key '{k}' not found in sheet '{sheet2}'. Available: {list(df2.columns)}")
        if df1[keys].duplicated().any() or df2[keys].duplicated().any():
            logger.warning("Merge keys are not unique in one or both datasets. Result may contain unexpected rows.")
        merged = pd.merge(df1, df2, on=keys, how=how, suffixes=suffixes)

    if output_sheet:
        wb = load_workbook_safe(file_path)
        try:
            _write_df_to_sheet(wb, output_sheet, merged)
            save_workbook_safe(wb, file_path)
            logger.info("Merged result written to sheet '%s'", output_sheet)
        finally:
            wb.close()

    return {
        "data": _records_sanitized(merged),
        "row_count": len(merged),
    }


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
    """Add a computed column using a pandas-eval expression or cumulative sum.

    column_type='formula' (default): evaluate ``expression`` via pandas eval.
    column_type='cumsum': compute a running total of ``source_col``.
    """
    if column_type == "cumsum":
        if source_col is None:
            raise ValueError("source_col is required when column_type='cumsum'")
        df = read_sheet_df(file_path, sheet_name, header_row=1 if has_header else 0)
        if source_col not in df.columns:
            raise ValueError(f"source_col '{source_col}' not found. Available: {list(df.columns)}")
        df[new_column_name] = df[source_col].cumsum()

        wb = load_workbook_safe(file_path)
        try:
            ws = get_sheet(wb, sheet_name)

            new_col_idx = list(df.columns).index(new_column_name) + 1
            if has_header:
                ws.cell(row=1, column=new_col_idx, value=new_column_name)

            start_row = 2 if has_header else 1
            for r_idx, val in enumerate(df[new_column_name].tolist(), start=start_row):
                ws.cell(row=r_idx, column=new_col_idx, value=sanitize_excel_value(val))

            save_workbook_safe(wb, file_path)
            logger.info(
                "Added cumsum column '%s' from '%s' to %s!%s", new_column_name, source_col, sheet_name, file_path
            )
            return {
                "status": "ok",
                "column_type": "cumsum",
                "new_column": new_column_name,
                "source_col": source_col,
                "rows_written": len(df),
                "message": f"Added cumsum column '{new_column_name}' from source column '{source_col}' to sheet '{sheet_name}' ({len(df)} rows written).",
            }
        finally:
            wb.close()

    if column_type == "rolling":
        if source_col is None:
            raise ValueError("source_col is required for column_type='rolling'.")
        if window is None or isinstance(window, bool) or window < 1:
            raise ValueError("window must be an integer >= 1 for column_type='rolling'.")
        if rolling_func not in ("mean", "sum"):
            raise ValueError("rolling_func must be 'mean' or 'sum'.")
        df = read_sheet_df(file_path, sheet_name, header_row=1 if has_header else 0)
        if source_col not in df.columns:
            raise ValueError(f"source_col '{source_col}' not found in sheet columns.")

        roller = df[source_col].rolling(window=window)
        rolled = getattr(roller, rolling_func)()

        wb2 = load_workbook_safe(file_path)
        try:
            ws2 = get_sheet(wb2, sheet_name)
            header_offset = 1 if has_header else 0
            new_col_idx = (
                list(df.columns).index(new_column_name) + 1 if new_column_name in df.columns else len(df.columns) + 1
            )
            if has_header:
                ws2.cell(row=1, column=new_col_idx, value=new_column_name)
            for i, val in enumerate(rolled):
                row_num = i + 1 + header_offset
                ws2.cell(row=row_num, column=new_col_idx, value=None if pd.isna(val) else round(float(val), 6))
            save_workbook_safe(wb2, file_path)
        finally:
            wb2.close()

        logger.info(
            "Rolling %s (window=%d) column '%s' added in '%s'", rolling_func, window, new_column_name, sheet_name
        )
        return f"Added rolling {rolling_func} column '{new_column_name}' (window={window}) from source column '{source_col}' to sheet '{sheet_name}'."

    # column_type == "formula" (default path)
    df = read_sheet_df(file_path, sheet_name, header_row=1 if has_header else 0)
    tree = validate_expression(expression)
    refs_column = any(isinstance(node, ast.Name) and node.id in df.columns for node in ast.walk(tree))

    try:
        df[new_column_name] = df.eval(expression, engine="numexpr")
    except (ImportError, NotImplementedError):
        try:
            df[new_column_name] = _eval_ast_node(tree, df)
        except Exception as e:
            raise ValueError(f"Failed to evaluate expression '{expression}': {e}") from e
    except Exception as e:
        raise ValueError(f"Failed to evaluate expression '{expression}': {e}") from e

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        new_col_idx = list(df.columns).index(new_column_name) + 1
        if has_header:
            ws.cell(row=1, column=new_col_idx, value=new_column_name)

        start_row = 2 if has_header else 1
        mode = "formula" if refs_column else "static"
        for r_idx, val in enumerate(df[new_column_name].tolist(), start=start_row):
            ws.cell(row=r_idx, column=new_col_idx, value=sanitize_excel_value(val))

        save_workbook_safe(wb, file_path)
        logger.info("Added column '%s' to %s!%s (mode=%s)", new_column_name, sheet_name, file_path, mode)
        return f"Added column '{new_column_name}' ({expression}) to '{sheet_name}' ({len(df)} rows, mode={mode})."
    finally:
        wb.close()


def deduplicate_data(
    file_path: str,
    sheet_name: str,
    columns: list[str] | None = None,
    keep: str | bool = "first",
) -> str:
    """Remove duplicate rows from a sheet."""
    if isinstance(keep, bool):
        if keep:
            raise ValueError("Invalid keep value True. Allowed: 'first', 'last', 'none', or False.")
        effective_keep: Literal["first", "last", False] = False
    elif keep == "none":
        effective_keep = False
    elif keep in ("first", "last"):
        effective_keep = keep  # type: ignore[assignment]
    else:
        raise ValueError(f"Invalid keep value '{keep}'. Allowed: 'first', 'last', 'none', or False.")

    df = read_sheet_df(file_path, sheet_name)

    if columns:
        for c in columns:
            if c not in df.columns:
                raise ValueError(f"Column '{c}' not found. Available: {list(df.columns)}")

    original_count = len(df)
    df = df.drop_duplicates(subset=columns, keep=effective_keep)
    removed = original_count - len(df)

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        # Clear existing data below header
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                cell.value = None

        for r_idx, row_data in enumerate(df.values.tolist(), start=2):
            for c_idx, val in enumerate(row_data, start=1):
                ws.cell(row=r_idx, column=c_idx, value=sanitize_excel_value(val))

        save_workbook_safe(wb, file_path)
        logger.info("Removed %d duplicates from %s!%s", removed, sheet_name, file_path)
        return f"Removed {removed} duplicate row(s) from '{sheet_name}'. {len(df)} rows remain."
    finally:
        wb.close()
