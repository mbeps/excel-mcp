"""Excel table (ListObject) operations."""

from __future__ import annotations

import re
from logging import Logger
from typing import Any

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

from mcp_server.models.common import CellScalar
from mcp_server.models.tables import TableInfo
from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


_TABLE_NAME_RE = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.]*$")


def _get_table(ws: Any, table_name: str) -> tuple[str, Table]:
    """Locate a table on *ws* by dict key or ``displayName``."""
    if table_name in ws.tables:
        return table_name, ws.tables[table_name]
    for tbl in ws.tables.values():
        if tbl.displayName == table_name:
            key = tbl.name or tbl.displayName
            return key, tbl
    raise ValueError(f"Table '{table_name}' not found in sheet '{ws.title}'.")


def _ranges_overlap(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    """Return True if two (min_col, min_row, max_col, max_row) rectangles overlap."""
    return not (a[2] < b[0] or a[0] > b[2] or a[3] < b[1] or a[1] > b[3])


def create_table(
    file_path: str,
    sheet_name: str,
    data_range: str,
    table_name: str,
    style_name: str = "TableStyleMedium9",
) -> str:
    """Create a native Excel table (ListObject)."""
    if not table_name or not _TABLE_NAME_RE.match(table_name):
        raise ValueError(
            f"Invalid table name '{table_name}'. Must start with a letter, '_', or '\\' and contain no spaces."
        )
    bounds = range_boundaries(data_range)
    if any(b is None for b in bounds):
        raise ValueError(f"Invalid table range '{data_range}': must be a bounded A1 range.")
    new_rect = (int(bounds[0]), int(bounds[1]), int(bounds[2]), int(bounds[3]))
    if new_rect[0] > new_rect[2] or new_rect[1] > new_rect[3]:
        raise ValueError(f"Invalid table range '{data_range}': reversed bounds.")

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        lower_name = table_name.lower()
        for existing in ws.tables.values():
            if (existing.displayName or "").lower() == lower_name or (existing.name or "").lower() == lower_name:
                raise ValueError(f"Table '{table_name}' already exists on sheet '{sheet_name}'.")
            ex_bounds = range_boundaries(existing.ref)
            ex_rect = (int(ex_bounds[0]), int(ex_bounds[1]), int(ex_bounds[2]), int(ex_bounds[3]))
            if _ranges_overlap(new_rect, ex_rect):
                raise ValueError(
                    f"Range '{data_range}' overlaps existing table '{existing.displayName}' ({existing.ref})."
                )

        style = TableStyleInfo(
            name=style_name,
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        table = Table(displayName=table_name, ref=data_range)
        table.tableStyleInfo = style
        ws.add_table(table)

        save_workbook_safe(wb, file_path)
        logger.info("Created table '%s' at %s in %s", table_name, data_range, file_path)
        return f"Created table '{table_name}' at '{data_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()


def list_tables(file_path: str, sheet_name: str) -> list[TableInfo]:
    """List all tables in a sheet."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        return [
            {
                "name": table.displayName,
                "ref": table.ref,
                "style": table.tableStyleInfo.name if table.tableStyleInfo else None,
            }
            for table in ws.tables.values()
        ]
    finally:
        wb.close()


def resize_table(file_path: str, sheet_name: str, table_name: str, new_range: str) -> str:
    """Change the cell reference range of a table and synchronize tableColumns."""
    bounds = range_boundaries(new_range)
    if any(b is None for b in bounds):
        raise ValueError(f"Invalid table range '{new_range}': must be a bounded A1 range.")
    min_col, min_row, max_col, max_row = int(bounds[0]), int(bounds[1]), int(bounds[2]), int(bounds[3])
    if min_col > max_col or min_row > max_row:
        raise ValueError(f"Invalid table range '{new_range}': reversed bounds.")

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        _, table = _get_table(ws, table_name)
        old_ref = table.ref
        table.ref = new_range

        new_col_count = max_col - min_col + 1
        old_cols = list(table.tableColumns or [])
        synced_cols: list[TableColumn] = []
        used_names: set[str] = set()
        for i in range(new_col_count):
            header_val = ws.cell(row=min_row, column=min_col + i).value
            if header_val is not None and str(header_val).strip():
                base_name = str(header_val).strip()
            elif i < len(old_cols) and old_cols[i].name:
                base_name = old_cols[i].name
            else:
                base_name = f"Column{i + 1}"

            col_name = base_name
            suffix = 2
            while col_name in used_names:
                col_name = f"{base_name}{suffix}"
                suffix += 1
            used_names.add(col_name)

            prev = old_cols[i] if i < len(old_cols) else None
            synced_cols.append(
                TableColumn(
                    id=i + 1,
                    name=col_name,
                    totalsRowFunction=prev.totalsRowFunction if prev else None,
                    totalsRowLabel=prev.totalsRowLabel if prev else None,
                )
            )
        table.tableColumns = synced_cols

        save_workbook_safe(wb, file_path)
        logger.info("Resized table '%s' from %s to %s in %s", table_name, old_ref, new_range, file_path)
        return f"Resized table '{table_name}' from '{old_ref}' to '{new_range}' on sheet '{sheet_name}'."
    finally:
        wb.close()


def set_table_totals_row(
    file_path: str,
    sheet_name: str,
    table_name: str,
    show_totals: bool,
    column_totals: dict[str, str] | None = None,
) -> str:
    """Toggle the totals row and set per-column aggregate functions.

    Valid function names: sum, count, average, max, min, countNums, stdDev, var, none.
    """
    _SUBTOTAL_IDS = {
        "average": 101,
        "count": 103,
        "countNums": 102,
        "max": 104,
        "min": 105,
        "stdDev": 107,
        "sum": 109,
        "var": 110,
    }

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        _, table = _get_table(ws, table_name)

        min_col, min_row, max_col, max_row = range_boundaries(table.ref)

        if show_totals:
            # Only extend ref if table doesn't already have a totals row
            if not table.totalsRowCount:
                totals_row = max_row + 1
                table.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{totals_row}"
            else:
                totals_row = max_row  # ref already includes totals row

            table.totalsRowCount = 1
            table.totalsRowShown = True

            if column_totals and table.tableColumns:
                for col_idx, col in enumerate(table.tableColumns, start=min_col):
                    if col.name in column_totals:
                        func = column_totals[col.name]
                        if func == "none":
                            col.totalsRowFunction = None
                            ws.cell(row=totals_row, column=col_idx).value = None
                        elif func in _SUBTOTAL_IDS:
                            col.totalsRowFunction = func
                            ws.cell(
                                row=totals_row, column=col_idx
                            ).value = f"=SUBTOTAL({_SUBTOTAL_IDS[func]},[{col.name}])"
                        else:
                            raise ValueError(
                                f"Unsupported totals function '{func}'. Valid: {sorted(_SUBTOTAL_IDS) + ['none']}."
                            )
        else:
            # Disable totals
            if table.totalsRowCount:
                # Clear cells in the totals row and shrink ref
                for c in range(min_col, max_col + 1):
                    ws.cell(row=max_row, column=c).value = None
                table.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row - 1}"
                if table.tableColumns:
                    for col in table.tableColumns:
                        col.totalsRowFunction = None
            table.totalsRowCount = None
            table.totalsRowShown = False

        save_workbook_safe(wb, file_path)
        logger.info("Set totals row for table '%s': show=%s", table_name, show_totals)
        return f"Totals row {'enabled' if show_totals else 'disabled'} for table '{table_name}'."
    finally:
        wb.close()


def get_table_data(
    file_path: str,
    sheet_name: str,
    table_name: str,
) -> dict[str, str | int | list[CellScalar] | list[list[CellScalar]] | None]:
    """Read table data as structured output with headers and rows."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        _, table = _get_table(ws, table_name)
        cell_rows = list(ws[table.ref])
        end_idx = len(cell_rows) - (table.totalsRowCount or 0)
        headers = [cell.value for cell in cell_rows[0]]
        rows = [[cell.value for cell in row] for row in cell_rows[1:end_idx]]
        logger.info("Read table '%s' data: %d rows from %s", table_name, len(rows), file_path)
        return {
            "table_name": table_name,
            "ref": table.ref,
            "headers": headers,
            "rows": rows,
            "row_count": len(rows),
        }
    finally:
        wb.close()


def convert_table_to_range(file_path: str, sheet_name: str, table_name: str) -> dict:
    """Remove a table definition while preserving all cell data and formatting.

    Cells containing structured references (e.g. ``=SUBTOTAL(109,[Amount])``)
    are replaced with their cached computed values before the table is deleted,
    because openpyxl cannot resolve structured references once the table
    definition is removed.
    """
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        table_key, table = _get_table(ws, table_name)

        table_ref = table.ref
        min_col, min_row, max_col, max_row = range_boundaries(table_ref)

        # Collect cells with structured references before deleting the table
        inside_pattern = re.compile(r"\[.*?\]")
        table_names = {n for n in (table_key, table.displayName) if n}
        outside_pattern = re.compile(r"(?:" + "|".join(re.escape(n) for n in table_names) + r")\[.*?\]")
        converted_cells: list[str] = []

        # Open a data_only copy to read cached values
        wb_data = load_workbook_safe(file_path, data_only=True)
        try:
            ws_data = get_sheet(wb_data, sheet_name)
            for row in ws.iter_rows():
                for cell in row:
                    val = cell.value
                    if not (isinstance(val, str) and val.startswith("=")):
                        continue
                    in_table = (min_row <= cell.row <= max_row) and (min_col <= cell.column <= max_col)
                    pattern = inside_pattern if in_table else outside_pattern
                    if pattern.search(val):
                        cached = ws_data.cell(row=cell.row, column=cell.column).value
                        if cached is not None:
                            cell.value = cached
                            converted_cells.append(cell.coordinate)
        finally:
            wb_data.close()

        del ws.tables[table_key]
        save_workbook_safe(wb, file_path)
        logger.info("Converted table '%s' to range in %s", table_name, sheet_name)

        result: dict = {"status": "ok", "sheet": sheet_name, "table": table_name}
        if converted_cells:
            result["structured_refs_converted"] = converted_cells
            result["note"] = (
                f"{len(converted_cells)} cell(s) contained structured references and were "
                "replaced with their cached computed values."
            )
        return result
    finally:
        wb.close()
