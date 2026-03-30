"""Excel table (ListObject) operations."""

from __future__ import annotations

from logging import Logger

from openpyxl.worksheet.table import Table, TableStyleInfo

from mcp_server.models.common import CellScalar
from mcp_server.models.tables import TableInfo
from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def create_table(
    file_path: str,
    sheet_name: str,
    data_range: str,
    table_name: str,
    style_name: str = "TableStyleMedium9",
) -> str:
    """Create a native Excel table (ListObject)."""
    wb = load_workbook_safe(file_path)
    ws = get_sheet(wb, sheet_name)

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
    """Change the cell reference range of a table."""
    wb = load_workbook_safe(file_path)
    ws = get_sheet(wb, sheet_name)
    if table_name not in ws.tables:
        raise ValueError(f"Table '{table_name}' not found in sheet '{sheet_name}'.")
    old_ref = ws.tables[table_name].ref
    ws.tables[table_name].ref = new_range
    save_workbook_safe(wb, file_path)
    logger.info("Resized table '%s' from %s to %s in %s", table_name, old_ref, new_range, file_path)
    return f"Resized table '{table_name}' from '{old_ref}' to '{new_range}' on sheet '{sheet_name}'."


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
    wb = load_workbook_safe(file_path)
    ws = get_sheet(wb, sheet_name)
    if table_name not in ws.tables:
        raise ValueError(f"Table '{table_name}' not found in sheet '{sheet_name}'.")
    table = ws.tables[table_name]
    table.totalsRowCount = 1 if show_totals else None
    if show_totals and column_totals and table.tableColumns:
        for col in table.tableColumns:
            if col.name in column_totals:
                col.totalsRowFunction = column_totals[col.name]
    save_workbook_safe(wb, file_path)
    logger.info("Set totals row for table '%s': show=%s", table_name, show_totals)
    return f"Totals row {'enabled' if show_totals else 'disabled'} for table '{table_name}'."


def get_table_data(
    file_path: str,
    sheet_name: str,
    table_name: str,
) -> dict[str, str | int | list[CellScalar] | list[list[CellScalar]] | None]:
    """Read table data as structured output with headers and rows."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        if table_name not in ws.tables:
            raise ValueError(f"Table '{table_name}' not found in sheet '{sheet_name}'.")
        table = ws.tables[table_name]
        cell_rows = list(ws[table.ref])
        headers = [cell.value for cell in cell_rows[0]]
        rows = [[cell.value for cell in row] for row in cell_rows[1:]]
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
