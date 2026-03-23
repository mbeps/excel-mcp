"""Excel table (ListObject) operations."""

from __future__ import annotations

from logging import Logger

from openpyxl.worksheet.table import Table, TableStyleInfo

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


def list_tables(file_path: str, sheet_name: str) -> list[dict]:
    """List all tables in a sheet."""
    wb = load_workbook_safe(file_path, read_only=False)
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
