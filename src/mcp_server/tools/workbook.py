"""Workbook-level operations: metadata, sheets, creation."""

from __future__ import annotations

from logging import Logger

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from mcp_server.models.workbook import (
    SheetCreatedInfo,
    SheetDefinition,
    SheetInfo,
    SheetSummary,
    WorkbookCreatedResult,
    WorkbookMetadata,
    WriteMultiSheetResult,
)
from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    read_sheet_df,
    save_workbook_safe,
    validate_file_path,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def get_workbook_metadata(file_path: str) -> WorkbookMetadata:
    """Return workbook metadata: sheets, active sheet, and named ranges."""
    wb = load_workbook_safe(file_path, read_only=True)
    try:
        sheets: list[SheetInfo] = []
        for ws in wb.worksheets:
            sheets.append(
                SheetInfo(
                    name=ws.title,
                    min_row=ws.min_row,
                    max_row=ws.max_row,
                    min_col=ws.min_column,
                    max_col=ws.max_column,
                )
            )

        named_ranges = [{"name": nr.name, "destination": str(nr.attr_text)} for nr in wb.defined_names.values()]

        return WorkbookMetadata(
            file_path=file_path,
            sheets=sheets,
            active_sheet=wb.active.title if wb.active else None,
            named_ranges=named_ranges,
        )
    finally:
        wb.close()


def create_workbook(
    file_path: str, sheet_names: list[str] | None = None, sheet_name: str | None = None
) -> WorkbookCreatedResult:
    """Create a new .xlsx workbook with optional sheet names."""
    validate_file_path(file_path, must_exist=False)
    wb = Workbook()

    if sheet_name and not sheet_names:
        sheet_names = [sheet_name]

    if sheet_names:
        for name in sheet_names:
            wb.create_sheet(title=name)
        if "Sheet" not in sheet_names and "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    save_workbook_safe(wb, file_path)
    logger.info("Created workbook: %s", file_path)
    return WorkbookCreatedResult(file_path=file_path, sheets=list(wb.sheetnames))


def get_sheet_summary(file_path: str, sheet_name: str) -> SheetSummary:
    """Return summary of a sheet: name, dimensions, headers, used range."""
    wb = load_workbook_safe(file_path, read_only=True)
    try:
        ws = get_sheet(wb, sheet_name)
        min_row = ws.min_row or 1
        min_col = ws.min_column or 1

        try:
            df = read_sheet_df(file_path, sheet_name)
            row_count = len(df) + 1  # +1 for header
            col_count = len(df.columns)
        except Exception:
            row_count = ws.max_row or 0
            col_count = ws.max_column or 0

        max_row = min_row + row_count - 1
        max_col = min_col + col_count - 1

        headers = []
        if col_count > 0:
            for row in ws.iter_rows(min_row=min_row, max_row=min_row, min_col=min_col, max_col=min_col + col_count - 1):
                headers = [cell.value for cell in row]

        used_range = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"

        return SheetSummary(
            name=ws.title,
            row_count=row_count,
            col_count=col_count,
            headers=[str(h) if h is not None else "" for h in headers],
            used_range=used_range,
        )
    finally:
        wb.close()


def rename_sheet(file_path: str, old_name: str, new_name: str) -> str:
    """Rename a worksheet."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, old_name)
        ws.title = new_name
        save_workbook_safe(wb, file_path)
        logger.info("Renamed sheet '%s' to '%s' in %s", old_name, new_name, file_path)
        return f"Sheet '{old_name}' renamed to '{new_name}'."
    finally:
        wb.close()


def delete_sheet(file_path: str, sheet_name: str) -> str:
    """Delete a worksheet. Raises ValueError if it's the only sheet."""
    wb = load_workbook_safe(file_path)
    try:
        if len(wb.sheetnames) == 1:
            raise ValueError("Cannot delete the only sheet in the workbook.")
        ws = get_sheet(wb, sheet_name)
        del wb[ws.title]
        save_workbook_safe(wb, file_path)
        logger.info("Deleted sheet '%s' from %s", sheet_name, file_path)
        return f"Sheet '{sheet_name}' deleted."
    finally:
        wb.close()


def copy_sheet(file_path: str, source_sheet: str, new_name: str) -> str:
    """Copy a sheet within the same workbook."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, source_sheet)
        target = wb.copy_worksheet(ws)
        target.title = new_name
        save_workbook_safe(wb, file_path)
        logger.info("Copied sheet '%s' as '%s' in %s", source_sheet, new_name, file_path)
        return f"Sheet '{source_sheet}' copied as '{new_name}'."
    finally:
        wb.close()


def write_multi_sheet(
    file_path: str,
    sheets: list[SheetDefinition],
) -> WriteMultiSheetResult:
    """Create a new workbook with multiple named sheets, data, and headers in one call."""
    validate_file_path(file_path, must_exist=False)
    wb = Workbook()

    # Remove the default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    summary: list[SheetCreatedInfo] = []
    for sheet_def in sheets:
        name = sheet_def["name"]
        headers: list[str] = sheet_def.get("headers") or []
        data: list[list[object]] | None = sheet_def.get("data") or sheet_def.get("values")
        column_widths: dict[str, float] | None = sheet_def.get("column_widths")

        ws = wb.create_sheet(title=name)
        current_row = 1

        if headers:
            for col_idx, header in enumerate(headers, start=1):
                cell = ws.cell(row=current_row, column=col_idx, value=header)
                cell.font = Font(bold=True)
            current_row += 1

        rows_written = 0
        if data:
            for row_data in data:
                for col_idx, value in enumerate(row_data, start=1):
                    ws.cell(row=current_row, column=col_idx, value=value)
                current_row += 1
                rows_written += 1

        if column_widths:
            for col_letter, width in column_widths.items():
                ws.column_dimensions[col_letter.upper()].width = width

        summary.append(
            SheetCreatedInfo(
                name=str(name),
                header_count=len(headers),
                row_count=rows_written,
                column_widths_set=list(column_widths.keys()) if column_widths else [],
            )
        )

    save_workbook_safe(wb, file_path)
    logger.info("Created multi-sheet workbook: %s with %d sheets", file_path, len(sheets))
    return WriteMultiSheetResult(file_path=file_path, sheets_created=summary)


def hide_sheet(file_path: str, sheet_name: str) -> dict[str, str]:
    """Hide a worksheet. Raises ValueError if it's the last visible sheet."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        visible_count = sum(1 for s in wb.worksheets if s.sheet_state == "visible")
        if visible_count <= 1 and ws.sheet_state == "visible":
            raise ValueError("Cannot hide the last visible sheet in the workbook.")
        ws.sheet_state = "hidden"
        save_workbook_safe(wb, file_path)
        logger.info("Hid sheet '%s' in %s", sheet_name, file_path)
        return {"status": "success", "message": f"Sheet '{sheet_name}' is now hidden"}
    finally:
        wb.close()


def unhide_sheet(file_path: str, sheet_name: str) -> dict[str, str]:
    """Unhide a hidden worksheet."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws.sheet_state = "visible"
        save_workbook_safe(wb, file_path)
        logger.info("Unhid sheet '%s' in %s", sheet_name, file_path)
        return {"status": "success", "message": f"Sheet '{sheet_name}' is now visible"}
    finally:
        wb.close()
