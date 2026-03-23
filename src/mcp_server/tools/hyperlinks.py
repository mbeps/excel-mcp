"""Hyperlink operations: add, read, delete, list."""

from __future__ import annotations

from logging import Logger

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def add_hyperlink(
    file_path: str,
    sheet_name: str,
    cell_ref: str,
    url: str,
    display_text: str | None = None,
    tooltip: str | None = None,
) -> str:
    """Add a hyperlink to a cell."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        cell = ws[cell_ref]
        cell.hyperlink = url
        if display_text is not None:
            cell.value = display_text
        if tooltip is not None:
            cell.hyperlink.tooltip = tooltip
        save_workbook_safe(wb, file_path)
        logger.info("Added hyperlink to %s!%s in %s", sheet_name, cell_ref, file_path)
        return f"Hyperlink added to cell {cell_ref} on sheet '{sheet_name}'."
    finally:
        wb.close()


def read_hyperlink(file_path: str, sheet_name: str, cell_ref: str) -> dict | None:
    """Read hyperlink from a cell. Returns dict with target/location/tooltip/display_text or None."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        cell = ws[cell_ref]
        if cell.hyperlink is None:
            return None
        hl = cell.hyperlink
        return {
            "target": hl.target,
            "location": hl.location,
            "tooltip": hl.tooltip,
            "display_text": cell.value,
        }
    finally:
        wb.close()


def delete_hyperlink(file_path: str, sheet_name: str, cell_ref: str) -> str:
    """Delete hyperlink from a cell."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws[cell_ref].hyperlink = None
        save_workbook_safe(wb, file_path)
        logger.info("Deleted hyperlink from %s!%s in %s", sheet_name, cell_ref, file_path)
        return f"Hyperlink deleted from cell {cell_ref} on sheet '{sheet_name}'."
    finally:
        wb.close()


def list_hyperlinks(file_path: str, sheet_name: str) -> list[dict]:
    """List all hyperlinks in a sheet."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        results = []
        for row in ws.iter_rows():
            for cell in row:
                if cell.hyperlink is not None:
                    hl = cell.hyperlink
                    results.append(
                        {
                            "cell_ref": cell.coordinate,
                            "target": hl.target,
                            "location": hl.location,
                            "tooltip": hl.tooltip,
                        }
                    )
        return results
    finally:
        wb.close()


def add_internal_hyperlink(
    file_path: str,
    sheet_name: str,
    cell: str,
    target_sheet: str,
    target_cell: str = "A1",
    display_text: str | None = None,
) -> str:
    """Add a hyperlink that navigates to another sheet/cell within the same workbook."""
    wb = load_workbook_safe(file_path)
    try:
        if target_sheet not in wb.sheetnames:
            raise ValueError(f"Sheet '{target_sheet}' not found in workbook")
        ws = get_sheet(wb, sheet_name)
        c = ws[cell]
        c.hyperlink = f"#{target_sheet}!{target_cell}"
        if display_text is not None:
            c.value = display_text
        save_workbook_safe(wb, file_path)
        logger.info(
            "Added internal link from %s!%s to %s!%s in %s",
            sheet_name,
            cell,
            target_sheet,
            target_cell,
            file_path,
        )
        return f"Added internal link from {cell} to {target_sheet}!{target_cell}"
    finally:
        wb.close()
