from __future__ import annotations

from typing import Literal

from mcp.types import ToolAnnotations

import mcp_server.tools.workbook as _workbook
from mcp_server.models.workbook import (
    SheetDefinition,
    SheetSummary,
    WorkbookCreatedResult,
    WorkbookMetadata,
    WriteMultiSheetResult,
)

__all__ = [
    "get_workbook_metadata",
    "create_workbook",
    "get_sheet_summary",
    "write_multi_sheet",
    "sheet_management",
]


def get_workbook_metadata(file_path: str) -> WorkbookMetadata:
    """Get workbook metadata including sheet names, dimensions, active sheet, and named ranges."""
    return _workbook.get_workbook_metadata(file_path)


def create_workbook(
    file_path: str, sheet_names: list[str] | None = None, sheet_name: str | None = None
) -> WorkbookCreatedResult:
    """Create a new .xlsx workbook.

    Optionally specify initial sheet names via sheet_names (list) or sheet_name (single).
    """
    return _workbook.create_workbook(file_path, sheet_names, sheet_name)


def get_sheet_summary(file_path: str, sheet_name: str) -> SheetSummary:
    """Get sheet summary: name, row/col counts, headers, and used range."""
    return _workbook.get_sheet_summary(file_path, sheet_name)


def write_multi_sheet(file_path: str, sheets: list[SheetDefinition]) -> WriteMultiSheetResult:
    """Create a new workbook with multiple named sheets, headers, data, and column widths in one call."""
    return _workbook.write_multi_sheet(file_path, sheets)


def sheet_management(
    action: Literal["rename", "delete", "copy", "hide", "unhide", "tab_color", "move"],
    file_path: str,
    sheet_name: str,
    new_name: str | None = None,
    color: str | None = None,
    offset: int | None = None,
) -> str | dict:
    """Manage worksheets within a workbook.

    action="rename": Rename a sheet. Requires: new_name.
    action="delete": DESTRUCTIVE. Delete a sheet. Raises if only sheet.
    action="copy": Copy a sheet. Requires: new_name for the copy.
    action="hide": Hide a sheet. Raises if it's the last visible sheet.
    action="unhide": Unhide a hidden sheet.
    action="tab_color": Set the tab colour. Requires: color (6-char hex, e.g. "FF0000"). Pass "000000" to clear.
    action="move": Reorder the sheet tab. Requires: offset (positive=right, negative=left).
    """
    if action == "rename":
        if not new_name:
            raise ValueError("new_name is required for action='rename'.")
        return _workbook.rename_sheet(file_path, sheet_name, new_name)
    if action == "delete":
        return _workbook.delete_sheet(file_path, sheet_name)
    if action == "copy":
        if not new_name:
            raise ValueError("new_name is required for action='copy'.")
        return _workbook.copy_sheet(file_path, sheet_name, new_name)
    if action == "hide":
        return _workbook.hide_sheet(file_path, sheet_name)
    if action == "unhide":
        return _workbook.unhide_sheet(file_path, sheet_name)
    if action == "tab_color":
        if not color:
            raise ValueError("color is required for action='tab_color'.")
        return _workbook.set_tab_color(file_path, sheet_name, color)
    if action == "move":
        if offset is None:
            raise ValueError("offset is required for action='move'.")
        return _workbook.move_sheet(file_path, sheet_name, offset)
    raise ValueError(f"Unknown action: {action}")


def register(mcp) -> None:
    """Register tools on *mcp*."""
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(get_workbook_metadata)
    mcp.tool()(create_workbook)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(get_sheet_summary)
    mcp.tool()(write_multi_sheet)
    mcp.tool()(sheet_management)
