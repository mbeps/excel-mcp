from __future__ import annotations

from typing import Any, Literal

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
    """Return workbook metadata including sheet names, active sheet, dimensions, and named ranges.

    Args:
        file_path: Path to the workbook to inspect.

    Returns:
        WorkbookMetadata: Pydantic model containing sheet list, active sheet, named ranges, and other metadata.

    Notes:
        - Read-only operation. Underlying implementation may use a lightweight reader for speed.
    """
    return _workbook.get_workbook_metadata(file_path)


def create_workbook(
    file_path: str, sheet_names: list[str] | None = None, sheet_name: str | None = None
) -> WorkbookCreatedResult:
    """Create a new Excel workbook at `file_path` with optional initial sheets.

    Args:
        file_path: Destination path for the new workbook.
        sheet_names: Optional list of sheet names to create.
        sheet_name: Optional single sheet name (legacy convenience).

    Returns:
        WorkbookCreatedResult: Contains file path and sheet information.

    Notes:
        - Mutates filesystem by creating a new .xlsx. Parent directory will be created if permitted by utils.
    """
    return _workbook.create_workbook(file_path, sheet_names, sheet_name)


def get_sheet_summary(file_path: str, sheet_name: str) -> SheetSummary:
    """Return a brief summary of a sheet: header row, used range, row/column counts and detected headers.

    Args:
        file_path: Workbook path.
        sheet_name: Worksheet to summarise.

    Returns:
        SheetSummary: Pydantic model with summary fields.

    Notes:
        - Read-only.
    """
    return _workbook.get_sheet_summary(file_path, sheet_name)


def write_multi_sheet(file_path: str, sheets: list[SheetDefinition]) -> WriteMultiSheetResult:
    """Create or overwrite a workbook with multiple sheets, headers and data in a single call.

    Args:
        file_path: Destination workbook path.
        sheets: List of SheetDefinition (name, headers, rows, column widths, etc.).

    Returns:
        WriteMultiSheetResult: Result model with file path and any warnings.

    Notes:
        - Destructive when targeting existing files — document overwrite semantics in higher-level docs.
    """
    return _workbook.write_multi_sheet(file_path, sheets)


def sheet_management(
    action: Literal["rename", "delete", "copy", "hide", "unhide", "tab_color", "move"],
    file_path: str,
    sheet_name: str,
    new_name: str | None = None,
    color: str | None = None,
    offset: int | None = None,
) -> str | dict:
    """Manage sheets within a workbook (rename, delete, copy, hide/unhide, set tab color, move order).

    Args:
        action: One of "rename", "delete", "copy", "hide", "unhide", "tab_color", "move".
            - "rename": requires `new_name`.
            - "delete": deletes the sheet; destructive.
            - "copy": requires `new_name` for the copy.
            - "hide": hides the sheet (cannot hide all visible sheets).
            - "unhide": unhides the sheet.
            - "tab_color": requires `color` (6-char hex) to set or "000000" to clear.
            - "move": requires `offset` (int) to shift position.
        file_path: Workbook path.
        sheet_name: Target sheet name for the action.
        new_name: New name for rename/copy.
        color: Tab color hex string for "tab_color".
        offset: Position offset for "move" (positive = right).

    Returns:
        str or dict: Operation result or metadata.

    Raises:
        ValueError: When required arguments for an action are missing.

    Notes:
        - Dispatch mapping: "rename"→`tools.workbook.rename_sheet`, "delete"→`tools.workbook.delete_sheet`,
          "copy"→`tools.workbook.copy_sheet`, "hide"→`tools.workbook.hide_sheet`, etc.
        - Deletions and moves are destructive operations and should be annotated in external docs and UIs.
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


def register(mcp: Any) -> None:
    """Register tools on *mcp*."""
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(get_workbook_metadata)
    mcp.tool()(create_workbook)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(get_sheet_summary)
    mcp.tool()(write_multi_sheet)
    mcp.tool()(sheet_management)
