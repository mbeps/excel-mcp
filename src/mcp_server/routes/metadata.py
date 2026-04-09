from __future__ import annotations

from typing import Literal

import mcp_server.tools.comments as _comments
import mcp_server.tools.hyperlinks as _hyperlinks
import mcp_server.tools.scenarios as _scenarios
import mcp_server.tools.named_ranges as _named_ranges
import mcp_server.tools.tables as _tables
from mcp_server.models.comments import CommentInfo
from mcp_server.models.hyperlinks import HyperlinkInfo, HyperlinkReadResult
from mcp_server.models.scenarios import ScenarioInfo, ScenarioApplyResult
from mcp_server.models.named_ranges import NamedRangeInfo
from mcp_server.models.tables import TableInfo
from mcp_server.models.common import ScenarioCellValue

__all__ = [
    "comment",
    "hyperlink",
    "scenario",
    "named_range",
    "table",
]


def comment(
    action: Literal["add", "read", "delete", "list"],
    file_path: str,
    sheet_name: str,
    cell_ref: str | None = None,
    text: str | None = None,
    author: str = "Excel MCP",
) -> str | CommentInfo | list[CommentInfo] | None:
    """Comment operations on cells.

    action="add": Add a comment. Requires: cell_ref, text. Optional: author.
    action="read": Read a comment. Read-only. Requires: cell_ref. Returns None if no comment.
    action="delete": DESTRUCTIVE. Delete a comment. Requires: cell_ref.
    action="list": List all comments in sheet. Read-only.
    """
    if action == "add":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='add'.")
        if text is None:
            raise ValueError("text is required for action='add'.")
        return _comments.add_comment(file_path, sheet_name, cell_ref, text, author)
    if action == "read":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='read'.")
        return _comments.read_comment(file_path, sheet_name, cell_ref)
    if action == "delete":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='delete'.")
        return _comments.delete_comment(file_path, sheet_name, cell_ref)
    if action == "list":
        return _comments.list_comments(file_path, sheet_name)
    raise ValueError(f"Unknown action: {action}")


def hyperlink(
    action: Literal["add", "read", "delete", "list"],
    file_path: str,
    sheet_name: str,
    cell_ref: str | None = None,
    url: str | None = None,
    display_text: str | None = None,
    tooltip: str | None = None,
) -> str | HyperlinkReadResult | list[HyperlinkInfo] | None:
    """Hyperlink operations.

    action="add": Add hyperlink. Requires: cell_ref, url. Optional: display_text, tooltip.
    action="read": Read hyperlink. Read-only. Requires: cell_ref. Returns None if none.
    action="delete": DESTRUCTIVE. Delete hyperlink. Requires: cell_ref.
    action="list": List all hyperlinks. Read-only.
    """
    if action == "add":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='add'.")
        if url is None:
            raise ValueError("url is required for action='add'.")
        return _hyperlinks.add_hyperlink(file_path, sheet_name, cell_ref, url, display_text, tooltip)
    if action == "read":
        if not cell_ref:
            raise ValueError("cell_ref is required for action='read'.")
        return _hyperlinks.read_hyperlink(file_path, sheet_name, cell_ref)
    if action == "delete":
        if cell_ref is None:
            raise ValueError("cell_ref is required for action='delete'.")
        return _hyperlinks.delete_hyperlink(file_path, sheet_name, cell_ref)
    if action == "list":
        return _hyperlinks.list_hyperlinks(file_path, sheet_name)
    raise ValueError(f"Unknown action: {action}")


def scenario(
    action: Literal["add", "list", "apply"],
    file_path: str,
    name: str | None = None,
    cell_values: dict[str, dict[str, ScenarioCellValue]] | None = None,
    description: str = "",
) -> str | list[ScenarioInfo] | ScenarioApplyResult:
    """Scenario management for what-if analysis.

    action="add": Save a scenario. Requires: name, cell_values ({sheet: {cell: value}}). Optional: description.
    action="list": List all scenarios. Read-only.
    action="apply": DESTRUCTIVE. Apply scenario values to sheet. Requires: name.
    """
    if action == "add":
        if not name:
            raise ValueError("name is required for action='add'.")
        if cell_values is None:
            raise ValueError("cell_values is required for action='add'.")
        return _scenarios.add_scenario(file_path, name, cell_values, description)
    if action == "list":
        return _scenarios.list_scenarios(file_path)
    if action == "apply":
        if not name:
            raise ValueError("name is required for action='apply'.")
        return _scenarios.apply_scenario(file_path, name)
    raise ValueError(f"Unknown action: {action}")


def named_range(
    action: Literal["list", "create", "delete", "update"],
    file_path: str,
    name: str | None = None,
    destination: str | None = None,
    scope: str = "workbook",
    new_destination: str | None = None,
) -> list[NamedRangeInfo] | str:
    """Manage named ranges.

    action="list": List all named ranges. Read-only. Requires: file_path only.
    action="create": Create a named range. Requires: name, destination. Optional: scope.
    action="delete": DESTRUCTIVE. Delete a named range. Requires: name.
    action="update": Update destination. Requires: name, new_destination.
    """
    if action == "list":
        return _named_ranges.list_named_ranges(file_path)
    if action == "create":
        if not name:
            raise ValueError("name is required for action='create'.")
        if not destination:
            raise ValueError("destination is required for action='create'.")
        return _named_ranges.create_named_range(file_path, name, destination, scope)
    if action == "delete":
        if not name:
            raise ValueError("name is required for action='delete'.")
        return _named_ranges.delete_named_range(file_path, name)
    if action == "update":
        if not name:
            raise ValueError("name is required for action='update'.")
        if not new_destination:
            raise ValueError("new_destination is required for action='update'.")
        return _named_ranges.update_named_range(file_path, name, new_destination)
    raise ValueError(f"Unknown action: {action}")


def table(
    action: Literal["create", "list", "resize", "totals", "data", "convert_to_range"],
    file_path: str,
    sheet_name: str,
    table_name: str | None = None,
    data_range: str | None = None,
    style_name: str = "TableStyleMedium9",
    new_range: str | None = None,
    show_totals: bool | None = None,
    column_totals: dict[str, str] | None = None,
) -> str | list[TableInfo] | dict:
    """Excel table (ListObject) operations.

    action="create": Create a table. Requires: data_range, table_name. Optional: style_name.
    action="list": List all tables. Read-only.
    action="resize": Resize a table. Requires: table_name, new_range.
    action="totals": Toggle totals row. Requires: table_name, show_totals. Optional: column_totals.
    action="data": Read table data. Read-only. Requires: table_name.
    """
    if action == "create":
        if not data_range:
            raise ValueError("data_range is required for action='create'.")
        if not table_name:
            raise ValueError("table_name is required for action='create'.")
        return _tables.create_table(file_path, sheet_name, data_range, table_name, style_name)
    if action == "list":
        return _tables.list_tables(file_path, sheet_name)
    if action == "resize":
        if table_name is None:
            raise ValueError("table_name is required for action='resize'.")
        if new_range is None:
            raise ValueError("new_range is required for action='resize'.")
        return _tables.resize_table(file_path, sheet_name, table_name, new_range)
    if action == "totals":
        if table_name is None:
            raise ValueError("table_name is required for action='totals'.")
        if show_totals is None:
            raise ValueError("show_totals is required for action='totals'.")
        return _tables.set_table_totals_row(file_path, sheet_name, table_name, show_totals, column_totals)
    if action == "data":
        if table_name is None:
            raise ValueError("table_name is required for action='data'.")
        return _tables.get_table_data(file_path, sheet_name, table_name)
    if action == "convert_to_range":
        if table_name is None:
            raise ValueError("table_name is required for action='convert_to_range'.")
        return _tables.convert_table_to_range(file_path, sheet_name, table_name)
    raise ValueError(f"Unknown action: {action}")


def register(mcp) -> None:
    """Register tools on *mcp*."""
    mcp.tool()(comment)
    mcp.tool()(hyperlink)
    mcp.tool()(scenario)
    mcp.tool()(named_range)
    mcp.tool()(table)
