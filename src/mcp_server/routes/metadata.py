from __future__ import annotations

from typing import Any, Literal

import mcp_server.tools.comments as _comments
import mcp_server.tools.hyperlinks as _hyperlinks
import mcp_server.tools.named_ranges as _named_ranges
import mcp_server.tools.scenarios as _scenarios
import mcp_server.tools.tables as _tables
from mcp_server.models.comments import CommentInfo
from mcp_server.models.common import ScenarioCellValue
from mcp_server.models.hyperlinks import HyperlinkInfo, HyperlinkReadResult
from mcp_server.models.named_ranges import NamedRangeInfo
from mcp_server.models.scenarios import ScenarioApplyResult, ScenarioInfo
from mcp_server.models.tables import TableInfo

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
    """Add, read, delete or list cell comments on a sheet.

    Args:
        action: "add", "read", "delete", or "list".
            - "add": requires `cell_ref` and `text`; optional `author`.
            - "read": requires `cell_ref`; returns CommentInfo or None.
            - "delete": requires `cell_ref`; destructive.
            - "list": returns list of CommentInfo for the sheet.
        file_path: Workbook path.
        sheet_name: Worksheet name.
        cell_ref: Cell reference for single-cell operations.
        text: Comment text for "add".
        author: Optional author string.

    Returns:
        str | CommentInfo | list[CommentInfo] | None: Depends on action.

    Notes:
        - Deletions mutate the workbook.
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
    """Add, read, delete, or list hyperlinks attached to cells.

    Args:
        action: One of "add", "read", "delete", "list".
        file_path: Workbook path.
        sheet_name: Worksheet name.
        cell_ref: Cell reference (required for single-cell ops).
        url: URL for "add".
        display_text, tooltip: Optional display and tooltip text.

    Returns:
        str | HyperlinkReadResult | list[HyperlinkInfo] | None

    Notes:
        - Deleting hyperlinks modifies the workbook.
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
    """Manage saved scenarios (what-if value sets) persisted in a hidden sheet.

    Args:
        action: "add", "list", or "apply".
            - "add": requires `name` and `cell_values` — a nested dict mapping sheet name to {cell_ref: value},
              e.g. {"Sheet1": {"A1": 100, "B2": 200}, "Sheet2": {"C3": "hello"}}.
            - "list": returns available scenarios.
            - "apply": requires `name` and will write stored cell values into the workbook (destructive).
        file_path: Workbook path.
        name: Scenario name for add/apply.
        cell_values: Nested dict mapping sheet_name → {cell_ref: scalar_value} for "add".
        description: Optional free-text description.

    Returns:
        str | list[ScenarioInfo] | ScenarioApplyResult

    Notes:
        - Scenarios are stored in a hidden `_mcp_scenarios` sheet — mention potential
          user-visible side-effects when users open the workbook.
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
    """List, create, delete, or update named ranges within a workbook.

    Args:
        action: "list", "create", "delete", or "update".
        file_path: Workbook path.
        name: Named range name for create/delete/update.
        destination: Destination range string for create.
        scope: "workbook" or sheet-scoped identifier.
        new_destination: New range for update.

    Returns:
        list[NamedRangeInfo] | str

    Notes:
        - Creating/updating named ranges mutates workbook metadata but typically does not alter cell values.
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
    """Create, list, resize, toggle totals, read data, or convert tables to ranges.

    Args:
        action: "create", "list", "resize", "totals", "data", or "convert_to_range".
        file_path: Workbook path.
        sheet_name: Worksheet containing the table.
        table_name: Table name for operations that require it.
        data_range: Range to use when creating a table.
        style_name: Named table style for creation.
        new_range: New range for resize.
        show_totals: Bool for toggling totals row.
        column_totals: Dict mapping column name to aggregation function name
            (e.g. {"Revenue": "sum", "Quantity": "count"}). Valid functions:
            sum, count, average, max, min, countNums, stdDev, var, none.

    Returns:
        str | list[TableInfo] | dict

    Notes:
        - Table creation/resizing mutates workbook structure.
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


def register(mcp: Any) -> None:
    """Register tools on *mcp*."""
    mcp.tool()(comment)
    mcp.tool()(hyperlink)
    mcp.tool()(scenario)
    mcp.tool()(named_range)
    mcp.tool()(table)
