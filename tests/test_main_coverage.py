import pytest
from mcp_server.main import mcp
import json
import os


def test_main_resources(sample_xlsx):
    # Test resources in main.py
    from mcp_server.main import resource_list_sheets, resource_sheet_preview

    res1 = resource_list_sheets(sample_xlsx)
    assert "Sheet1" in res1

    res2 = resource_sheet_preview(sample_xlsx, "Sheet1")
    # Actually check for 'rows' or just printed content
    data = json.loads(res2)
    assert "rows" in data
    assert len(data["rows"]) > 0


def test_main_tools(sample_xlsx):
    # Test tool wrappers in main.py
    from mcp_server.main import (
        get_workbook_metadata,
        create_workbook,
        get_sheet_summary,
        write_multi_sheet,
        sheet_management,
    )

    # Simple tool call
    meta = get_workbook_metadata(sample_xlsx)
    assert meta["active_sheet"] == "Sheet1"

    # Sheet management consolidated tool
    res = sheet_management("rename", sample_xlsx, "Sheet1", "RenamedSheet")
    # Be case-insensitive and flexible with the exact message
    assert "renamed to 'renamedsheet'" in res.lower()

    res = sheet_management("copy", sample_xlsx, "RenamedSheet", "CopiedSheet")
    assert "copied" in res.lower() and "copiedsheet" in res.lower()

    res = sheet_management("delete", sample_xlsx, "CopiedSheet")
    assert "deleted" in res.lower()
