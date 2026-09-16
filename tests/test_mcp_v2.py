"""MCP v2 stateless in-memory test suite for excel-mcp-server.

Tests tools, resources, prompts, error message preservation, and ASGI Starlette app
using the in-memory MCP Client without networking overhead.
"""

from __future__ import annotations

import json
import urllib.parse
from collections.abc import AsyncIterator
from pathlib import Path

import pytest
from mcp import Client
from starlette.applications import Starlette

from mcp_server.main import app, mcp


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"


@pytest.fixture
async def client() -> AsyncIterator[Client]:
    """In-memory client connected directly to the MCPServer instance."""
    async with Client(mcp, raise_exceptions=False) as c:
        yield c


@pytest.mark.anyio
async def test_mcp_v2_tools_discovery(client: Client) -> None:
    """Verify all 69 tools are registered with schemas and descriptions."""
    res = await client.list_tools()
    tools = res.tools

    assert len(tools) == 69, f"Expected 69 tools, got {len(tools)}"

    tool_map = {t.name: t for t in tools}

    # Core representative tools
    assert "create_workbook" in tool_map
    assert "get_workbook_metadata" in tool_map
    assert "read_cells" in tool_map
    assert "write_cells" in tool_map
    assert "clear_range" in tool_map
    assert "sheet_management" in tool_map
    assert "upload_file" in tool_map
    assert "download_file" in tool_map
    assert "release_file" in tool_map

    # Validate schema shape and descriptions
    for tool in tools:
        assert tool.description, f"Tool {tool.name} must have a description"
        assert tool.input_schema is not None, f"Tool {tool.name} must have input_schema"
        assert tool.input_schema.get("type") == "object", f"Tool {tool.name} input_schema must be object"

    # Verify snake_case annotations
    meta_tool = tool_map["get_workbook_metadata"]
    assert meta_tool.annotations is not None
    assert meta_tool.annotations.read_only_hint is True

    clear_tool = tool_map["clear_range"]
    assert clear_tool.annotations is not None
    assert clear_tool.annotations.destructive_hint is True

    upload_tool = tool_map["upload_file"]
    assert upload_tool.annotations is not None
    assert upload_tool.annotations.read_only_hint is False
    assert upload_tool.annotations.destructive_hint is False


@pytest.mark.anyio
async def test_mcp_v2_tool_call_success(client: Client, tmp_path: Path) -> None:
    """Verify invoking tools through the MCP v2 client succeeds."""
    file_path = str(tmp_path / "v2_test.xlsx")

    # 1. Create workbook
    create_res = await client.call_tool("create_workbook", {"file_path": file_path, "sheet_name": "Sales"})
    assert not create_res.is_error
    assert len(create_res.content) > 0

    # 2. Get sheet summary
    summary_res = await client.call_tool("get_sheet_summary", {"file_path": file_path, "sheet_name": "Sales"})
    assert not summary_res.is_error
    assert len(summary_res.content) > 0

    # 3. Write cells
    write_res = await client.call_tool(
        "write_cells",
        {
            "mode": "range",
            "file_path": file_path,
            "sheet_name": "Sales",
            "start_cell": "A1",
            "data": [["Item", "Price"], ["Widget", 10.5]],
        },
    )
    assert not write_res.is_error

    # 4. Read cells
    read_res = await client.call_tool(
        "read_cells",
        {
            "mode": "range",
            "file_path": file_path,
            "sheet_name": "Sales",
            "start_cell": "A1",
            "end_cell": "B2",
        },
    )
    assert not read_res.is_error
    read_text = read_res.content[0].text
    assert "Widget" in read_text


@pytest.mark.anyio
async def test_mcp_v2_error_preservation(client: Client, tmp_path: Path) -> None:
    """Verify domain exceptions preserve model-visible messages rather than being masked."""
    file_path = str(tmp_path / "err_test.xlsx")

    # Calling sheet_management with an invalid action raises ValueError in the route
    res = await client.call_tool(
        "sheet_management",
        {
            "action": "nonexistent_action",
            "file_path": file_path,
            "sheet_name": "Sheet1",
        },
    )

    assert res.is_error is True
    assert len(res.content) > 0
    error_message = res.content[0].text

    # The error message MUST contain the actual failure reason, NOT generic "Error executing tool"
    assert "Unknown action" in error_message or "nonexistent_action" in error_message


@pytest.mark.anyio
async def test_mcp_v2_resource_templates_and_read(client: Client, tmp_path: Path) -> None:
    """Verify resource templates discovery and reading."""
    templates_res = await client.list_resource_templates()
    templates = [t.uri_template for t in templates_res.resource_templates]

    assert "excel://workbook/{file_path}/sheets" in templates
    assert "excel://workbook/{file_path}/sheet/{sheet_name}/preview" in templates

    # Create a real workbook to read via resource
    file_path = str(tmp_path / "resource_test.xlsx")
    await client.call_tool("create_workbook", {"file_path": file_path, "sheet_name": "Data"})

    # Read workbook sheets resource (URL-encoded path)
    encoded_path = urllib.parse.quote(file_path, safe="")
    res_content = await client.read_resource(f"excel://workbook/{encoded_path}/sheets")
    assert len(res_content.contents) > 0

    data = json.loads(res_content.contents[0].text)
    assert isinstance(data, list)
    assert any(s["name"] == "Data" for s in data)


@pytest.mark.anyio
async def test_mcp_v2_prompts_discovery_and_get(client: Client) -> None:
    """Verify all 20 prompts are listed and can be retrieved."""
    prompts_res = await client.list_prompts()
    prompts = prompts_res.prompts

    assert len(prompts) == 20, f"Expected 20 prompts, got {len(prompts)}"
    prompt_names = {p.name for p in prompts}

    assert "excel-quickstart" in prompt_names
    assert "excel-data-analysis" in prompt_names
    assert "excel-financial-model" in prompt_names

    # Retrieve a specific prompt
    prompt = await client.get_prompt("excel-quickstart", arguments={"file_path": "report.xlsx", "sheet_name": "Summary"})
    assert prompt.messages
    message_text = prompt.messages[0].content.text
    assert "report.xlsx" in message_text
    assert "Summary" in message_text


def test_mcp_v2_asgi_app() -> None:
    """Ensure stateless Streamable HTTP Starlette app is properly configured."""
    assert isinstance(app, Starlette)
    fresh_app = mcp.streamable_http_app(stateless_http=True)
    assert isinstance(fresh_app, Starlette)

