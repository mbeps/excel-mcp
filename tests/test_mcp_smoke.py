import pytest
from mcp_server.main import mcp

@pytest.mark.anyio
async def test_mcp_tools_registration():
    """Verify that MCP tools are correctly registered and accessible."""
    tools = await mcp.list_tools()
    assert len(tools) > 0
    tool_names = [t.name for t in tools]
    assert "create_pivot_table_native" in tool_names
    assert "write_cells" in tool_names

