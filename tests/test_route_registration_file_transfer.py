"""Route registration coverage for file transfer tools."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from mcp_server.routes import register_all_routes


class _FakeMCP:
    """Tiny MCP stub that captures registered tool names."""

    def __init__(self) -> None:
        self.tool_names: list[str] = []

    def tool(self, *args: Any, **kwargs: Any) -> Callable[..., Any]:
        def decorator(fn: Callable[..., Any]) -> Callable[..., Any]:
            self.tool_names.append(fn.__name__)
            return fn

        return decorator


def test_register_all_routes_includes_file_transfer_tools() -> None:
    """`register_all_routes` should register upload/download/release tools."""

    fake = _FakeMCP()
    register_all_routes(fake)

    assert "upload_file" in fake.tool_names
    assert "download_file" in fake.tool_names
    assert "release_file" in fake.tool_names
