"""Domain-specific tool registration modules.

Each module exports a ``register(mcp)`` function that binds
``@mcp.tool`` handlers for its domain.
"""

from __future__ import annotations

import functools
import inspect
from typing import Any, Callable, cast

try:
    from mcp.server.mcpserver.exceptions import ToolError
except ImportError:  # pragma: no cover

    class ToolError(Exception):  # type: ignore[no-redef]
        """Fallback ToolError when mcp is not installed."""

        pass


def _wrap_tool_fn(fn: Callable[..., Any]) -> Callable[..., Any]:
    """Wrap tool handler so domain exceptions surface as ToolError in MCP v2."""
    if inspect.iscoroutinefunction(fn):

        @functools.wraps(fn)
        async def async_wrapped(*args: Any, **kwargs: Any) -> Any:
            try:
                return await fn(*args, **kwargs)
            except (ValueError, FileNotFoundError, PermissionError, KeyError) as e:
                raise ToolError(str(e)) from e

        return async_wrapped

    @functools.wraps(fn)
    def sync_wrapped(*args: Any, **kwargs: Any) -> Any:
        try:
            return fn(*args, **kwargs)
        except (ValueError, FileNotFoundError, PermissionError, KeyError) as e:
            raise ToolError(str(e)) from e

    return sync_wrapped


class _MCPToolWrapper:
    """Delegates attribute access to mcp while wrapping tool registrations with error preservation."""

    def __init__(self, mcp: Any) -> None:
        self._mcp = mcp

    def __getattr__(self, name: str) -> Any:
        return getattr(self._mcp, name)

    def tool(self, *args: Any, **kwargs: Any) -> Callable[..., Any]:
        decorator = self._mcp.tool(*args, **kwargs)

        def wrap(fn: Callable[..., Any]) -> Callable[..., Any]:
            return cast(Callable[..., Any], decorator(_wrap_tool_fn(fn)))

        return wrap


def register_all_routes(mcp: object) -> None:
    """Register every domain's tools on *mcp*."""
    from mcp_server.routes import (
        analysis,
        cell_ops,
        charts,
        cleaning,
        custom_code,
        file_transfer,
        financial,
        formatting,
        formulas,
        governance,
        metadata,
        multi_file,
        pivot_etl,
        statistical,
        workbook,
        worksheet_ops,
    )

    wrapped_mcp = _MCPToolWrapper(mcp) if hasattr(mcp, "tool") else mcp

    for mod in (
        workbook,
        cell_ops,
        formatting,
        formulas,
        charts,
        worksheet_ops,
        analysis,
        pivot_etl,
        financial,
        cleaning,
        statistical,
        governance,
        metadata,
        multi_file,
        custom_code,
        file_transfer,
    ):
        mod.register(wrapped_mcp)
