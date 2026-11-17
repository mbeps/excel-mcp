"""excel-mcp MCP server entrypoint and route registration.

This module configures logging and initialises the MCP `mcp` object (an
`MCPServer` instance) unless tool registration is disabled for testing. It
registers JSON resources and all tool routes, exposes a backwards-compatible
set of route exports for tests/consumers, and provides `run()` as the
convenience entrypoint to start the server.

Usage:
    Call `run()` to start the MCP server. Tests may set
    `MCP_SERVER_DISABLE_TOOL_REGISTRATION=1` to avoid real registration.
"""

from __future__ import annotations

import json
import os
import urllib.parse
from collections.abc import Callable
from logging import Logger
from typing import Any

from mcp_server.utils.logger import configure_logging


class _NoopMCP:
    """Stub used when tool registration is disabled (e.g. during tests)."""

    def resource(self, *args: Any, **kwargs: Any) -> Callable[..., Any]:
        def decorator(fn: Callable[..., Any]) -> Callable[..., Any]:
            return fn

        return decorator

    def tool(self, *args: Any, **kwargs: Any) -> Callable[..., Any]:
        def decorator(fn: Callable[..., Any]) -> Callable[..., Any]:
            return fn

        return decorator

    def prompt(self, *args: Any, **kwargs: Any) -> Callable[..., Any]:
        def decorator(fn: Callable[..., Any]) -> Callable[..., Any]:
            return fn

        return decorator

    def run(self, *args: Any, **kwargs: Any) -> None:
        pass

    def streamable_http_app(self, *args: Any, **kwargs: Any) -> Any:
        return None


logger: Logger = configure_logging("mcp_server.main")

# ── Resources ────────────────────────────────────────────────────────────────

import mcp_server.tools.workbook as _workbook  # noqa: E402
import mcp_server.tools.cell_ops as _cell_ops  # noqa: E402
from mcp_server.prompts import register_prompts as _register_prompts  # noqa: E402
from mcp_server.routes import register_all_routes  # noqa: E402


def resource_list_sheets(file_path: str) -> str:
    """List all sheets in a workbook as JSON."""
    decoded_path = urllib.parse.unquote(file_path)
    metadata = _workbook.get_workbook_metadata(decoded_path)
    return json.dumps([s.model_dump() for s in metadata.sheets], indent=2)


def resource_sheet_preview(file_path: str, sheet_name: str) -> str:
    """Preview the first 20 rows of a sheet as JSON."""
    decoded_path = urllib.parse.unquote(file_path)
    decoded_sheet = urllib.parse.unquote(sheet_name)
    data = _cell_ops.read_range(decoded_path, decoded_sheet, "A1", "Z20")
    return json.dumps(data, indent=2, default=str)


def create_server(disable_registration: bool | None = None) -> Any:
    """Create and configure an MCPServer (or _NoopMCP when registration is disabled)."""
    is_disabled = (
        os.environ.get("MCP_SERVER_DISABLE_TOOL_REGISTRATION") == "1"
        if disable_registration is None
        else disable_registration
    )
    if is_disabled:
        server: Any = _NoopMCP()
    else:
        from mcp.server.mcpserver import MCPServer, ResourceSecurity

        server = MCPServer(
            name="excel-mcp-server",
            version="0.1.0",
            description="MCP server for Excel automation and data manipulation",
            resource_security=ResourceSecurity(exempt_params={"file_path"}),
        )

    server.resource("excel://workbook/{file_path}/sheets", mime_type="application/json")(resource_list_sheets)
    server.resource("excel://workbook/{file_path}/sheet/{sheet_name}/preview", mime_type="application/json")(
        resource_sheet_preview
    )
    register_all_routes(server)
    _register_prompts(server)
    return server


mcp: Any = create_server()
app: Any = mcp.streamable_http_app(stateless_http=True) if hasattr(mcp, "streamable_http_app") else None

# ── Backward-compatible exports ──────────────────────────────────────────────
# Tests and external code may import tool functions from main.py directly.

from mcp_server.routes.workbook import *  # noqa: E402, F401, F403
from mcp_server.routes.cell_ops import *  # noqa: E402, F401, F403
from mcp_server.routes.formatting import *  # noqa: E402, F401, F403
from mcp_server.routes.formulas import *  # noqa: E402, F401, F403
from mcp_server.routes.charts import *  # noqa: E402, F401, F403
from mcp_server.routes.worksheet_ops import *  # noqa: E402, F401, F403
from mcp_server.routes.analysis import *  # noqa: E402, F401, F403
from mcp_server.routes.pivot_etl import *  # noqa: E402, F401, F403
from mcp_server.routes.financial import *  # noqa: E402, F401, F403
from mcp_server.routes.cleaning import *  # noqa: E402, F401, F403
from mcp_server.routes.statistical import *  # noqa: E402, F401, F403
from mcp_server.routes.governance import *  # noqa: E402, F401, F403
from mcp_server.routes.metadata import *  # noqa: E402, F401, F403
from mcp_server.routes.multi_file import *  # noqa: E402, F401, F403
from mcp_server.routes.custom_code import *  # noqa: E402, F401, F403
from mcp_server.routes.file_transfer import *  # noqa: E402, F401, F403

# ── Internal module aliases (for test monkeypatching) ────────────────────────
import mcp_server.tools.formatting as _formatting  # noqa: E402
import mcp_server.tools.formulas as _formulas  # noqa: E402
import mcp_server.tools.csv_ops as _csv_ops  # noqa: E402
import mcp_server.tools.conditional_formatting as _cond_fmt  # noqa: E402
import mcp_server.tools.tables as _tables  # noqa: E402
import mcp_server.tools.data_validation as _data_val  # noqa: E402
import mcp_server.tools.protection as _protection  # noqa: E402
import mcp_server.tools.doc_properties as _doc_props  # noqa: E402
import mcp_server.tools.charts as _charts  # noqa: E402
import mcp_server.tools.named_ranges as _named_ranges  # noqa: E402
import mcp_server.tools.comments as _comments  # noqa: E402
import mcp_server.tools.hyperlinks as _hyperlinks  # noqa: E402
import mcp_server.tools.scenarios as _scenarios  # noqa: E402
import mcp_server.tools.multi_file as _multi_file  # noqa: E402
import mcp_server.tools.worksheet_ops as _ws_ops  # noqa: E402
import mcp_server.tools.analysis as _analysis  # noqa: E402
import mcp_server.tools.pivot_etl as _pivot_etl  # noqa: E402
import mcp_server.tools.financial as _financial  # noqa: E402
import mcp_server.tools.statistical as _statistical  # noqa: E402
import mcp_server.tools.solver as _solver  # noqa: E402
import mcp_server.tools.cleaning as _cleaning  # noqa: E402

# ── Entry Point ──────────────────────────────────────────────────────────────


def run(argv: list[str] | None = None) -> None:
    """Launch the MCP server.

    Supports two transport modes selected via --transport:
      stdio (default): reads/writes on stdin/stdout. Suitable for local clients.
      http: starts a streamable-HTTP server. Use --host and --port to configure.
    """
    import argparse

    parser = argparse.ArgumentParser(description="Excel MCP Server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "http", "streamable-http"],
        default="stdio",
    )
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    args = parser.parse_args(argv)

    if args.transport in ("http", "streamable-http"):
        logger.info("Starting MCP server in streamable-http mode on %s:%s", args.host, args.port)
    else:
        logger.info("Starting MCP server in stdio mode")

    try:
        if args.transport in ("http", "streamable-http"):
            mcp.run(
                transport="streamable-http",
                host=args.host,
                port=args.port,
                stateless_http=True,
            )
        else:
            mcp.run()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception:
        logger.critical("Fatal server error", exc_info=True)
        raise


if __name__ == "__main__":
    run()
