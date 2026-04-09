from __future__ import annotations

import json
from logging import Logger

from mcp_server.utils.logger import configure_logging


class _NoopMCP:
    """Stub used when tool registration is disabled (e.g. during tests)."""

    def resource(self, *args, **kwargs):
        def decorator(fn):
            return fn

        return decorator

    def tool(self, *args, **kwargs):
        def decorator(fn):
            return fn

        return decorator

    def prompt(self, *args, **kwargs):
        def decorator(fn):
            return fn

        return decorator


if __import__("os").environ.get("MCP_SERVER_DISABLE_TOOL_REGISTRATION") == "1":
    mcp = _NoopMCP()
else:
    from mcp.server.fastmcp import FastMCP

    mcp = FastMCP("excel-mcp-server")

logger: Logger = configure_logging("mcp_server.main")

# ── Resources ────────────────────────────────────────────────────────────────

import mcp_server.tools.workbook as _workbook  # noqa: E402
import mcp_server.tools.cell_ops as _cell_ops  # noqa: E402


@mcp.resource("excel://workbook/{file_path}/sheets", mime_type="application/json")
def resource_list_sheets(file_path: str) -> str:
    """List all sheets in a workbook as JSON."""
    metadata = _workbook.get_workbook_metadata(file_path)
    return json.dumps([s.model_dump() for s in metadata.sheets], indent=2)


@mcp.resource("excel://workbook/{file_path}/sheet/{sheet_name}/preview", mime_type="application/json")
def resource_sheet_preview(file_path: str, sheet_name: str) -> str:
    """Preview the first 20 rows of a sheet as JSON."""
    data = _cell_ops.read_range(file_path, sheet_name, "A1", "Z20")
    return json.dumps(data, indent=2, default=str)


# ── Tool Registration ────────────────────────────────────────────────────────

from mcp_server.routes import register_all_routes  # noqa: E402

register_all_routes(mcp)

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

# ── Prompt Registration ──────────────────────────────────────────────────────

from mcp_server.prompts import register_prompts as _register_prompts  # noqa: E402

_register_prompts(mcp)

# ── Entry Point ──────────────────────────────────────────────────────────────


def run() -> None:
    """Launch the MCP server."""
    try:
        mcp.run()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception:
        logger.critical("Fatal server error", exc_info=True)
        raise


if __name__ == "__main__":
    run()
