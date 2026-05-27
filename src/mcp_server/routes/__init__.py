"""Domain-specific tool registration modules.

Each module exports a ``register(mcp)`` function that binds
``@mcp.tool`` handlers for its domain.
"""

from __future__ import annotations


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
        mod.register(mcp)
