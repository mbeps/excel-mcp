"""Helpers for creating a stderr-only logger suitable for stdio MCP servers.

This module provides a convenience function to produce a configured
`logging.Logger` instance that writes only to stderr so stdout remains
available for JSON-RPC or other protocol traffic.
"""

from __future__ import annotations

import logging
import sys


def configure_logging(name: str = "mcp_server") -> logging.Logger:
    """Create a stderr-only logger suitable for stdio MCP servers.

    Args:
        name: Logger namespace used for filtering and formatting.

    Returns:
        Configured logger with a single stream handler.

    References:
        - Python logging guide: https://docs.python.org/3/library/logging.html
    """
    logger: logging.Logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.propagate = False

    if not logger.handlers:
        handler: logging.StreamHandler = logging.StreamHandler(sys.stderr)
        formatter: logging.Formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    return logger
