from __future__ import annotations

import mcp_server.tools.custom_code as _custom_code
import mcp_server.tools.images as _images

__all__ = [
    "insert_image",
    "execute_custom_code",
]


def insert_image(
    file_path: str,
    sheet: str,
    image_path: str,
    cell: str,
    width: int | None = None,
    height: int | None = None,
) -> dict:
    """Insert an image into a worksheet at the specified cell."""
    return _images.insert_image(file_path, sheet, image_path, cell, width, height)


def execute_custom_code(
    file_path: str,
    code: str,
    sheet: str | None = None,
    output_file: str | None = None,
) -> dict:
    """Execute custom Python/pandas code against an Excel file in a sandboxed environment.

    Use this for operations not covered by other tools.
    The code has access to 'df' (the DataFrame), 'pd' (pandas), and 'np' (numpy).
    Set 'result' variable to return data.
    """
    return _custom_code.execute_custom_code(file_path, code, sheet, output_file)


def register(mcp) -> None:
    """Register tools on *mcp*."""
    mcp.tool()(insert_image)
    mcp.tool()(execute_custom_code)
