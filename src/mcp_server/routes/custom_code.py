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
    """Insert an image into a worksheet anchored at a target cell.

    Args:
        file_path: Workbook path to modify.
        sheet: Worksheet name.
        image_path: Path to image (PNG/JPG/GIF).
        cell: Anchor cell where image is placed.
        width: Optional width in Excel units.
        height: Optional height in Excel units.

    Returns:
        dict: Metadata about the inserted image (anchor, size, file used).

    Notes:
        - Mutates workbook and depends on Pillow. Document supported image formats and sizing behaviour.
    """
    return _images.insert_image(file_path, sheet, image_path, cell, width, height)


def execute_custom_code(
    file_path: str,
    code: str,
    sheet: str | None = None,
    output_file: str | None = None,
) -> dict:
    """Execute sandboxed Python/pandas code against a workbook or sheet and return `result`.

    Args:
        file_path: Path to source workbook to load into the sandbox.
        code: Python code string. Sandbox exposes `df` (pandas.DataFrame), `pd` (pandas), `np` (numpy).
        sheet: Optional sheet name to load into `df`. If omitted, the first sheet or a default is used.
        output_file: Optional path to write results back to a workbook.

    Returns:
        dict: Execution result, typically containing `result` (from user code), `stdout` and `errors`.

    Raises:
        ValueError: If code fails safety checks in the sandbox.

    Notes:
        - High-risk: sandbox uses AST checks — document the allowed AST nodes and forbidden names.
        - Recommend returning a short example snippet of a safe operation in the route docs.
    """
    return _custom_code.execute_custom_code(file_path, code, sheet, output_file)


def register(mcp) -> None:
    """Register tools on *mcp*."""
    mcp.tool()(insert_image)
    mcp.tool()(execute_custom_code)
