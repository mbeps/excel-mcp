"""Image insertion operations for Excel workbooks."""

from __future__ import annotations

from logging import Logger

from openpyxl.drawing.image import Image

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
    validate_file_path,
    validate_single_cell,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

VALID_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp"}


def insert_image(
    file_path: str,
    sheet: str,
    image_path: str,
    cell: str,
    width: int | None = None,
    height: int | None = None,
) -> dict:
    """Insert an image into an Excel worksheet at the specified cell.

    Args:
        file_path: Path to the Excel workbook (.xlsx or .xlsm).
        sheet: Target worksheet name.
        image_path: Path to the image file to insert.
        cell: Cell reference where the image will be anchored (e.g. 'B2').
        width: Optional image width in pixels.
        height: Optional image height in pixels.

    Returns:
        Dict with status, message, and image_path.
    """
    validate_file_path(file_path)
    validate_single_cell(cell)
    validated_image = validate_file_path(image_path, allowed_extensions=VALID_IMAGE_EXTENSIONS)
    if width is not None and width <= 0:
        raise ValueError(f"width must be positive, got {width}")
    if height is not None and height <= 0:
        raise ValueError(f"height must be positive, got {height}")

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet)
        try:
            img = Image(str(validated_image))
        except Exception as exc:
            raise ValueError(f"Failed to load image '{image_path}': {exc}") from exc

        if width is not None and height is not None:
            img.width = width
            img.height = height
        elif width is not None:
            if img.width and img.height:
                img.height = max(1, round(img.height * (width / img.width)))
            img.width = width
        elif height is not None:
            if img.width and img.height:
                img.width = max(1, round(img.width * (height / img.height)))
            img.height = height

        ws.add_image(img, cell)
        save_workbook_safe(wb, file_path)
        logger.info("Inserted image '%s' at %s in sheet '%s'", image_path, cell, sheet)
    finally:
        wb.close()

    return {
        "status": "success",
        "message": f"Image inserted at {cell}",
        "image_path": image_path,
    }
