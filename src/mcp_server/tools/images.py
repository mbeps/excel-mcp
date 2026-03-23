"""Image operations: insert, list, and delete images in worksheets."""

from __future__ import annotations

from logging import Logger
from pathlib import Path

from openpyxl.drawing.image import Image

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)

ALLOWED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp"}


def insert_image(
    file_path: str,
    sheet_name: str,
    image_path: str,
    cell_ref: str,
    width: int | None = None,
    height: int | None = None,
) -> str:
    """Insert an image into a sheet at the specified cell."""
    img_path = Path(image_path).resolve()
    if not img_path.exists():
        raise ValueError(f"Image file not found: {image_path}")
    if img_path.suffix.lower() not in ALLOWED_IMAGE_EXTENSIONS:
        raise ValueError(f"Unsupported image type: {img_path.suffix}. Allowed: {ALLOWED_IMAGE_EXTENSIONS}")

    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        img = Image(str(img_path))
        if width is not None:
            img.width = width
        if height is not None:
            img.height = height

        ws.add_image(img, cell_ref)
        save_workbook_safe(wb, file_path)
        logger.info("Inserted image '%s' at %s in sheet '%s'", image_path, cell_ref, sheet_name)
        return f"Image inserted at {cell_ref} in sheet '{sheet_name}'."
    finally:
        wb.close()


def list_images(file_path: str, sheet_name: str) -> list[dict]:
    """List all images in a sheet."""
    wb = load_workbook_safe(file_path, read_only=False)
    try:
        ws = get_sheet(wb, sheet_name)
        results = []
        for i, img in enumerate(ws._images):
            anchor_str = ""
            if hasattr(img, "anchor") and img.anchor is not None:
                anchor_obj = img.anchor
                if hasattr(anchor_obj, "_from") and anchor_obj._from is not None:
                    col = anchor_obj._from.col
                    row = anchor_obj._from.row
                    anchor_str = f"col={col}, row={row}"
                else:
                    anchor_str = str(anchor_obj)
            results.append(
                {
                    "index": i,
                    "width": img.width,
                    "height": img.height,
                    "anchor": anchor_str,
                }
            )
        return results
    finally:
        wb.close()


def delete_image(file_path: str, sheet_name: str, image_index: int = 0) -> str:
    """Delete an image by index from the sheet."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)

        images = ws._images
        if image_index < 0 or image_index >= len(images):
            raise ValueError(f"Image index {image_index} out of range. Sheet has {len(images)} image(s).")

        del images[image_index]
        save_workbook_safe(wb, file_path)
        logger.info("Deleted image at index %d from sheet '%s'", image_index, sheet_name)
        return f"Image at index {image_index} deleted from sheet '{sheet_name}'."
    finally:
        wb.close()
