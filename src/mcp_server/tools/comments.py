"""Comment operations: add, read, delete, list."""

from __future__ import annotations

from logging import Logger

from openpyxl.comments import Comment

from mcp_server.utils.excel_helpers import (
    get_sheet,
    load_workbook_safe,
    save_workbook_safe,
)
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def add_comment(
    file_path: str,
    sheet_name: str,
    cell_ref: str,
    text: str,
    author: str = "Excel MCP",
) -> str:
    """Add a comment/note to a cell."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws[cell_ref].comment = Comment(text, author)
        save_workbook_safe(wb, file_path)
        logger.info("Added comment to %s!%s in %s", sheet_name, cell_ref, file_path)
        return f"Comment added to cell {cell_ref} on sheet '{sheet_name}'."
    finally:
        wb.close()


def update_comment(
    file_path: str,
    sheet_name: str,
    cell_ref: str,
    text: str,
    author: str | None = None,
) -> str:
    """Update an existing comment on a cell. Raises ValueError if no comment exists."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        existing = ws[cell_ref].comment
        if existing is None:
            raise ValueError(f"No comment found on cell {cell_ref} in sheet '{sheet_name}'.")
        new_author = author if author is not None else existing.author
        ws[cell_ref].comment = Comment(text, new_author)
        save_workbook_safe(wb, file_path)
        logger.info("Updated comment on %s!%s in %s", sheet_name, cell_ref, file_path)
        return f"Comment on cell {cell_ref} updated."
    finally:
        wb.close()


def read_comment(file_path: str, sheet_name: str, cell_ref: str) -> dict | None:
    """Read comment from a cell. Returns dict with text/author or None."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        comment = ws[cell_ref].comment
        if comment is None:
            return None
        return {"text": comment.text, "author": comment.author}
    finally:
        wb.close()


def delete_comment(file_path: str, sheet_name: str, cell_ref: str) -> str:
    """Delete comment from a cell."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        ws[cell_ref].comment = None
        save_workbook_safe(wb, file_path)
        logger.info("Deleted comment from %s!%s in %s", sheet_name, cell_ref, file_path)
        return f"Comment deleted from cell {cell_ref} on sheet '{sheet_name}'."
    finally:
        wb.close()


def list_comments(file_path: str, sheet_name: str) -> list[dict]:
    """List all comments in a sheet."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        results = []
        for row in ws.iter_rows(
            min_row=ws.min_row,
            max_row=ws.max_row,
            min_col=ws.min_column,
            max_col=ws.max_column,
        ):
            for cell in row:
                if cell.comment is not None:
                    results.append(
                        {
                            "cell_ref": cell.coordinate,
                            "text": cell.comment.text,
                            "author": cell.comment.author,
                        }
                    )
        return results
    finally:
        wb.close()


def add_comments_bulk(
    file_path: str,
    sheet_name: str,
    comments: list[dict],
) -> dict:
    """Add comments to multiple cells at once."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        count = 0
        for entry in comments:
            cell_ref = entry["cell"]
            text = entry["text"]
            author = entry.get("author", "Excel MCP")
            ws[cell_ref].comment = Comment(text, author)
            count += 1
        save_workbook_safe(wb, file_path)
        logger.info("Bulk added %d comments to %s in %s", count, sheet_name, file_path)
        return {"added": count}
    finally:
        wb.close()


def delete_comments_bulk(
    file_path: str,
    sheet_name: str,
    cells: list[str],
) -> dict:
    """Delete comments from multiple cells at once."""
    wb = load_workbook_safe(file_path)
    try:
        ws = get_sheet(wb, sheet_name)
        count = 0
        for cell_ref in cells:
            if ws[cell_ref].comment is not None:
                ws[cell_ref].comment = None
                count += 1
        save_workbook_safe(wb, file_path)
        logger.info("Bulk deleted %d comments from %s in %s", count, sheet_name, file_path)
        return {"deleted": count}
    finally:
        wb.close()
