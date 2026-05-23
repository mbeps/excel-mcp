from __future__ import annotations

import os
from typing import Any

from mcp.types import ToolAnnotations

from mcp_server.utils.excel_helpers import validate_file_path
from mcp_server.utils.file_resolver import encode_file_to_base64, resolve_to_local, session_store

__all__ = [
    "upload_file",
    "download_file",
    "release_file",
]


def upload_file(
    file_content: str | list[str],
    filename: str | list[str] = "workbook.xlsx",
) -> dict:
    """Upload a file for server-side processing.

    Accepts a base64-encoded file string or an HTTP/HTTPS URL. Returns a session_id
    and file_path that can be used with any other tool. The file is stored temporarily
    on the server and auto-deleted after 1 hour or when release_file is called.

    Args:
        file_content: Base64-encoded file bytes OR an HTTP/HTTPS URL pointing to the file.
        filename: Filename hint used to determine the file extension (e.g. "data.xlsx", "report.csv").

    Returns:
        dict: A mapping containing:
            - session_id (str): Opaque ID for this file session.
            - file_path (str): Server-side local path — use this as file_path in other tools.
            - message (str): Human-readable confirmation.
    """
    if isinstance(file_content, list):
        if isinstance(filename, list):
            if len(filename) != len(file_content):
                raise ValueError(
                    "When passing multiple files, filename list length must match file_content list length"
                )
            filename_hints = filename
        else:
            filename_hints = [filename] * len(file_content)

        items: list[dict] = []
        for content, name_hint in zip(file_content, filename_hints, strict=True):
            path, is_temp = resolve_to_local(content, name_hint)
            session_id = session_store.register(path, is_temp=is_temp)
            items.append(
                {
                    "session_id": session_id,
                    "file_path": str(path),
                    "message": f"File '{os.path.basename(path)}' is ready for processing.",
                }
            )

        return {
            "files": items,
            "count": len(items),
            "message": f"Uploaded {len(items)} files.",
        }

    filename_hint = filename[0] if isinstance(filename, list) and filename else "workbook.xlsx"
    path, is_temp = resolve_to_local(file_content, filename_hint)
    session_id = session_store.register(path, is_temp=is_temp)
    return {
        "session_id": session_id,
        "file_path": str(path),
        "message": f"File '{os.path.basename(path)}' is ready for processing.",
    }


def download_file(file_path: str) -> dict:
    """Download a file from the server as base64.

    Use after tools that modify a workbook to retrieve the updated file content.

    Args:
        file_path: Local path to the file (typically from upload_file or a tool that created a file).

    Returns:
        dict: A mapping containing:
            - file_content (str): Base64-encoded file bytes.
            - filename (str): Basename of the file.
            - size_bytes (int): Size of the file in bytes.
    """
    validated = validate_file_path(file_path, must_exist=True)
    encoded = encode_file_to_base64(str(validated))
    return {
        "file_content": encoded,
        "filename": os.path.basename(str(validated)),
        "size_bytes": os.path.getsize(str(validated)),
    }


def release_file(session_id: str) -> dict:
    """Release a session file and delete it from the server.

    Call this when you are done with a file to free server disk space.

    Args:
        session_id: The session_id returned by upload_file.

    Returns:
        dict: A mapping containing:
            - success (bool): Whether the file was successfully released.
            - message (str): Human-readable result description.
    """
    result = session_store.delete(session_id)
    return {
        "success": result,
        "message": "File released." if result else "Session not found or already released.",
    }


def register(mcp: Any) -> None:
    """Register file transfer tools on *mcp*."""
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False))(upload_file)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=True))(download_file)
    mcp.tool(annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=True))(release_file)
