"""Resolve remote/encoded file inputs to local paths.

This module supports local paths, HTTP URLs, base64-encoded payloads, and a
session-backed temporary-file lifecycle for tool workflows.
"""

from __future__ import annotations

import base64
import os
import re
import tempfile
import threading
import time
import uuid
from pathlib import Path
from typing import Literal

import httpx

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_CLOUD_URI_PATTERN: re.Pattern[str] = re.compile(r"^[a-zA-Z][a-zA-Z0-9+\-.]+://")
_HTTP_SCHEMES: tuple[str, ...] = ("http://", "https://")

# ---------------------------------------------------------------------------
# Input detection
# ---------------------------------------------------------------------------


def detect_input_type(
    value: str,
) -> Literal["local_path", "http_url", "cloud_url", "base64"]:
    """Detect the type of a file input string.

    Applies the following rules in order:

    1. Starts with ``http://`` or ``https://`` → ``"http_url"``.
    2. Matches a non-HTTP URI scheme (e.g. ``s3://``, ``gs://``) → ``"cloud_url"``.
    3. Contains ``os.sep``, starts with ``.`` or ``/``, or the path exists on
       disk → ``"local_path"``.
    4. Decodes cleanly as standard base64 → ``"base64"``.
    5. Fallback → ``"local_path"`` (validation is deferred to downstream callers).

    Args:
        value: The raw input string to classify.

    Returns:
        A literal string identifying the detected input type.
    """
    if value.startswith(_HTTP_SCHEMES):
        return "http_url"

    if _CLOUD_URI_PATTERN.match(value) and not value.startswith(_HTTP_SCHEMES):
        return "cloud_url"

    try:
        base64.b64decode(value, validate=True)
        return "base64"
    except Exception:
        pass

    if os.sep in value or value.startswith((".", "/")) or Path(value).exists():
        return "local_path"

    return "local_path"


# ---------------------------------------------------------------------------
# Temp file helpers
# ---------------------------------------------------------------------------


def _make_temp_path(filename_hint: str) -> str:
    """Return a temp file path with a suffix derived from *filename_hint*.

    Args:
        filename_hint: A filename whose suffix (e.g. ``.xlsx``) is used for the
            temp file.

    Returns:
        Absolute path string for the new temp file (not yet created on disk).
    """
    suffix = Path(filename_hint).suffix or ".xlsx"
    fd, path = tempfile.mkstemp(prefix="excel_mcp_", suffix=suffix, dir=tempfile.gettempdir())
    os.close(fd)
    return path


# ---------------------------------------------------------------------------
# Primary resolver
# ---------------------------------------------------------------------------


def resolve_to_local(
    value: str,
    filename_hint: str = "file.xlsx",
) -> tuple[str, bool]:
    """Resolve a file input to a local filesystem path.

    Args:
        value: The raw input — a local path, HTTP URL, cloud URL, or base64
            string.
        filename_hint: Used to derive the temp file suffix when a temp file must
            be created.  Defaults to ``"file.xlsx"``.

    Returns:
        A two-tuple ``(local_path, is_temp)`` where *is_temp* is ``True`` when a
        temporary file was created and the caller is responsible for cleanup (or
        may delegate to :class:`SessionFileStore`).

    Raises:
        ValueError: When a cloud URL is supplied (not implemented in v1).
        httpx.HTTPStatusError: When the HTTP download fails with a non-2xx
            status.
        Exception: Propagated from base64 decoding failures.
    """
    input_type = detect_input_type(value)

    if input_type == "local_path":
        return value, False

    if input_type == "cloud_url":
        raise ValueError("Cloud URL support requires fsspec. Install fsspec and s3fs for S3 support.")

    if input_type == "http_url":
        return _download_http(value, filename_hint)

    # base64
    return _decode_base64(value, filename_hint)


def _download_http(url: str, filename_hint: str) -> tuple[str, bool]:
    """Stream-download an HTTP URL to a temp file.

    Args:
        url: The HTTP or HTTPS URL to download.
        filename_hint: Used to derive the temp file suffix.

    Returns:
        ``(temp_path, True)`` on success.

    Raises:
        httpx.HTTPStatusError: On non-2xx HTTP responses.
        Exception: On any other download error, after cleaning up the temp file.
    """
    temp_path = _make_temp_path(filename_hint)
    try:
        with httpx.stream("GET", url, follow_redirects=True) as response:
            response.raise_for_status()
            with open(temp_path, "wb") as fh:
                for chunk in response.iter_bytes(chunk_size=8192):
                    fh.write(chunk)
    except Exception:
        _safe_remove(temp_path)
        raise
    return temp_path, True


def _decode_base64(encoded: str, filename_hint: str) -> tuple[str, bool]:
    """Decode a base64 string and write it to a temp file.

    Args:
        encoded: Standard base64-encoded file content.
        filename_hint: Used to derive the temp file suffix.

    Returns:
        ``(temp_path, True)`` on success.

    Raises:
        Exception: On decoding failure, after cleaning up the temp file.
    """
    temp_path = _make_temp_path(filename_hint)
    try:
        data = base64.b64decode(encoded)
        with open(temp_path, "wb") as fh:
            fh.write(data)
    except Exception:
        _safe_remove(temp_path)
        raise
    return temp_path, True


def _safe_remove(path: str) -> None:
    """Delete *path* from disk, ignoring errors.

    Args:
        path: Filesystem path to remove.
    """
    try:
        os.remove(path)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Base64 encoding helper
# ---------------------------------------------------------------------------


def encode_file_to_base64(path: str) -> str:
    """Read a file and return its standard base64-encoded representation.

    Args:
        path: Absolute or relative path to the file to encode.

    Returns:
        A base64-encoded string of the file's binary contents.
    """
    with open(path, "rb") as fh:
        return base64.b64encode(fh.read()).decode("ascii")


# ---------------------------------------------------------------------------
# SessionFileStore
# ---------------------------------------------------------------------------


class SessionFileStore:
    """Thread-safe store mapping session IDs to temporary file paths with TTL.

    Files that exceed the TTL are lazily cleaned up from disk on the next
    :meth:`register` or :meth:`get` call.

    Args:
        ttl: Time-to-live in seconds for each registered session.  Defaults to
            ``3600`` (one hour).

    Example::

        store = SessionFileStore()
        session_id = store.register("/tmp/excel_mcp_abc.xlsx")
        path = store.get(session_id)   # returns the path while within TTL
        store.delete(session_id)       # removes session and deletes the file
    """

    def __init__(self, ttl: float = 3600.0) -> None:
        self._ttl: float = ttl
        self._lock: threading.Lock = threading.Lock()
        # session_id -> (abs_path, created_at, is_temp)
        self._store: dict[str, tuple[str, float, bool]] = {}

    def register(self, path: str, is_temp: bool = True) -> str:
        """Register a file path and return a new session ID.

        Expired sessions are purged before the new entry is added.

        Args:
            path: Absolute path to the file to register.
            is_temp: Whether this file is server-managed temporary state and may
                be deleted during release/expiry cleanup.

        Returns:
            A UUID4 string that can be used to retrieve or delete the session.
        """
        session_id = str(uuid.uuid4())
        abs_path = os.path.abspath(path)
        with self._lock:
            self._cleanup_expired()
            self._store[session_id] = (abs_path, time.monotonic(), is_temp)
        return session_id

    def get(self, session_id: str) -> str | None:
        """Return the file path for *session_id* if it exists and has not expired.

        Expired entries are removed lazily.

        Args:
            session_id: The session identifier returned by :meth:`register`.

        Returns:
            The registered file path, or ``None`` if the session is missing or
            expired.
        """
        with self._lock:
            entry = self._store.get(session_id)
            if entry is None:
                return None
            abs_path, created_at, _is_temp = entry
            if time.monotonic() - created_at > self._ttl:
                del self._store[session_id]
                if _is_temp:
                    _safe_remove(abs_path)
                return None
            return abs_path

    def delete(self, session_id: str) -> bool:
        """Remove a session and delete its associated file from disk.

        Args:
            session_id: The session identifier to remove.

        Returns:
            ``True`` if the session was found and removed, ``False`` otherwise.
        """
        with self._lock:
            entry = self._store.pop(session_id, None)
        if entry is None:
            return False
        abs_path, _created_at, is_temp = entry
        if is_temp:
            _safe_remove(abs_path)
        return True

    def _cleanup_expired(self) -> None:
        """Remove all expired sessions and delete their files from disk.

        Must be called with :attr:`_lock` already held.
        """
        now = time.monotonic()
        expired = [sid for sid, (_, created_at, _) in self._store.items() if now - created_at > self._ttl]
        for sid in expired:
            path, _created_at, is_temp = self._store.pop(sid)
            if is_temp:
                _safe_remove(path)


# ---------------------------------------------------------------------------
# Module-level singleton
# ---------------------------------------------------------------------------

session_store: SessionFileStore = SessionFileStore()
