"""Resolve remote/encoded file inputs to local paths.

This module supports local paths, HTTP URLs, base64-encoded payloads, and a
session-backed temporary-file lifecycle for tool workflows.
"""

from __future__ import annotations

import base64
import ipaddress
import os
import re
import socket
import tempfile
import threading
import time
import urllib.parse
import uuid
from pathlib import Path
from typing import Literal

import httpx

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_CLOUD_URI_PATTERN: re.Pattern[str] = re.compile(r"^[a-zA-Z][a-zA-Z0-9+\-.]+://")
_HTTP_SCHEMES: tuple[str, ...] = ("http://", "https://")
_DEFAULT_HTTP_TIMEOUT: float = 30.0
_MAX_REDIRECTS: int = 5
_DEFAULT_MAX_FILE_BYTES: int = 50 * 1024 * 1024  # 50 MiB


def _get_max_file_bytes() -> int:
    """Return the configured maximum file size in bytes."""
    raw = os.environ.get("EXCEL_MCP_MAX_FILE_BYTES", "").strip()
    if raw:
        return int(raw)
    return _DEFAULT_MAX_FILE_BYTES


def _allow_private_urls() -> bool:
    """Return whether private/loopback URL targets are permitted via env opt-in."""
    return os.environ.get("EXCEL_MCP_ALLOW_PRIVATE_URLS", "").strip().lower() in ("1", "true", "yes")


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
    if not value or not value.strip():
        return "local_path"

    if value.startswith(_HTTP_SCHEMES):
        return "http_url"

    if _CLOUD_URI_PATTERN.match(value):
        return "cloud_url"

    try:
        if Path(value).exists():
            return "local_path"
    except OSError:
        pass

    try:
        base64.b64decode(value, validate=True)
        return "base64"
    except Exception:
        pass

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


def _validate_url_target(url: str) -> None:
    """Validate that *url* uses HTTP(S) and does not resolve to private/local IP ranges."""
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme not in ("http", "https") or not parsed.hostname:
        raise ValueError(f"Invalid HTTP(S) URL: {url}")

    if _allow_private_urls():
        return

    port = parsed.port or (443 if parsed.scheme == "https" else 80)
    try:
        addr_infos = socket.getaddrinfo(parsed.hostname, port, type=socket.SOCK_STREAM)
    except socket.gaierror as e:
        raise ValueError(f"Failed to resolve host '{parsed.hostname}': {e}") from e

    for _family, _type, _proto, _canonname, sockaddr in addr_infos:
        ip_str = str(sockaddr[0]).split("%", 1)[0]
        ip = ipaddress.ip_address(ip_str)
        if (
            ip.is_private
            or ip.is_loopback
            or ip.is_link_local
            or ip.is_multicast
            or ip.is_reserved
            or ip.is_unspecified
        ):
            raise ValueError(f"Blocked URL resolving to disallowed IP address: {ip}")


def _download_http(url: str, filename_hint: str) -> tuple[str, bool]:
    """Stream-download an HTTP URL to a temp file with SSRF, redirect, timeout, and size guards.

    Args:
        url: The HTTP or HTTPS URL to download.
        filename_hint: Used to derive the temp file suffix.

    Returns:
        ``(temp_path, True)`` on success.

    Raises:
        ValueError: On disallowed IP targets, excessive redirects, or size limit exceeded.
        httpx.HTTPStatusError: On non-2xx HTTP responses.
        Exception: On any other download error, after cleaning up the temp file.
    """
    max_bytes = _get_max_file_bytes()
    current_url = url
    _validate_url_target(current_url)

    temp_path = _make_temp_path(filename_hint)
    try:
        for _ in range(_MAX_REDIRECTS + 1):
            with httpx.stream(
                "GET",
                current_url,
                follow_redirects=False,
                timeout=_DEFAULT_HTTP_TIMEOUT,
            ) as response:
                if response.status_code in (301, 302, 303, 307, 308):
                    location = response.headers.get("location")
                    if not location:
                        raise ValueError("Redirect response missing Location header")
                    current_url = urllib.parse.urljoin(current_url, location)
                    _validate_url_target(current_url)
                    continue

                response.raise_for_status()
                content_length = response.headers.get("content-length")
                if isinstance(content_length, str) and content_length.isdigit() and int(content_length) > max_bytes:
                    raise ValueError(f"Remote file exceeds max allowed size ({max_bytes} bytes)")

                total_bytes = 0
                with open(temp_path, "wb") as fh:
                    for chunk in response.iter_bytes(chunk_size=8192):
                        total_bytes += len(chunk)
                        if total_bytes > max_bytes:
                            raise ValueError(f"Downloaded file exceeds max allowed size ({max_bytes} bytes)")
                        fh.write(chunk)
                return temp_path, True

        raise ValueError(f"Exceeded maximum redirects ({_MAX_REDIRECTS})")
    except Exception:
        _safe_remove(temp_path)
        raise


def _decode_base64(encoded: str, filename_hint: str) -> tuple[str, bool]:
    """Decode a base64 string and write it to a temp file.

    Args:
        encoded: Standard base64-encoded file content.
        filename_hint: Used to derive the temp file suffix.

    Returns:
        ``(temp_path, True)`` on success.

    Raises:
        ValueError: If decoded payload exceeds ``EXCEL_MCP_MAX_FILE_BYTES``.
        Exception: On decoding failure, after cleaning up the temp file.
    """
    max_bytes = _get_max_file_bytes()
    temp_path = _make_temp_path(filename_hint)
    try:
        data = base64.b64decode(encoded, validate=True)
        if len(data) > max_bytes:
            raise ValueError(f"Decoded payload exceeds max allowed size ({max_bytes} bytes)")
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

    Raises:
        ValueError: If file size exceeds ``EXCEL_MCP_MAX_FILE_BYTES``.
    """
    max_bytes = _get_max_file_bytes()
    if os.path.getsize(path) > max_bytes:
        raise ValueError(f"File exceeds max allowed size ({max_bytes} bytes)")
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
