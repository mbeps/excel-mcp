"""Tests for mcp_server.utils.file_resolver."""

from __future__ import annotations

import base64
import os
import time
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from mcp_server.utils.file_resolver import (
    SessionFileStore,
    detect_input_type,
    encode_file_to_base64,
    resolve_to_local,
)

# ---------------------------------------------------------------------------
# detect_input_type
# ---------------------------------------------------------------------------


def test_detect_input_type_https() -> None:
    assert detect_input_type("https://example.com/file.xlsx") == "http_url"


def test_detect_input_type_http() -> None:
    assert detect_input_type("http://example.com/file.xlsx") == "http_url"


def test_detect_input_type_cloud_s3() -> None:
    assert detect_input_type("s3://my-bucket/data/file.xlsx") == "cloud_url"


def test_detect_input_type_local_existing_file(tmp_path: Path) -> None:
    p = tmp_path / "real.xlsx"
    p.write_bytes(b"fake content")
    assert detect_input_type(str(p)) == "local_path"


def test_detect_input_type_local_absolute_path() -> None:
    assert detect_input_type("/some/path/file.xlsx") == "local_path"


def test_detect_input_type_base64() -> None:
    encoded = base64.b64encode(b"hello world").decode("ascii")
    assert detect_input_type(encoded) == "base64"


def test_detect_input_type_base64_with_slash_still_detected() -> None:
    encoded = base64.b64encode(b"\xff\xff").decode("ascii")
    assert "/" in encoded
    assert detect_input_type(encoded) == "base64"


def test_detect_input_type_fallback_unrecognised() -> None:
    # A string that doesn't match any special pattern falls back to local_path
    assert detect_input_type("unrecognised-random-value-!!@@") == "local_path"


# ---------------------------------------------------------------------------
# resolve_to_local — local path
# ---------------------------------------------------------------------------


def test_resolve_to_local_existing_file_returns_path_unchanged(tmp_path: Path) -> None:
    p = tmp_path / "existing.xlsx"
    p.write_bytes(b"data")
    result_path, is_temp = resolve_to_local(str(p))
    assert result_path == str(p)
    assert is_temp is False


def test_resolve_to_local_local_does_not_create_temp_file(tmp_path: Path) -> None:
    p = tmp_path / "file.xlsx"
    p.write_bytes(b"data")
    before = set(os.listdir(os.path.dirname(p)))
    resolve_to_local(str(p))
    after = set(os.listdir(os.path.dirname(p)))
    assert before == after


# ---------------------------------------------------------------------------
# resolve_to_local — cloud URL
# ---------------------------------------------------------------------------


def test_resolve_to_local_cloud_url_raises_value_error() -> None:
    with pytest.raises(ValueError, match="fsspec"):
        resolve_to_local("s3://bucket/key/file.xlsx")


# ---------------------------------------------------------------------------
# resolve_to_local — base64
# ---------------------------------------------------------------------------


def test_resolve_to_local_base64_creates_temp_file(tmp_path: Path) -> None:
    # Use bytes whose base64 encoding contains no '/' (os.sep on Linux), so
    # detect_input_type classifies the string as "base64" rather than "local_path".
    # b"\x00" * 20  →  "AAAAAAAAAAAAAAAAAAAAAAAAAAAAA=" (no '/' or '.')
    safe_bytes = b"\x00" * 20
    encoded = base64.b64encode(safe_bytes).decode("ascii")

    result_path, is_temp = resolve_to_local(encoded, "test.xlsx")
    try:
        assert is_temp is True
        assert os.path.exists(result_path)
        assert os.path.basename(result_path).startswith("excel_mcp_")
    finally:
        if os.path.exists(result_path):
            os.remove(result_path)


def test_resolve_to_local_base64_content_is_correct(tmp_path: Path) -> None:
    original_bytes = b"arbitrary file bytes for testing"
    encoded = base64.b64encode(original_bytes).decode("ascii")
    result_path, _ = resolve_to_local(encoded, "payload.bin")
    try:
        assert Path(result_path).read_bytes() == original_bytes
    finally:
        if os.path.exists(result_path):
            os.remove(result_path)


# ---------------------------------------------------------------------------
# resolve_to_local — HTTP URL (mocked)
# ---------------------------------------------------------------------------


def test_resolve_to_local_http_url_creates_temp_file() -> None:
    fake_bytes = b"PK\x03\x04fake xlsx bytes"

    # Build a mock response that supports context manager and iter_bytes
    mock_response = MagicMock()
    mock_response.__enter__ = MagicMock(return_value=mock_response)
    mock_response.__exit__ = MagicMock(return_value=False)
    mock_response.raise_for_status = MagicMock()
    mock_response.iter_bytes = MagicMock(return_value=iter([fake_bytes]))

    with patch("httpx.stream", return_value=mock_response) as mock_stream:
        result_path, is_temp = resolve_to_local("https://example.com/file.xlsx", "file.xlsx")
        mock_stream.assert_called_once()

    try:
        assert is_temp is True
        assert os.path.exists(result_path)
        assert Path(result_path).read_bytes() == fake_bytes
    finally:
        if os.path.exists(result_path):
            os.remove(result_path)


# ---------------------------------------------------------------------------
# encode_file_to_base64
# ---------------------------------------------------------------------------


def test_encode_file_to_base64_round_trips(tmp_path: Path) -> None:
    original = b"known content for base64 round-trip test"
    p = tmp_path / "encode_me.bin"
    p.write_bytes(original)
    encoded = encode_file_to_base64(str(p))
    decoded = base64.b64decode(encoded)
    assert decoded == original


def test_encode_file_to_base64_returns_ascii_string(tmp_path: Path) -> None:
    p = tmp_path / "ascii_check.bin"
    p.write_bytes(b"\x00\xff\xfe")
    encoded = encode_file_to_base64(str(p))
    assert isinstance(encoded, str)
    encoded.encode("ascii")  # must not raise


# ---------------------------------------------------------------------------
# SessionFileStore
# ---------------------------------------------------------------------------


@pytest.fixture()
def file_store() -> SessionFileStore:
    return SessionFileStore(ttl=3600.0)


def test_session_store_register_returns_nonempty_id(file_store: SessionFileStore, tmp_path: Path) -> None:
    p = tmp_path / "reg.xlsx"
    p.write_bytes(b"x")
    session_id = file_store.register(str(p))
    assert session_id
    assert len(session_id) > 0


def test_session_store_get_returns_path_after_register(file_store: SessionFileStore, tmp_path: Path) -> None:
    p = tmp_path / "get_test.xlsx"
    p.write_bytes(b"x")
    session_id = file_store.register(str(p))
    retrieved = file_store.get(session_id)
    assert retrieved is not None
    assert os.path.abspath(str(p)) == retrieved


def test_session_store_delete_returns_true_and_removes_file(file_store: SessionFileStore, tmp_path: Path) -> None:
    p = tmp_path / "delete_test.xlsx"
    p.write_bytes(b"x")
    session_id = file_store.register(str(p), is_temp=True)
    result = file_store.delete(session_id)
    assert result is True
    assert not p.exists()


def test_session_store_delete_non_temp_keeps_file(file_store: SessionFileStore, tmp_path: Path) -> None:
    p = tmp_path / "non_temp.xlsx"
    p.write_bytes(b"x")
    session_id = file_store.register(str(p), is_temp=False)

    result = file_store.delete(session_id)

    assert result is True
    assert p.exists()


def test_session_store_delete_returns_false_for_unknown_id(file_store: SessionFileStore) -> None:
    result = file_store.delete("nonexistent-session-id-xyz")
    assert result is False


def test_session_store_expired_session_returns_none(tmp_path: Path) -> None:
    short_ttl_store = SessionFileStore(ttl=0.01)
    p = tmp_path / "expire_test.xlsx"
    p.write_bytes(b"x")
    session_id = short_ttl_store.register(str(p))
    time.sleep(0.05)
    result = short_ttl_store.get(session_id)
    assert result is None
