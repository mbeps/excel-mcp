"""Tests for mcp_server.routes.file_transfer tool functions."""

from __future__ import annotations

import base64
import os
from pathlib import Path

import pytest
from openpyxl import Workbook

from mcp_server.routes.file_transfer import download_file, release_file, upload_file
from mcp_server.utils.file_resolver import session_store

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_xlsx(path: Path) -> Path:
    """Write a minimal xlsx workbook to *path* and return it."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "hello"
    wb.save(str(path))
    return path


# ---------------------------------------------------------------------------
# upload_file — base64 input
# ---------------------------------------------------------------------------


def test_upload_file_base64_returns_required_keys(tmp_path: Path) -> None:
    p = _make_xlsx(tmp_path / "upload_b64.xlsx")
    encoded = base64.b64encode(p.read_bytes()).decode("ascii")

    result = upload_file(file_content=encoded, filename="test.xlsx")

    assert "session_id" in result
    assert "file_path" in result
    assert "message" in result

    # Clean up
    release_file(result["session_id"])


def test_upload_file_base64_file_exists_on_disk(tmp_path: Path) -> None:
    # Use bytes whose base64 has no '/' (os.sep on Linux) so detect_input_type
    # correctly classifies the payload as "base64" and creates a real temp file.
    safe_bytes = b"\x00" * 20
    encoded = base64.b64encode(safe_bytes).decode("ascii")

    result = upload_file(file_content=encoded, filename="test.xlsx")
    assert os.path.exists(result["file_path"])

    # Clean up
    release_file(result["session_id"])


# ---------------------------------------------------------------------------
# upload_file — local path input
# ---------------------------------------------------------------------------


def test_upload_file_local_path_returns_required_keys(tmp_path: Path) -> None:
    p = _make_xlsx(tmp_path / "upload_local.xlsx")

    result = upload_file(file_content=str(p), filename="test.xlsx")

    assert "session_id" in result
    assert "file_path" in result
    assert "message" in result

    # Clean up
    release_file(result["session_id"])


def test_upload_file_local_path_file_path_matches_input(tmp_path: Path) -> None:
    p = _make_xlsx(tmp_path / "upload_local_path.xlsx")

    result = upload_file(file_content=str(p), filename="test.xlsx")
    assert result["file_path"] == str(p)

    # Clean up
    release_file(result["session_id"])


def test_release_file_local_upload_does_not_delete_source_file(tmp_path: Path) -> None:
    p = _make_xlsx(tmp_path / "local_source.xlsx")
    result = upload_file(file_content=str(p), filename="test.xlsx")

    release_result = release_file(result["session_id"])

    assert release_result["success"] is True
    assert p.exists()


def test_upload_file_multiple_returns_file_list(tmp_path: Path) -> None:
    p1 = _make_xlsx(tmp_path / "first.xlsx")
    p2 = _make_xlsx(tmp_path / "second.xlsx")

    payload_1 = base64.b64encode(p1.read_bytes()).decode("ascii")
    payload_2 = base64.b64encode(p2.read_bytes()).decode("ascii")

    result = upload_file(
        file_content=[payload_1, payload_2],
        filename=["one.xlsx", "two.xlsx"],
    )

    assert result["count"] == 2
    assert len(result["files"]) == 2
    assert "session_id" in result["files"][0]
    assert "file_path" in result["files"][0]
    assert "session_id" in result["files"][1]
    assert "file_path" in result["files"][1]

    for item in result["files"]:
        assert os.path.exists(item["file_path"])
        release_file(item["session_id"])


def test_upload_file_multiple_rejects_mismatched_filename_lengths(tmp_path: Path) -> None:
    p1 = _make_xlsx(tmp_path / "one.xlsx")
    p2 = _make_xlsx(tmp_path / "two.xlsx")
    payload_1 = base64.b64encode(p1.read_bytes()).decode("ascii")
    payload_2 = base64.b64encode(p2.read_bytes()).decode("ascii")

    with pytest.raises(ValueError, match="filename list length"):
        upload_file(
            file_content=[payload_1, payload_2],
            filename=["only-one.xlsx"],
        )


# ---------------------------------------------------------------------------
# download_file
# ---------------------------------------------------------------------------


def test_download_file_returns_required_keys(tmp_path: Path) -> None:
    p = _make_xlsx(tmp_path / "download_me.xlsx")

    result = download_file(str(p))

    assert "file_content" in result
    assert "filename" in result
    assert "size_bytes" in result


def test_download_file_content_matches_original_bytes(tmp_path: Path) -> None:
    p = _make_xlsx(tmp_path / "download_content.xlsx")
    original_bytes = p.read_bytes()

    result = download_file(str(p))
    decoded = base64.b64decode(result["file_content"])

    assert decoded == original_bytes


def test_download_file_size_bytes_is_correct(tmp_path: Path) -> None:
    p = _make_xlsx(tmp_path / "download_size.xlsx")

    result = download_file(str(p))

    assert result["size_bytes"] == p.stat().st_size


def test_download_file_filename_is_basename(tmp_path: Path) -> None:
    p = _make_xlsx(tmp_path / "my_workbook.xlsx")

    result = download_file(str(p))

    assert result["filename"] == "my_workbook.xlsx"


# ---------------------------------------------------------------------------
# release_file — success path
# ---------------------------------------------------------------------------


def test_release_file_success_true(tmp_path: Path) -> None:
    p = _make_xlsx(tmp_path / "release_me.xlsx")
    session_id = session_store.register(str(p))

    result = release_file(session_id)

    assert result["success"] is True


def test_release_file_removes_file_from_disk(tmp_path: Path) -> None:
    p = _make_xlsx(tmp_path / "release_disk.xlsx")
    session_id = session_store.register(str(p), is_temp=True)

    release_file(session_id)

    assert not p.exists()


# ---------------------------------------------------------------------------
# release_file — unknown session
# ---------------------------------------------------------------------------


def test_release_file_unknown_session_success_false() -> None:
    result = release_file("nonexistent-session-id-abc123")
    assert result["success"] is False
