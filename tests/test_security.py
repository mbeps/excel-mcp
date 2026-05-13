"""Security and governance tests for Excel MCP server.

Validates EXCEL_MCP_ALLOWED_DIRS, path traversal, and symlink protections.
"""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from mcp_server.routes.cell_ops import write_cells
from mcp_server.tools.workbook import create_workbook
from mcp_server.utils.excel_helpers import validate_file_path

# ---------------------------------------------------------------------------
# 8.1 EXCEL_MCP_ALLOWED_DIRS — Not Set (No Restriction)
# ---------------------------------------------------------------------------

def test_allowed_dirs_unset_no_restriction(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """When EXCEL_MCP_ALLOWED_DIRS is unset, any path should be allowed."""
    monkeypatch.delenv("EXCEL_MCP_ALLOWED_DIRS", raising=False)
    
    file_path = str(tmp_path / "test_unset.xlsx")
    result = create_workbook(file_path=file_path)
    
    assert result.file_path == file_path
    assert Path(file_path).exists()


# ---------------------------------------------------------------------------
# 8.2 EXCEL_MCP_ALLOWED_DIRS Set — Path Within Allowed Directory
# ---------------------------------------------------------------------------

def test_allowed_dirs_set_path_inside_succeeds(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """When path is inside allowed directory, operation should succeed."""
    allowed_dir = tmp_path / "allowed"
    allowed_dir.mkdir()
    monkeypatch.setenv("EXCEL_MCP_ALLOWED_DIRS", str(allowed_dir))
    
    file_path = str(allowed_dir / "allowed.xlsx")
    result = create_workbook(file_path=file_path)
    
    assert result.file_path == file_path
    assert Path(file_path).exists()
    
    # Also test write_cells
    write_result = write_cells(mode="single", file_path=file_path, sheet_name="Sheet", cell_ref="A1", value="test")
    assert "set" in str(write_result).lower()


# ---------------------------------------------------------------------------
# 8.3 EXCEL_MCP_ALLOWED_DIRS Set — Path Outside Allowed Directory
# ---------------------------------------------------------------------------

def test_allowed_dirs_set_path_outside_fails(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """When path is outside allowed directory, operation should fail with PermissionError."""
    allowed_dir = tmp_path / "allowed"
    allowed_dir.mkdir()
    forbidden_dir = tmp_path / "forbidden"
    forbidden_dir.mkdir()
    monkeypatch.setenv("EXCEL_MCP_ALLOWED_DIRS", str(allowed_dir))
    
    file_path = str(forbidden_dir / "forbidden.xlsx")
    with pytest.raises(ValueError, match="outside allowed directories"):
        create_workbook(file_path=file_path)


# ---------------------------------------------------------------------------
# 8.4 Path Traversal Attempts
# ---------------------------------------------------------------------------

def test_security_traversal_blocked(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Paths containing '..' should be blocked if they resolve outside allowed dirs."""
    allowed_dir = tmp_path / "allowed"
    allowed_dir.mkdir()
    monkeypatch.setenv("EXCEL_MCP_ALLOWED_DIRS", str(allowed_dir))
    
    # Resolves to tmp_path/test.xlsx which is outside allowed_dir
    traversal_path = str(allowed_dir / ".." / "test.xlsx")
    with pytest.raises(ValueError, match="outside allowed directories"):
        create_workbook(file_path=traversal_path)


def test_security_absolute_traversal_blocked(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Absolute paths outside allowed dirs should be blocked even if passed as string."""
    allowed_dir = tmp_path / "allowed"
    allowed_dir.mkdir()
    monkeypatch.setenv("EXCEL_MCP_ALLOWED_DIRS", str(allowed_dir))
    
    with pytest.raises(ValueError, match="outside allowed directories"):
        create_workbook(file_path="/etc/passwd.xlsx")


# ---------------------------------------------------------------------------
# 8.5 Symlink Protection
# ---------------------------------------------------------------------------

def test_security_symlink_blocked(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Symlinks resolving outside allowed dirs should be blocked."""
    allowed_dir = tmp_path / "allowed"
    allowed_dir.mkdir()
    forbidden_dir = tmp_path / "forbidden"
    forbidden_dir.mkdir()
    forbidden_file = forbidden_dir / "secret.xlsx"
    forbidden_file.touch()
    
    symlink_path = allowed_dir / "link.xlsx"
    os.symlink(forbidden_file, symlink_path)
    
    monkeypatch.setenv("EXCEL_MCP_ALLOWED_DIRS", str(allowed_dir))
    
    # Should fail when validating the path for read/write
    with pytest.raises(ValueError, match="outside allowed directories"):
        validate_file_path(str(symlink_path))
