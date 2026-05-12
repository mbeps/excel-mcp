"""Sandbox security tests for Excel MCP server.

Validates execute_custom_code restrictions, import blocking, and stdout capture.
"""

from __future__ import annotations

import sys
import io
from pathlib import Path

import pytest

from mcp_server.tools.custom_code import execute_custom_code
from mcp_server.tools.workbook import create_workbook


# ---------------------------------------------------------------------------
# 8.6, 8.7, 8.8 Sandbox — Imports Blocked
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("import_stmt", [
    "import os",
    "import subprocess",
    "import socket",
    "from os import path",
    "import sys",
])
def test_custom_code_imports_blocked(tmp_path: Path, import_stmt: str) -> None:
    """Import statements should be blocked by AST validation."""
    fp = str(tmp_path / "test.xlsx")
    create_workbook(fp)
    
    code = f"{import_stmt}\nresult = 1"
    result = execute_custom_code(file_path=fp, code=code)
    
    assert result["status"] == "error"
    assert "import statements are not allowed" in result["message"]


# ---------------------------------------------------------------------------
# 8.9 Sandbox — __import__ Dynamic Import Blocked
# ---------------------------------------------------------------------------

def test_custom_code_dunder_import_blocked(tmp_path: Path) -> None:
    """__import__ should be blocked to prevent dynamic import bypass."""
    fp = str(tmp_path / "test.xlsx")
    create_workbook(fp)
    
    code = "__import__('os').listdir('/')"
    result = execute_custom_code(file_path=fp, code=code)
    
    assert result["status"] == "error"
    assert "not allowed" in result["message"]


# ---------------------------------------------------------------------------
# 8.10 Sandbox — print(...) Stdout Corruption Risk
# ---------------------------------------------------------------------------

def test_custom_code_print_does_not_corrupt_response(tmp_path: Path) -> None:
    """print() should be handled gracefully. 
    Note: The report says this is a known gap where it might corrupt stdout.
    This test verifies current behavior and ensures the tool returns a valid dict.
    """
    fp = str(tmp_path / "test.xlsx")
    create_workbook(fp)
    
    code = "print('hello from sandbox')\nresult = 42"
    
    # Capture stdout to see if it leaks
    captured_stdout = io.StringIO()
    old_stdout = sys.stdout
    sys.stdout = captured_stdout
    try:
        result = execute_custom_code(file_path=fp, code=code)
    finally:
        sys.stdout = old_stdout
        
    assert result["status"] == "success"
    assert result["result"] == 42


# ---------------------------------------------------------------------------
# 8.11 Sandbox — ValueError Raised by User Code
# ---------------------------------------------------------------------------

def test_custom_code_user_raises_value_error(tmp_path: Path) -> None:
    """Exceptions in user code should be caught and returned as structured errors."""
    fp = str(tmp_path / "test.xlsx")
    create_workbook(fp)
    
    code = "raise ValueError('intentional error')"
    result = execute_custom_code(file_path=fp, code=code)
    
    assert result["status"] == "error"
    assert "intentional error" in result["message"]


# ---------------------------------------------------------------------------
# 8.12 Sandbox — Syntactically Invalid Code
# ---------------------------------------------------------------------------

def test_custom_code_syntax_error(tmp_path: Path) -> None:
    """Syntax errors should be caught and returned as structured errors."""
    fp = str(tmp_path / "test.xlsx")
    create_workbook(fp)
    
    code = "def broken(\nresult = 1"
    result = execute_custom_code(file_path=fp, code=code)
    
    assert result["status"] == "error"
    assert "Syntax error" in result["message"]
