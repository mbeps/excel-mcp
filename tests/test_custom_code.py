"""Tests for custom_code.py — sandboxed Python/pandas code execution."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from mcp_server.tools.custom_code import execute_custom_code


@pytest.fixture()
def sample_xlsx(tmp_path: Path) -> str:
    """Create a sample xlsx with Name/Age/Salary data."""
    fp = tmp_path / "test.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Name", "Age", "Salary"])
    ws.append(["Alice", 30, 50000])
    ws.append(["Bob", 25, 60000])
    ws.append(["Charlie", 35, 70000])
    wb.save(str(fp))
    wb.close()
    return str(fp)


def test_execute_basic_code(sample_xlsx: str) -> None:
    code = "result = len(df)"
    r = execute_custom_code(sample_xlsx, code)
    assert r["status"] == "success"


def test_execute_pandas_operations(sample_xlsx: str) -> None:
    code = "result = df['Age'].mean()"
    r = execute_custom_code(sample_xlsx, code)
    assert r["status"] == "success"


def test_execute_df_mutation(sample_xlsx: str, tmp_path: Path) -> None:
    output = str(tmp_path / "output.xlsx")
    code = "df['Bonus'] = df['Salary'] * 0.1\nresult = df"
    r = execute_custom_code(sample_xlsx, code, output_file=output)
    assert r["status"] == "success"


def test_execute_blocks_import(sample_xlsx: str) -> None:
    code = "import os"
    r = execute_custom_code(sample_xlsx, code)
    assert r["status"] == "error"


def test_execute_blocks_eval(sample_xlsx: str) -> None:
    code = "eval('1+1')"
    r = execute_custom_code(sample_xlsx, code)
    assert r["status"] == "error"


def test_execute_blocks_dunder(sample_xlsx: str) -> None:
    code = "x = df.__class__.__bases__"
    r = execute_custom_code(sample_xlsx, code)
    assert r["status"] == "error"


def test_execute_blocks_open(sample_xlsx: str) -> None:
    code = "f = open('/etc/passwd')"
    r = execute_custom_code(sample_xlsx, code)
    assert r["status"] == "error"
