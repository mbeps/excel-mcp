"""Phase 1 regression, security, and coverage tests."""

from __future__ import annotations

import asyncio
import base64
import binascii
import gc
import json
import os
import re
import socket
import subprocess
import tempfile
import urllib.parse
from pathlib import Path
from unittest.mock import MagicMock, patch

import httpx
import numpy as np
import openpyxl
import pandas as pd
import pytest
from openpyxl import Workbook
from openpyxl.utils.exceptions import InvalidFileException

from mcp_server.main import _NoopMCP, create_server, resource_list_sheets, resource_sheet_preview
from mcp_server.routes import _MCPToolWrapper, _wrap_tool_fn
from mcp_server.routes.file_transfer import upload_file
from mcp_server.tools.custom_code import execute_custom_code
from mcp_server.utils.excel_helpers import (
    _FILE_LOCKS,
    _get_file_lock,
    col_letter_to_index,
    get_sheet,
    load_hidden_json,
    load_workbook_safe,
    read_sheet_df,
    resolve_dataframe_column,
    sanitize_excel_value,
    save_hidden_json,
    save_workbook_safe,
    validate_excel_range,
    validate_file_path,
    write_dataframe_to_workbook,
)
from mcp_server.utils.expression_validator import validate_expression
from mcp_server.utils.file_resolver import (
    detect_input_type,
    encode_file_to_base64,
    resolve_to_local,
)


# ---------------------------------------------------------------------------
# main.py: create_server, _NoopMCP, URI decoding
# ---------------------------------------------------------------------------


def test_create_server_noop_and_real() -> None:
    noop = create_server(disable_registration=True)
    assert isinstance(noop, _NoopMCP)
    assert noop.streamable_http_app() is None
    noop.run()

    real = create_server(disable_registration=False)
    assert not isinstance(real, _NoopMCP)


def test_resource_uri_decoding(tmp_path: Path) -> None:
    fp = tmp_path / "my book.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "My Sheet"
    ws["A1"] = "Header"
    ws["A2"] = "Value"
    wb.save(str(fp))
    wb.close()

    encoded_path = urllib.parse.quote(str(fp), safe="")
    encoded_sheet = urllib.parse.quote("My Sheet", safe="")

    sheets = json.loads(resource_list_sheets(encoded_path))
    assert any(s["name"] == "My Sheet" for s in sheets)

    preview = json.loads(resource_sheet_preview(encoded_path, encoded_sheet))
    assert preview["rows"][0][0] == "Header"


# ---------------------------------------------------------------------------
# routes/__init__.py: async_wrapped & expanded ToolError boundary
# ---------------------------------------------------------------------------


def test_wrap_tool_fn_async_and_expanded_errors() -> None:
    from mcp.server.mcpserver.exceptions import ToolError

    async def ok_async(x: int) -> int:
        return x + 1

    async def fail_async() -> None:
        raise OSError("disk error")

    wrapped_ok = _wrap_tool_fn(ok_async)
    wrapped_fail = _wrap_tool_fn(fail_async)

    assert asyncio.run(wrapped_ok(41)) == 42
    with pytest.raises(ToolError, match="disk error"):
        asyncio.run(wrapped_fail())

    for exc in (
        re.error("bad regex"),
        binascii.Error("bad b64"),
        httpx.HTTPError("http fail"),
        InvalidFileException("bad xlsx"),
    ):

        def raise_exc(e: Exception = exc) -> None:
            raise e

        with pytest.raises(ToolError):
            _wrap_tool_fn(raise_exc)()


# ---------------------------------------------------------------------------
# excel_helpers.py: validate_file_path, save_workbook_safe, read_sheet_df,
# validate_excel_range, save_hidden_json, sanitize, resolve_col, write_df
# ---------------------------------------------------------------------------


def test_validate_file_path_temp_subdir_not_bypassed(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    monkeypatch.setenv("EXCEL_MCP_ALLOWED_DIRS", str(allowed))

    # A subdir inside system tempdir with excel_mcp_ prefix must NOT bypass ALLOWED_DIRS
    fake_temp_dir = Path(tempfile.gettempdir()) / "excel_mcp_subdir_test"
    fake_temp_dir.mkdir(exist_ok=True)
    target = fake_temp_dir / "evil.xlsx"
    target.write_bytes(b"x")
    try:
        with pytest.raises(ValueError, match="outside allowed directories"):
            validate_file_path(str(target))
    finally:
        target.unlink(missing_ok=True)
        fake_temp_dir.rmdir()


def test_validate_file_path_no_existence_oracle_outside_allowed(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    monkeypatch.setenv("EXCEL_MCP_ALLOWED_DIRS", str(allowed))

    existing_outside = tmp_path / "secret.xlsx"
    existing_outside.write_bytes(b"secret")
    missing_outside = tmp_path / "missing.xlsx"

    with pytest.raises(ValueError, match="outside allowed directories"):
        validate_file_path(str(existing_outside), must_exist=True)
    with pytest.raises(ValueError, match="outside allowed directories"):
        validate_file_path(str(missing_outside), must_exist=True)


def test_validate_file_path_allowed_extensions(tmp_path: Path) -> None:
    img = tmp_path / "logo.png"
    img.write_bytes(b"PNG")
    assert validate_file_path(str(img), allowed_extensions={".png", ".jpg"}) == img.resolve()
    with pytest.raises(ValueError, match="Unsupported file type"):
        validate_file_path(str(img))


def test_save_workbook_safe_guards_and_weak_locks(tmp_path: Path) -> None:
    fp = tmp_path / "ro.xlsx"
    wb = Workbook()
    wb.save(str(fp))
    wb.close()

    ro_wb = openpyxl.load_workbook(str(fp), read_only=True)
    try:
        with pytest.raises(ValueError, match="read-only"):
            save_workbook_safe(ro_wb, str(fp))
    finally:
        ro_wb.close()

    rw_wb = Workbook()
    try:
        with pytest.raises(ValueError, match="non-Excel"):
            save_workbook_safe(rw_wb, str(tmp_path / "out.csv"))
    finally:
        rw_wb.close()

    ephemeral = tmp_path / "ephemeral_lock.xlsx"
    lock = _get_file_lock(ephemeral)
    assert ephemeral in _FILE_LOCKS
    del lock
    gc.collect()
    assert ephemeral not in _FILE_LOCKS


def test_read_sheet_df_csv_and_guards(tmp_path: Path) -> None:
    csv_utf8 = tmp_path / "data.csv"
    csv_utf8.write_text("Name,Val\nAlice,10\n", encoding="utf-8")
    df = read_sheet_df(str(csv_utf8), "ignored", header_row=1)
    assert list(df.columns) == ["Name", "Val"]

    csv_latin1 = tmp_path / "latin1.csv"
    csv_latin1.write_bytes("Name,City\nJosé,São Paulo\n".encode("latin-1"))
    df_lat = read_sheet_df(str(csv_latin1), "ignored", header_row=1)
    assert df_lat.iloc[0]["Name"] == "José"

    with pytest.raises(ValueError, match="header_row"):
        read_sheet_df(str(csv_utf8), "ignored", header_row=True)
    with pytest.raises(ValueError, match="header_row"):
        read_sheet_df(str(csv_utf8), "ignored", header_row=-1)


def test_validate_excel_range_normalization_and_inverted() -> None:
    res = validate_excel_range(" a01 : b02 ")
    assert res["valid"] is True
    assert res["start_cell"] == "A1"
    assert res["end_cell"] == "B2"

    inv_row = validate_excel_range("A5:B2")
    assert inv_row["valid"] is False
    assert "precedes" in inv_row["message"]

    inv_col = validate_excel_range("C1:A3")
    assert inv_col["valid"] is False


def test_save_hidden_json_limit() -> None:
    wb = Workbook()
    try:
        with pytest.raises(ValueError, match="32767"):
            save_hidden_json(wb, "_meta", {"big": "x" * 33000})
    finally:
        wb.close()


def test_sanitize_excel_value_and_resolve_column() -> None:
    assert sanitize_excel_value(None) is None
    assert sanitize_excel_value(pd.NA) is None
    assert sanitize_excel_value(pd.NaT) is None
    assert sanitize_excel_value(np.bool_(True)) is True
    assert type(sanitize_excel_value(np.bool_(True))) is bool
    assert sanitize_excel_value(False) is False
    assert sanitize_excel_value(np.int64(7)) == 7
    assert type(sanitize_excel_value(np.int64(7))) is int
    assert sanitize_excel_value(np.float64("nan")) is None
    assert sanitize_excel_value(float("nan")) is None
    assert sanitize_excel_value(np.float64(1.5)) == 1.5
    assert sanitize_excel_value(pd.Timestamp("2026-01-01")).year == 2026
    assert sanitize_excel_value(pd.Timestamp("NaT")) is None

    df = pd.DataFrame({"ID": [1], "Age": [20], 0: ["zero"]})
    assert resolve_dataframe_column(df, "ID") == "ID"
    assert resolve_dataframe_column(df, "Age") == "Age"
    assert resolve_dataframe_column(df, "0") == 0
    assert resolve_dataframe_column(df, "age") == "Age"
    assert resolve_dataframe_column(df, "B") == "Age"
    with pytest.raises(ValueError, match="not found"):
        resolve_dataframe_column(df, "Z")


def test_write_dataframe_to_workbook_preserves_other_sheets(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.chdir(tmp_path)
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Data"
    ws1.append(["A", "B"])
    ws1.append([1, 2])
    ws2 = wb.create_sheet("KeepMe")
    ws2["A1"] = "Preserved"
    wb.save("book.xlsx")
    wb.close()

    new_df = pd.DataFrame({"X": [10, np.nan], "Y": [np.bool_(False), np.int64(5)]})
    write_dataframe_to_workbook(new_df, "./book.xlsx", sheet_name="Data")

    wb_check = openpyxl.load_workbook("book.xlsx")
    try:
        assert "KeepMe" in wb_check.sheetnames
        assert wb_check["KeepMe"]["A1"].value == "Preserved"
        assert wb_check["Data"]["A1"].value == "X"
        assert wb_check["Data"]["A3"].value is None
        assert wb_check["Data"]["B2"].value is False
    finally:
        wb_check.close()


# ---------------------------------------------------------------------------
# file_resolver.py & file_transfer.py: SSRF, redirects, size caps, upload
# ---------------------------------------------------------------------------


def test_detect_input_type_edge_cases(tmp_path: Path) -> None:
    assert detect_input_type("") == "local_path"
    assert detect_input_type("   ") == "local_path"
    long_b64 = base64.b64encode(b"a" * 300).decode("ascii")
    assert detect_input_type(long_b64) == "base64"


@pytest.mark.parametrize(
    "ip_str",
    ["127.0.0.1", "10.0.0.1", "192.168.1.10", "169.254.169.254", "::1"],
)
def test_ssrf_blocks_private_and_loopback_ips(ip_str: str) -> None:
    family = socket.AF_INET6 if ":" in ip_str else socket.AF_INET
    fake_addr = [(family, socket.SOCK_STREAM, 6, "", (ip_str, 80))]
    with patch("socket.getaddrinfo", return_value=fake_addr):
        with pytest.raises(ValueError, match="disallowed IP"):
            resolve_to_local("http://internal.example/book.xlsx")


def test_ssrf_blocks_redirect_to_loopback() -> None:
    public_addr = [(socket.AF_INET, socket.SOCK_STREAM, 6, "", ("93.184.216.34", 443))]
    private_addr = [(socket.AF_INET, socket.SOCK_STREAM, 6, "", ("127.0.0.1", 80))]

    redirect_resp = MagicMock()
    redirect_resp.__enter__ = MagicMock(return_value=redirect_resp)
    redirect_resp.__exit__ = MagicMock(return_value=False)
    redirect_resp.status_code = 302
    redirect_resp.headers = {"location": "http://127.0.0.1/secret.xlsx"}

    with patch("socket.getaddrinfo", side_effect=[public_addr, private_addr]):
        with patch("httpx.stream", return_value=redirect_resp):
            with pytest.raises(ValueError, match="disallowed IP"):
                resolve_to_local("https://example.com/file.xlsx")


def test_size_limits_on_base64_and_encode(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("EXCEL_MCP_MAX_FILE_BYTES", "10")
    payload = base64.b64encode(b"0123456789ABCDEF").decode("ascii")
    with pytest.raises(ValueError, match="max allowed size"):
        resolve_to_local(payload, "small.xlsx")

    fp = tmp_path / "big.xlsx"
    fp.write_bytes(b"0123456789ABCDEF")
    with pytest.raises(ValueError, match="max allowed size"):
        encode_file_to_base64(str(fp))


def test_upload_file_validates_local_path_and_extension(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    outside = tmp_path / "outside.xlsx"
    outside.write_bytes(b"x")
    monkeypatch.setenv("EXCEL_MCP_ALLOWED_DIRS", str(allowed))

    with pytest.raises(ValueError, match="outside allowed directories"):
        upload_file(file_content=str(outside), filename="outside.xlsx")

    b64 = base64.b64encode(b"hello").decode("ascii")
    with pytest.raises(ValueError, match="Unsupported file type"):
        upload_file(file_content=b64, filename="evil.exe")


# ---------------------------------------------------------------------------
# expression_validator.py: AST & Pow DoS hardening
# ---------------------------------------------------------------------------


def test_expression_validator_blocks_attr_subscript_and_pow_dos() -> None:
    with pytest.raises(ValueError, match="Unsafe expression node"):
        validate_expression("x.real", allowed_names=frozenset({"x"}))
    with pytest.raises(ValueError, match="Unsafe expression node"):
        validate_expression("x[0]", allowed_names=frozenset({"x"}))
    with pytest.raises(ValueError, match="Chained exponentiation"):
        validate_expression("2**3**4")
    with pytest.raises(ValueError, match="exceeds allowed limit"):
        validate_expression("2**1001")
    with pytest.raises(ValueError, match="exceeds allowed limit"):
        validate_expression("pow(2, -1001)")

    assert validate_expression("2**10 + pow(2, -3)") is not None


# ---------------------------------------------------------------------------
# custom_code.py: AST escapes, timeout, CSV, multi-sheet preservation
# ---------------------------------------------------------------------------


def test_custom_code_blocks_format_chain_and_xml(tmp_path: Path) -> None:
    fp = tmp_path / "book.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["A"])
    ws.append([1])
    wb.save(str(fp))
    wb.close()

    for bad_code in (
        'result = "{0.__class__}".format(df)',
        'result = "{0}".format_map({"0": 1})',
        "result = pd.compat.os.getcwd()",
        "result = abs.__self__",
        "result = df.to_xml()",
    ):
        res = execute_custom_code(str(fp), bad_code)
        assert res["status"] == "error", f"Expected error for: {bad_code}"


def test_custom_code_preserves_other_sheets_and_supports_csv(tmp_path: Path) -> None:
    fp = tmp_path / "multi.xlsx"
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Main"
    ws1.append(["Val"])
    ws1.append([10])
    ws2 = wb.create_sheet("Other")
    ws2["A1"] = "Keep"
    wb.save(str(fp))
    wb.close()

    res = execute_custom_code(str(fp), "df['Double'] = df['Val'] * 2", sheet="Main")
    assert res["status"] == "success"

    wb_after = openpyxl.load_workbook(str(fp))
    try:
        assert "Other" in wb_after.sheetnames
        assert wb_after["Other"]["A1"].value == "Keep"
        assert wb_after["Main"]["B2"].value == 20
    finally:
        wb_after.close()

    csv_in = tmp_path / "in.csv"
    csv_in.write_text("Val\n5\n", encoding="utf-8")
    csv_out = tmp_path / "out.csv"
    res_csv = execute_custom_code(str(csv_in), "result = df.assign(Double=df['Val'] * 2)", output_file=str(csv_out))
    assert res_csv["status"] == "success"
    assert "Double" in csv_out.read_text(encoding="utf-8")


def test_custom_code_timeout_handled(tmp_path: Path) -> None:
    fp = tmp_path / "t.xlsx"
    wb = Workbook()
    wb.active.append(["A"])
    wb.save(str(fp))
    wb.close()

    with patch("subprocess.run", side_effect=subprocess.TimeoutExpired(cmd="python", timeout=15)):
        res = execute_custom_code(str(fp), "result = 1")
    assert res["status"] == "error"
    assert "timed out" in res["message"]


def test_phase1_remaining_branches(tmp_path: Path) -> None:
    # _MCPToolWrapper.__getattr__
    dummy_mcp = MagicMock()
    dummy_mcp.custom_attr = "ok"
    assert _MCPToolWrapper(dummy_mcp).custom_attr == "ok"

    # upload_file list/str filename variants
    b64 = base64.b64encode(b"data").decode("ascii")
    res_multi = upload_file(file_content=[b64, b64], filename="single_hint.xlsx")
    assert res_multi["count"] == 2
    res_list_name = upload_file(file_content=b64, filename=["one.xlsx"])
    assert res_list_name["file_path"].endswith(".xlsx")
    res_empty_list_name = upload_file(file_content=b64, filename=[])
    assert res_empty_list_name["file_path"].endswith(".xlsx")

    # expression_validator remaining branches
    with pytest.raises(ValueError, match="Invalid expression syntax"):
        validate_expression("1 +")
    with pytest.raises(ValueError, match="forbidden name"):
        validate_expression("exec")
    with pytest.raises(ValueError, match="Unknown variable"):
        validate_expression("y", allowed_names=frozenset({"x"}))
    with pytest.raises(ValueError, match="Unsafe function call"):
        validate_expression("unknown_fn(1)")
    assert validate_expression("2 ** (x)", allowed_names=frozenset({"x"})) is not None

    # file_resolver remaining SSRF / size / redirect branches
    with pytest.raises(ValueError, match="Invalid HTTP"):
        resolve_to_local("http:///no-host/a.xlsx")
    with patch("socket.getaddrinfo", side_effect=socket.gaierror("nope")):
        with pytest.raises(ValueError, match="Failed to resolve host"):
            resolve_to_local("https://unresolvable.invalid/a.xlsx")

    public_addr = [(socket.AF_INET, socket.SOCK_STREAM, 6, "", ("93.184.216.34", 443))]
    no_loc_resp = MagicMock()
    no_loc_resp.__enter__ = MagicMock(return_value=no_loc_resp)
    no_loc_resp.__exit__ = MagicMock(return_value=False)
    no_loc_resp.status_code = 302
    no_loc_resp.headers = {}
    with patch("socket.getaddrinfo", return_value=public_addr):
        with patch("httpx.stream", return_value=no_loc_resp):
            with pytest.raises(ValueError, match="missing Location"):
                resolve_to_local("https://example.com/a.xlsx")

    loop_resp = MagicMock()
    loop_resp.__enter__ = MagicMock(return_value=loop_resp)
    loop_resp.__exit__ = MagicMock(return_value=False)
    loop_resp.status_code = 302
    loop_resp.headers = {"location": "https://example.com/next.xlsx"}
    with patch("socket.getaddrinfo", return_value=public_addr):
        with patch("httpx.stream", return_value=loop_resp):
            with pytest.raises(ValueError, match="Exceeded maximum redirects"):
                resolve_to_local("https://example.com/a.xlsx")

    big_cl_resp = MagicMock()
    big_cl_resp.__enter__ = MagicMock(return_value=big_cl_resp)
    big_cl_resp.__exit__ = MagicMock(return_value=False)
    big_cl_resp.status_code = 200
    big_cl_resp.headers = {"content-length": str(999_999_999)}
    with patch("socket.getaddrinfo", return_value=public_addr):
        with patch("httpx.stream", return_value=big_cl_resp):
            with pytest.raises(ValueError, match="max allowed size"):
                resolve_to_local("https://example.com/a.xlsx")

    big_stream_resp = MagicMock()
    big_stream_resp.__enter__ = MagicMock(return_value=big_stream_resp)
    big_stream_resp.__exit__ = MagicMock(return_value=False)
    big_stream_resp.status_code = 200
    big_stream_resp.headers = {}
    big_stream_resp.iter_bytes = MagicMock(return_value=iter([b"x" * (51 * 1024 * 1024)]))
    with patch("socket.getaddrinfo", return_value=public_addr):
        with patch("httpx.stream", return_value=big_stream_resp):
            with pytest.raises(ValueError, match="max allowed size"):
                resolve_to_local("https://example.com/a.xlsx")

    # excel_helpers remaining branches
    assert col_letter_to_index("AA") == 27
    new_dest = tmp_path / "sub" / "new_book.xlsx"
    src_book = tmp_path / "src.xlsx"
    wb_src = Workbook()
    wb_src.active.append(["Col"])
    wb_src.active.append([1])
    wb_src.save(str(src_book))
    wb_src.close()
    write_dataframe_to_workbook(
        pd.DataFrame({"Col": [2]}),
        str(src_book),
        sheet_name="Custom",
        output_file=str(new_dest),
        include_header=False,
    )
    assert new_dest.exists()
    wb_chk = load_workbook_safe(str(new_dest))
    try:
        assert get_sheet(wb_chk, "Custom")["A1"].value == 2
        assert load_hidden_json(wb_chk, "Missing") == {}
        with pytest.raises(ValueError, match="not found"):
            get_sheet(wb_chk, "Nope")
    finally:
        wb_chk.close()

    # custom_code remaining branches: >50 rows, Series, ndarray, no-op with print, error file
    assert execute_custom_code(str(tmp_path / "missing.xlsx"), "result = 1")["status"] == "error"
    big_df_res = execute_custom_code(
        str(src_book),
        "result = pd.DataFrame({'A': list(range(55))})",
        output_file=str(tmp_path / "big_out.xlsx"),
    )
    assert big_df_res["status"] == "success"
    assert "preview" in big_df_res["result"]
    assert execute_custom_code(str(src_book), "result = pd.Series([1, 2], name='s')")["status"] == "success"
    assert execute_custom_code(str(src_book), "result = np.array([1, 2])")["status"] == "success"
    noop_res = execute_custom_code(str(src_book), "print('hi')")
    assert noop_res["status"] == "success"
    assert "hi" in noop_res["stdout"]

