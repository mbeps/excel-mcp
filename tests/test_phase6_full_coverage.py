from __future__ import annotations

import subprocess
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook
from PIL import Image

from mcp_server.routes import custom_code as custom_code_routes
from mcp_server.routes import file_transfer as file_transfer_routes
from mcp_server.routes import worksheet_ops as worksheet_ops_routes
from mcp_server.tools import analysis, cell_ops, custom_code, formulas, multi_file, pivot_etl
from mcp_server.utils import excel_helpers, file_resolver


def _make_workbook(path: Path, sheets: dict[str, list[list[object]]]) -> Path:
    wb = Workbook()
    first = True
    for name, rows in sheets.items():
        ws = wb.active if first else wb.create_sheet(name)
        if first:
            assert ws is not None
            ws.title = name
            first = False
        for row in rows:
            ws.append(row)
    wb.save(str(path))
    wb.close()
    return path


def test_routes_custom_code_file_transfer_and_worksheet_ops(tmp_path: Path) -> None:
    wb_path = _make_workbook(tmp_path / "book.xlsx", {"Sheet1": [["A"], [1]]})
    img_path = tmp_path / "dot.png"
    Image.new("RGB", (4, 4), color="red").save(img_path)

    img_res = custom_code_routes.insert_image(str(wb_path), "Sheet1", str(img_path), "B2")
    assert img_res["status"] == "success"

    code_res = custom_code_routes.execute_custom_code(str(wb_path), "result = int(df['A'].sum())")
    assert code_res["status"] == "success"
    assert code_res["result"] == 1

    with pytest.raises(ValueError, match="Either 'file_content' or 'url'"):
        file_transfer_routes.upload_file()

    with pytest.raises(ValueError, match="print_area is required"):
        worksheet_ops_routes.worksheet_print("set_print_area", str(wb_path), "Sheet1", print_area=None)


def test_analysis_non_ascii_col_ref(tmp_path: Path) -> None:
    wb_path = _make_workbook(tmp_path / "a.xlsx", {"Sheet1": [["Name", "Val"], ["x", 1]]})
    with pytest.raises(ValueError, match="Column 'Ä' not found"):
        analysis.vlookup_helper(str(wb_path), str(wb_path), "Ä", "Name", ["Val"])


def test_cell_ops_remaining_branches(tmp_path: Path) -> None:
    wb_path = _make_workbook(
        tmp_path / "cells.xlsx",
        {"Sheet1": [["A", "B", "C"], [True, False, None], [None, None, None]]},
    )
    wb = load_workbook(str(wb_path))
    ws = wb["Sheet1"]
    ws.merge_cells("B3:C3")
    wb.save(str(wb_path))
    wb.close()

    with pytest.raises(ValueError, match="Cannot write to merged cell"):
        cell_ops.copy_range(str(wb_path), "Sheet1", "A1:B1", "Sheet1", "B3")

    res_series = cell_ops.fill_series(
        str(wb_path),
        "Sheet1",
        "A4",
        count=2,
        series_type="custom",
        step="TAG",
        direction="down",
    )
    assert res_series["values_written"] == ["TAG", "TAG"]

    res_replace = cell_ops.find_replace(str(wb_path), "Sheet1", "False", "true", match_entire_cell=True)
    assert res_replace["replacements_made"] == 1
    wb_check = load_workbook(str(wb_path))
    assert wb_check["Sheet1"]["B2"].value is True
    wb_check.close()


def test_custom_code_remaining_branches(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    wb_path = _make_workbook(tmp_path / "cc.xlsx", {"Sheet1": [["A"], [2]]})

    # Fallback str(value) in _serialize_result
    res_tuple = custom_code.execute_custom_code(str(wb_path), "result = {10, 20}")
    assert res_tuple["status"] == "success"
    assert isinstance(res_tuple["result"], str)

    # Missing sheet -> read_sheet_df error
    res_missing_sheet = custom_code.execute_custom_code(str(wb_path), "result = 1", sheet="NoSuchSheet")
    assert res_missing_sheet["status"] == "error"
    assert "Failed to read file" in res_missing_sheet["message"]

    # stdout captured when writing DataFrame back
    res_stdout_df = custom_code.execute_custom_code(
        str(wb_path),
        "print('hello sandbox')\nresult = df",
    )
    assert res_stdout_df["status"] == "success"
    assert "hello sandbox" in res_stdout_df["stdout"]

    # Write failure when output_file has invalid extension
    res_bad_out = custom_code.execute_custom_code(
        str(wb_path),
        "result = df",
        output_file=str(tmp_path / "bad.xyz"),
    )
    assert res_bad_out["status"] == "error"
    assert "Failed to write result" in res_bad_out["message"]

    # Subprocess exits without output_pkl
    with patch.object(
        subprocess,
        "run",
        return_value=SimpleNamespace(returncode=1, stdout="", stderr="boom"),
    ):
        res_no_pkl = custom_code.execute_custom_code(str(wb_path), "result = 1")
    assert res_no_pkl == {"status": "error", "message": "Execution error: boom"}

    # Subprocess raises unexpected OSError
    with patch.object(subprocess, "run", side_effect=OSError("spawn failed")):
        res_os_err = custom_code.execute_custom_code(str(wb_path), "result = 1")
    assert res_os_err == {"status": "error", "message": "Execution error: spawn failed"}


def test_formulas_single_row_range_errors(tmp_path: Path) -> None:
    wb_path = _make_workbook(tmp_path / "err.xlsx", {"Sheet1": [["#DIV/0!", 1, "#REF!"]]})
    res = formulas.get_formula_errors(str(wb_path), "Sheet1", cell_range="1")
    assert res["count"] == 2


def test_multi_file_empty_column_dtype(tmp_path: Path) -> None:
    f1 = _make_workbook(tmp_path / "f1.xlsx", {"Sheet1": [["ID", "Note"], [1, None], [2, None]]})
    f2 = _make_workbook(tmp_path / "f2.xlsx", {"Sheet1": [["ID", "Note"], [1, "ok"], [2, "ok"]]})
    res = multi_file.validate_data_consistency([str(f1), str(f2)], key_column="ID", check_columns=["Note"])
    assert res["schema_mismatches"] == []


def test_pivot_etl_remaining_branches(tmp_path: Path) -> None:
    wb_path = _make_workbook(
        tmp_path / "pivot.xlsx",
        {
            "Sales": [["Cat", "Amt"], ["A", 10], ["B", 20]],
            "Left": [["k", "v"], [1, "x"], [1, "y"]],
            "Right": [["k", "w"], [1, "z"]],
        },
    )
    pivot_etl.create_pivot_table(str(wb_path), "Sales", ["Cat"], ["Amt"], output_sheet="Piv")
    refreshed = pivot_etl.refresh_pivot_table(str(wb_path), "Piv", source_file_path=str(wb_path))
    assert refreshed["rows"] == 2

    with pytest.raises(ValueError, match="Unsupported join type"):
        pivot_etl.merge_datasets(str(wb_path), "Left", "Right", join_key="k", how="cross")

    with pytest.raises(ValueError, match="not found in sheet 'Right'"):
        pivot_etl.merge_datasets(str(wb_path), "Left", "Right", join_key="v")

    merged = pivot_etl.merge_datasets(str(wb_path), "Left", "Right", left_on="k", right_on="k")
    assert merged["row_count"] == 2


def test_excel_helpers_remaining_branches(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    # Directory instead of file
    dir_xlsx = tmp_path / "dir.xlsx"
    dir_xlsx.mkdir()
    with pytest.raises(ValueError, match="Path is not a file"):
        excel_helpers.validate_file_path(str(dir_xlsx), must_exist=True)

    # Corrupt workbook
    bad_xlsx = tmp_path / "corrupt.xlsx"
    bad_xlsx.write_bytes(b"not a zip")
    with pytest.raises(ValueError, match="Failed to open workbook"):
        excel_helpers.load_workbook_safe(str(bad_xlsx))

    # save_workbook_safe unlink OSError in finally
    wb = Workbook()
    orig_unlink = Path.unlink

    def failing_unlink(self: Path, missing_ok: bool = False) -> None:
        if self.suffix == ".tmp":
            raise OSError("cannot unlink")
        orig_unlink(self, missing_ok=missing_ok)

    with (
        patch.object(wb, "save", side_effect=RuntimeError("save boom")),
        patch.object(Path, "unlink", failing_unlink),
        pytest.raises(ValueError, match="Failed to save workbook"),
    ):
        excel_helpers.save_workbook_safe(wb, str(tmp_path / "out.xlsx"))
    wb.close()

    # Empty CSV -> EmptyDataError wrapped in ValueError
    empty_csv = tmp_path / "empty.csv"
    empty_csv.write_text("", encoding="utf-8")
    with pytest.raises(ValueError, match="Failed to read CSV"):
        excel_helpers.read_sheet_df(str(empty_csv), 0)

    # Undecodable CSV in both utf-8 and latin-1
    with patch.object(pd, "read_csv", side_effect=UnicodeDecodeError("utf-8", b"", 0, 1, "bad")):
        with pytest.raises(ValueError, match="Failed to decode CSV"):
            excel_helpers.read_sheet_df(str(empty_csv), 0)

    # .xls failure does not fall back to openpyxl
    bad_xls = tmp_path / "bad.xls"
    bad_xls.write_bytes(b"not xls")
    with pytest.raises(ValueError, match="Failed to read sheet"):
        excel_helpers.read_sheet_df(str(bad_xls), 0)

    # Range validation edge cases
    assert not excel_helpers.validate_excel_range("   ")["valid"]
    assert not excel_helpers.validate_excel_range("A1:B2:C3")["valid"]
    assert not excel_helpers.validate_excel_range("A0")["valid"]
    assert not excel_helpers.validate_excel_range("A1048577")["valid"]

    # Hidden JSON empty and malformed
    wb_json = Workbook()
    ws_hidden = wb_json.create_sheet("_hidden")
    assert excel_helpers.load_hidden_json(wb_json, "_hidden") == {}
    ws_hidden["A1"] = "{bad json"
    assert excel_helpers.load_hidden_json(wb_json, "_hidden") == {}
    wb_json.close()

    # write_dataframe_to_workbook with sheet_name=None on existing file
    existing_wb = _make_workbook(tmp_path / "existing.xlsx", {"Main": [["X"], [1]]})
    excel_helpers.write_dataframe_to_workbook(pd.DataFrame({"X": [99]}), str(existing_wb), sheet_name=None)
    df_back = excel_helpers.read_sheet_df(str(existing_wb), "Main")
    assert df_back["X"].tolist() == [99]


def test_file_resolver_remaining_branches(tmp_path: Path) -> None:
    file_resolver._safe_remove(str(tmp_path / "nonexistent.tmp"))

    store = file_resolver.SessionFileStore(ttl=-1.0)
    assert store.get("missing-session") is None

    temp_file = tmp_path / "expired.xlsx"
    temp_file.write_bytes(b"abc")
    kept_file = tmp_path / "kept.xlsx"
    kept_file.write_bytes(b"def")

    store.register(str(temp_file), is_temp=True)
    # Trigger _cleanup_expired on next register
    store.register(str(kept_file), is_temp=False)
    assert not temp_file.exists()
    # Trigger _cleanup_expired for is_temp=False
    store.register(str(kept_file), is_temp=False)
    assert kept_file.exists()
