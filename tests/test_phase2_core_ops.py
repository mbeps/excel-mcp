"""Phase 2 regression and coverage tests for cell_ops, formulas, workbook, worksheet_ops, and csv_ops."""

from __future__ import annotations

import datetime
from pathlib import Path
from unittest.mock import patch

import openpyxl
import pytest
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.formula import ArrayFormula

from mcp_server.routes import cell_ops as cell_ops_route
from mcp_server.routes import worksheet_ops as ws_ops_route
from mcp_server.tools.cell_ops import (
    auto_sum,
    copy_range,
    fill_formula,
    fill_series,
    find_replace,
    read_cell,
    read_range,
    transpose_range,
    write_cell,
    write_range,
)
from mcp_server.tools.csv_ops import csv_to_xlsx, read_csv_preview
from mcp_server.tools.formulas import (
    get_formula_dependents,
    get_formula_errors,
    get_formula_precedents,
    get_formula_value,
    list_formulas,
    set_formulas_batch,
)
from mcp_server.tools.workbook import get_sheet_summary, set_tab_color
from mcp_server.tools.worksheet_ops import (
    copy_range_across_sheets,
    copy_sheet_across_workbooks,
    delete_cols,
    delete_rows,
    group_cols,
    group_rows,
    insert_cols,
    insert_rows,
    merge_workbooks,
    remove_page_break,
    set_col_width,
    stack_sheets,
    ungroup_cols,
    ungroup_rows,
)


def _make_wb(tmp_path: Path, name: str = "test.xlsx", setup=None) -> str:
    path = str(tmp_path / name)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    if setup is not None:
        setup(wb, ws)
    wb.save(path)
    wb.close()
    return path


# ---------------------------------------------------------------------------
# cell_ops
# ---------------------------------------------------------------------------


def test_read_range_style_and_array_formula_and_bounds(tmp_path: Path) -> None:
    def setup(_wb, ws):
        ws["A1"] = "Header"
        ws["A1"].font = Font(bold=True, color="FF112233")
        ws["B1"] = ArrayFormula("B1:B2", "=A1:A2")
        ws["A2"] = 10
        ws["B2"] = 20

    path = _make_wb(tmp_path, setup=setup)

    res = read_range(path, "Sheet1", "A1", "B2", show_formula=False, show_style=True)
    assert res["styles"][0][0]["bold"] is True
    assert res["rows"][0][1] is None

    res_formula = read_range(path, "Sheet1", "A1", "B2", show_formula=True, show_style=True)
    assert res_formula["styles"][0][0]["bold"] is True
    assert res_formula["rows"][0][1] == "=A1:A2"

    # Reversed range rejected
    with pytest.raises(ValueError, match="precedes"):
        read_range(path, "Sheet1", "B2", "A1")

    # max_cells < 0 rejected
    with pytest.raises(ValueError, match="max_cells"):
        read_range(path, "Sheet1", "A1", "B2", max_cells=-1)

    # max_cells == 0 returns empty rows
    zero = read_range(path, "Sheet1", "A1", "B2", max_cells=0)
    assert zero["rows"] == []
    assert zero["truncated"] is True

    # max_cells < col_count_available obeys strict cap
    capped = read_range(path, "Sheet1", "A1", "B2", max_cells=1)
    assert capped["row_count"] == 1
    assert capped["col_count"] == 1
    assert capped["truncated"] is True


def test_merged_cell_write_guards(tmp_path: Path) -> None:
    def setup(_wb, ws):
        ws["A1"] = "Merged"
        ws.merge_cells("A1:B2")
        ws["C1"] = "=A1"

    path = _make_wb(tmp_path, setup=setup)

    with pytest.raises(ValueError, match="merged cell"):
        write_cell(path, "Sheet1", "B1", "x")
    with pytest.raises(ValueError, match="merged cell"):
        write_range(path, "Sheet1", "B1", [["x"]])
    with pytest.raises(ValueError, match="merged cell"):
        fill_formula(path, "Sheet1", "C1", "B1:B2")
    with pytest.raises(ValueError, match="merged cell"):
        transpose_range(path, "Sheet1", "C1:C1", "B1")


def test_copy_range_overlapping_and_formula_translation(tmp_path: Path) -> None:
    def setup(_wb, ws):
        ws["A1"] = 10
        ws["B1"] = 20
        ws["A2"] = "=A1*2"
        ws["B2"] = "=B1*2"
        ws["A1"].font = Font(bold=True)

    path = _make_wb(tmp_path, setup=setup)
    copy_range(path, "Sheet1", "A1:B2", "Sheet1", "B1", copy_values=True, copy_styles=True)

    wb = openpyxl.load_workbook(path)
    ws = wb["Sheet1"]
    assert ws["B1"].value == 10
    assert ws["C1"].value == 20
    assert ws["B2"].value == "=B1*2"
    assert ws["C2"].value == "=C1*2"
    assert ws["B1"].font.bold is True
    wb.close()


def test_fill_series_validations_and_date_default_step(tmp_path: Path) -> None:
    def setup(_wb, ws):
        ws["A1"] = datetime.datetime(2025, 1, 1)
        ws["B1"] = True

    path = _make_wb(tmp_path, setup=setup)

    # Date series with default integer step=1
    res = fill_series(path, "Sheet1", "A1", "date", count=3, step=1)
    assert res["values_written"] == ["2025-01-01", "2025-01-02", "2025-01-03"]

    with pytest.raises(ValueError, match="count must be >= 1"):
        fill_series(path, "Sheet1", "A1", "number", count=0)
    with pytest.raises(ValueError, match="Boolean"):
        fill_series(path, "Sheet1", "B1", "number", count=2)
    with pytest.raises(ValueError, match="Boolean"):
        fill_series(path, "Sheet1", "A1", "date", count=2, step=True)
    with pytest.raises(ValueError, match="non-empty"):
        fill_series(path, "Sheet1", "A1", "custom", count=2, step="")


def test_find_replace_types_backslashes_and_array_formula(tmp_path: Path) -> None:
    def setup(_wb, ws):
        ws["A1"] = 100
        ws["A2"] = 3.5
        ws["A3"] = True
        ws["A4"] = "path_here"
        ws["A5"] = ArrayFormula("A5:A5", "=SUM(A1:A2)")

    path = _make_wb(tmp_path, setup=setup)

    find_replace(path, "Sheet1", "100", "200")
    find_replace(path, "Sheet1", "3.5", "4.25")
    find_replace(path, "Sheet1", "True", "false")
    find_replace(path, "Sheet1", "here", r"\1\path", regex=False)
    # ArrayFormula skipped when search_formulas=False, updated when True
    skipped = find_replace(path, "Sheet1", "SUM", "AVERAGE", search_formulas=False)
    assert skipped["replacements_made"] == 0
    updated = find_replace(path, "Sheet1", "SUM", "AVERAGE", search_formulas=True)
    assert updated["replacements_made"] == 1

    wb = openpyxl.load_workbook(path)
    ws = wb["Sheet1"]
    assert ws["A1"].value == 200 and type(ws["A1"].value) is int
    assert ws["A2"].value == 4.25 and type(ws["A2"].value) is float
    assert ws["A3"].value is False
    assert ws["A4"].value == r"path_\1\path"
    assert isinstance(ws["A5"].value, ArrayFormula)
    assert ws["A5"].value.text == "=AVERAGE(A1:A2)"
    wb.close()


def test_auto_sum_excludes_bool_and_includes_formulas(tmp_path: Path) -> None:
    def setup(_wb, ws):
        ws["A1"] = True
        ws["A2"] = 10
        ws["A3"] = "=5+5"
        ws["B4"] = False
        ws["C4"] = 20
        ws["D4"] = "=10+10"

    path = _make_wb(tmp_path, setup=setup)
    up = auto_sum(path, "Sheet1", "A4")
    assert up["formula"] == "=SUM(A2:A3)"
    left = auto_sum(path, "Sheet1", "E4")
    assert left["formula"] == "=SUM(C4:D4)"

    # read_cell with ArrayFormula
    def setup_arr(_wb, ws):
        ws["A1"] = ArrayFormula("A1:A1", "=1+1")

    arr_path = _make_wb(tmp_path, name="arr.xlsx", setup=setup_arr)
    rc = read_cell(arr_path, "Sheet1", "A1", include_formula=True)
    assert rc["formula"] == "=1+1"

    # Route datetime value
    dt = datetime.date(2025, 5, 1)
    cell_ops_route.write_cells("single", path, "Sheet1", cell_ref="F1", value=dt)
    assert "F1" in cell_ops_route.read_cells("single", path, "Sheet1", cell_ref="F1")["cell_ref"]


# ---------------------------------------------------------------------------
# formulas
# ---------------------------------------------------------------------------


def test_formulas_phase2_fixes(tmp_path: Path) -> None:
    def setup(wb, ws):
        ws["A1"] = 10
        ws["B2"] = 20
        ws["C3"] = "#CALC!"
        ws["C4"] = "#BLOCKED!"
        ws["D1"] = ArrayFormula("D1:D1", "=SUM(A1:B2)")
        ws["E1"] = '=LOG10(B2) + ATAN2(A1, B2) + SUM(B2:A1) + "A99" + \'Bob\'\'s Sheet\'!C5'
        ws["F1"] = '="B2" + Sheet1!B2'
        ws2 = wb.create_sheet("Sheet2")
        ws2["B2"] = 99
        ws["G1"] = "=Sheet2!B2"

    path = _make_wb(tmp_path, setup=setup)

    set_formulas_batch(path, "Sheet1", {"H1": "{=SUM(A1:A1)}"})
    assert get_formula_value(path, "Sheet1", "H1")["formula"] == "=SUM(A1:A1)"

    errs = get_formula_errors(path, "Sheet1")
    err_codes = {e["error"] for e in errs["errors"]}
    assert {"#CALC!", "#BLOCKED!"} <= err_codes

    prec = get_formula_precedents(path, "Sheet1", "E1")["precedents"]
    assert "A99" not in prec
    assert "LOG10" not in prec
    assert "ATAN2" not in prec
    assert "'Bob''s Sheet'!C5" in prec
    assert "A1" in prec and "B2" in prec

    deps = get_formula_dependents(path, "Sheet1", "B2")
    dep_cells = {d["cell"] for d in deps["dependents"]}
    assert "F1" in dep_cells
    assert "E1" in dep_cells
    assert "G1" not in dep_cells  # Sheet2!B2 must not match Sheet1!B2

    listed = {item["cell_ref"]: item["formula"] for item in list_formulas(path, "Sheet1")}
    assert listed["D1"] == "=SUM(A1:B2)"

    # Max expansion cap
    set_formulas_batch(path, "Sheet1", {"H2": "=SUM(A1:Z10000)"})
    with pytest.raises(ValueError, match="maximum precedent expansion"):
        get_formula_precedents(path, "Sheet1", "H2")


# ---------------------------------------------------------------------------
# workbook, worksheet_ops, csv_ops
# ---------------------------------------------------------------------------


def test_workbook_summary_fallback_and_tab_color(tmp_path: Path) -> None:
    def setup(_wb, ws):
        ws["B3"] = "Col1"
        ws["C3"] = "Col2"
        ws["B4"] = 1
        ws["C4"] = 2

    path = _make_wb(tmp_path, setup=setup)
    with patch("mcp_server.tools.workbook.read_sheet_df", side_effect=RuntimeError("boom")):
        summary = get_sheet_summary(path, "Sheet1")
    assert summary.row_count == 2
    assert summary.col_count == 2
    assert summary.used_range == "B3:C4"

    # #000000 is allowed (black tab), None/"" clears
    res_black = set_tab_color(path, "Sheet1", "#000000")
    assert res_black["color"] == "000000"
    wb = openpyxl.load_workbook(path)
    assert wb["Sheet1"].sheet_properties.tabColor is not None
    wb.close()

    res_clear = set_tab_color(path, "Sheet1", "")
    assert res_clear["color"] is None
    wb = openpyxl.load_workbook(path)
    assert wb["Sheet1"].sheet_properties.tabColor is None
    wb.close()


def test_worksheet_ops_phase2_fixes(tmp_path: Path) -> None:
    def setup(wb, ws):
        ws["A1"] = 1
        ws["B1"] = 2
        ws["A1"].font = Font(bold=True)
        long_ws = wb.create_sheet("VeryLongSheetNameThatIs29Char")
        long_ws["A1"] = "hi"

    path1 = _make_wb(tmp_path, name="wb1.xlsx", setup=setup)
    path2 = _make_wb(tmp_path, name="wb2.xlsx", setup=setup)

    # Single-cell copy_range_across_sheets + overlapping copy
    copy_range_across_sheets(path1, "Sheet1", "A1", "Sheet1", "B1", copy_styles=True)
    wb = openpyxl.load_workbook(path1)
    assert wb["Sheet1"]["B1"].value == 1
    assert wb["Sheet1"]["B1"].font.bold is True
    wb.close()

    # copy_sheet_across_workbooks & merge_workbooks
    dest_sheet_wb = str(tmp_path / "copied.xlsx")
    copy_sheet_across_workbooks(path1, "Sheet1", dest_sheet_wb)
    assert Path(dest_sheet_wb).exists()

    empty_merged = str(tmp_path / "empty_merged.xlsx")
    res_empty = merge_workbooks([], empty_merged)
    assert res_empty["total_sheets"] == 1

    merged = str(tmp_path / "merged.xlsx")
    res_m = merge_workbooks([path1, path2], merged, conflict_strategy="rename")
    assert all(len(s) <= 31 for s in res_m["sheets"])

    # Row/col bounds and route falsiness
    for fn in (insert_rows, delete_rows):
        with pytest.raises(ValueError, match=">= 1"):
            fn(path1, "Sheet1", 0, 1)
        with pytest.raises(ValueError, match=">= 1"):
            fn(path1, "Sheet1", 1, 0)
    for fn in (insert_cols, delete_cols):
        with pytest.raises(ValueError, match=">= 1"):
            fn(path1, "Sheet1", 0, 1)
        with pytest.raises(ValueError, match=">= 1"):
            fn(path1, "Sheet1", 1, 0)
    with pytest.raises(ValueError, match=">= 1"):
        ws_ops_route.worksheet_structure("insert_rows", path1, "Sheet1", row=1, count=0)

    # ungroup clears collapsed
    group_rows(path1, "Sheet1", 2, 3, hidden=True)
    ungroup_rows(path1, "Sheet1", 2, 3)
    group_cols(path1, "Sheet1", 2, 3, hidden=True)
    ungroup_cols(path1, "Sheet1", 2, 3)
    wb = openpyxl.load_workbook(path1)
    assert wb["Sheet1"].row_dimensions[4].collapsed is False
    assert wb["Sheet1"].column_dimensions["D"].collapsed is False
    wb.close()

    # set_col_width lowercase
    set_col_width(path1, "Sheet1", ["a", "b"], 25.0)
    wb = openpyxl.load_workbook(path1)
    assert wb["Sheet1"].column_dimensions["A"].width == 25.0
    wb.close()

    # stack_sheets empty & NaN sanitization
    with pytest.raises(ValueError, match="must not be empty"):
        stack_sheets(path1, [], "Out")

    def setup_nan(wb, ws):
        ws.append(["Name", "Val"])
        ws.append(["Alice", 10])
        ws.append(["Bob", None])
        ws2 = wb.create_sheet("Sheet2")
        ws2.append(["Name", "Val"])
        ws2.append(["Cara", 20])

    nan_path = _make_wb(tmp_path, name="nan.xlsx", setup=setup_nan)
    stack_out = str(tmp_path / "stacked.xlsx")
    stack_sheets(nan_path, ["Sheet1", "Sheet2"], "Combined", output_path=stack_out)
    wb = openpyxl.load_workbook(stack_out)
    assert wb["Combined"]["B3"].value is None
    wb.close()

    # remove_page_break validation
    with pytest.raises(ValueError, match=">= 1"):
        remove_page_break(path1, "Sheet1", row=0)
    with pytest.raises(ValueError, match=">= 1"):
        remove_page_break(path1, "Sheet1", col=0)


def test_csv_ops_nan_and_latin1_fallback(tmp_path: Path) -> None:
    csv_path = tmp_path / "data.csv"
    csv_path.write_bytes("name,val\nCaf\xe9,\n".encode("latin-1"))

    prev = read_csv_preview(str(csv_path), encoding="utf-8")
    assert prev["rows"] == [["Café", None]]

    xlsx_out = tmp_path / "out.xlsx"
    msg = csv_to_xlsx(str(csv_path), str(xlsx_out), encoding="utf-8")
    assert "Converted" in msg
    wb = openpyxl.load_workbook(xlsx_out)
    assert wb["Sheet1"]["A2"].value == "Café"
    wb.close()


def test_phase2_remaining_edge_cases(tmp_path: Path) -> None:
    """Cover remaining error/edge branches in Phase 2 tools and routes."""
    from openpyxl.styles import PatternFill

    from mcp_server.routes import formulas as formulas_route
    from mcp_server.routes import workbook as workbook_route
    from mcp_server.tools.cell_ops import _coerce_replaced_value, clear_range
    from mcp_server.tools.formulas import _cell_in_range
    from mcp_server.tools.worksheet_ops import set_row_height

    def setup(wb, ws):
        ws["A1"] = "hello"
        ws["A1"].fill = PatternFill(start_color="FF00FF00", end_color="FF00FF00", fill_type="solid")
        ws["A2"] = 10
        ws["A3"] = 20
        ws["B1"] = "=SUM(A2:A4)"
        ws["B2"] = "=SUM(Sheet2!A1:A2)"
        ws["C1"] = "#CALC!"
        ws.column_dimensions["A"].width = 18.0
        ws2 = wb.create_sheet("Sheet2")
        ws2["A1"] = 5
        ws2["A2"] = 6

    path = _make_wb(tmp_path, name="edges.xlsx", setup=setup)

    # cell_ops edges
    assert read_cell(path, "Sheet1", "A1", include_formula=True)["formula"] is None
    r_cap = read_range(path, "Sheet1", "A1", "B3", show_style=True, max_cells=4)
    assert r_cap["row_count"] == 2 and r_cap["truncated"] is True
    with pytest.raises(ValueError, match="rectangular"):
        write_range(path, "Sheet1", "D1", [[1, 2], [3]])
    with pytest.raises(ValueError):
        clear_range(path, "Sheet1", "B2", "A1")
    with pytest.raises(ValueError):
        copy_range(path, "Sheet1", "B2:A1", "Sheet1", "D1")
    with pytest.raises(ValueError):
        copy_range(path, "Sheet1", "A1:A1", "Sheet1", "ZZZ1")
    copy_range(path, "Sheet1", "B1:B1", "Sheet1", "D1", paste_values_only=True)
    copy_range(path, "Sheet1", "A1:A1", "Sheet1", "D2", copy_values=False, copy_styles=True)
    with pytest.raises(ValueError, match="direction"):
        fill_series(path, "Sheet1", "A1", "number", count=2, direction="up")
    fill_series(path, "Sheet1", "E10", "number", count=2)
    fill_series(path, "Sheet1", "E1", "date", count=2, start_value=datetime.datetime(2025, 1, 1), direction="right")
    fill_series(path, "Sheet1", "E2", "text_increment", count=2, start_value="Item1", direction="right")
    with pytest.raises(ValueError, match="non-empty"):
        fill_series(path, "Sheet1", "A1", "custom", count=2, step=None)  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="date value"):
        fill_series(path, "Sheet1", "D5", "date", count=2)
    with pytest.raises(ValueError, match="text value"):
        fill_series(path, "Sheet1", "D5", "text_increment", count=2)
    with pytest.raises(ValueError, match="trailing number"):
        fill_series(path, "Sheet1", "A1", "text_increment", count=2)
    with pytest.raises(ValueError, match="Invalid series_type"):
        fill_series(path, "Sheet1", "A1", "unknown", count=2)
    fill_series(path, "Sheet1", "D1", "custom", count=2, step="X", direction="right")

    assert _coerce_replaced_value(True, "notabool") == "notabool"
    assert _coerce_replaced_value(10, "inf") == "inf"
    assert _coerce_replaced_value(10, "3.5") == 3.5
    assert _coerce_replaced_value(10, "abc") == "abc"
    assert _coerce_replaced_value(1.5, "abc") == "abc"
    assert find_replace(path, "Sheet1", "^hello$", "world", match_entire_cell=True, regex=True)["replacements_made"] == 1
    assert find_replace(path, "Sheet1", "world", "hello", match_entire_cell=True, regex=False)["replacements_made"] == 1

    # formulas edges
    assert get_formula_errors(path, "Sheet1", cell_range="C1")["count"] == 1
    assert get_formula_errors(path, "Sheet1", cell_range="C1:D1")["count"] == 1
    assert "Sheet2!A1" in get_formula_precedents(path, "Sheet1", "B2")["precedents"]
    assert _cell_in_range("A1", "bad") is False
    assert get_formula_dependents(path, "Sheet1", "A3")["count"] >= 1

    # worksheet_ops edges
    with pytest.raises(ValueError):
        copy_range_across_sheets(path, "Sheet1", "B2:A1", "Sheet2", "A1")
    with pytest.raises(ValueError):
        copy_range_across_sheets(path, "Sheet1", "A1", "Sheet2", "ZZZ1")
    with pytest.raises(ValueError, match="conflict_strategy"):
        merge_workbooks([path], str(tmp_path / "m.xlsx"), conflict_strategy="skip")
    existing_m = str(tmp_path / "existing_m.xlsx")
    merge_workbooks([path], existing_m)
    res_m2 = merge_workbooks([path, path], existing_m, conflict_strategy="rename")
    assert "Sheet1_2" in res_m2["sheets"] and "Sheet1_3" in res_m2["sheets"]
    with pytest.raises(ValueError, match=">= 1"):
        set_row_height(path, "Sheet1", [0], 20.0)
    stack_sheets(path, ["Sheet1"], "Stacked", output_path=existing_m)

    # routes missing-arg validation
    for fn, args in [
        (cell_ops_route.read_cells, ("range", path, "Sheet1")),
        (cell_ops_route.write_cells, ("range", path, "Sheet1")),
        (cell_ops_route.write_cells, ("series", path, "Sheet1")),
        (cell_ops_route.write_cells, ("merge", path, "Sheet1")),
        (cell_ops_route.write_cells, ("unmerge", path, "Sheet1")),
        (formulas_route.formula_write, ("set", path, "Sheet1")),
        (formulas_route.formula_write, ("batch", path, "Sheet1")),
        (formulas_route.formula_write, ("fill", path, "Sheet1")),
        (formulas_route.formula_write, ("auto_sum", path, "Sheet1")),
        (formulas_route.formula_write, ("unknown", path, "Sheet1")),  # type: ignore[arg-type]
        (formulas_route.formula_audit, ("value", path, "Sheet1")),
        (formulas_route.formula_audit, ("precedents", path, "Sheet1")),
        (formulas_route.formula_audit, ("dependents", path, "Sheet1")),
        (workbook_route.sheet_management, ("rename", path, "Sheet1")),
        (workbook_route.sheet_management, ("copy", path, "Sheet1")),
        (workbook_route.sheet_management, ("tab_color", path, "Sheet1")),
        (workbook_route.sheet_management, ("move", path, "Sheet1")),
        (ws_ops_route.worksheet_structure, ("insert_rows", path, "Sheet1")),
        (ws_ops_route.worksheet_structure, ("unknown", path, "Sheet1")),  # type: ignore[arg-type]
        (ws_ops_route.worksheet_print, ("unknown", path, "Sheet1")),  # type: ignore[arg-type]
        (ws_ops_route.worksheet_transfer, ("unknown",)),  # type: ignore[arg-type]
    ]:
        with pytest.raises(ValueError):
            fn(*args)
    wb_route_path = str(tmp_path / "wb_route.xlsx")
    workbook_route.create_workbook(wb_route_path, sheet_names=["S1"])
    assert workbook_route.get_sheet_summary(wb_route_path, "S1").name == "S1"
    workbook_route.write_multi_sheet(str(tmp_path / "wb_multi.xlsx"), [{"name": "M1", "headers": ["A"], "data": [[1]]}])
    ws_ops_route.worksheet_print("remove_page_break", path, "Sheet1", row=2)
    with pytest.raises(ValueError):
        cell_ops_route.read_cells("range", path, "Sheet1", start_cell="A1")
    with pytest.raises(ValueError):
        cell_ops_route.write_cells("range", path, "Sheet1", start_cell="A1")
    with pytest.raises(ValueError):
        cell_ops_route.write_cells("series", path, "Sheet1", start_cell="A1")
    with pytest.raises(ValueError):
        formulas_route.formula_write("fill", path, "Sheet1", cell_ref="A1")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_view("auto_filter", path, "Sheet1", remove=False)
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("delete_rows", path, "Sheet1")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("insert_cols", path, "Sheet1")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("delete_cols", path, "Sheet1")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("set_row_height", path, "Sheet1")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("set_row_height", path, "Sheet1", rows=[1])
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("set_col_width", path, "Sheet1")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("set_col_width", path, "Sheet1", cols_list=["A"])
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("group_rows", path, "Sheet1")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("group_rows", path, "Sheet1", start_row=1)
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("group_cols", path, "Sheet1")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("group_cols", path, "Sheet1", start_col=1)
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("ungroup_rows", path, "Sheet1")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("ungroup_rows", path, "Sheet1", start_row=1)
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("ungroup_cols", path, "Sheet1")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_structure("ungroup_cols", path, "Sheet1", start_col=1)
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_transfer("copy_range_across")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_transfer("copy_range_across", file_path=path)
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_transfer("copy_range_across", file_path=path, source_sheet="Sheet1")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_transfer("copy_range_across", file_path=path, source_sheet="Sheet1", source_range="A1")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_transfer("copy_sheet_across")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_transfer("copy_sheet_across", source_file=path)
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_transfer("copy_sheet_across", source_file=path, source_sheet="Sheet1")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_transfer("merge_workbooks")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_transfer("merge_workbooks", source_files=[path])
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_transfer("stack_sheets")
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_transfer("stack_sheets", file_path=path)
    with pytest.raises(ValueError):
        ws_ops_route.worksheet_transfer("stack_sheets", file_path=path, sheet_names=["Sheet1"])
