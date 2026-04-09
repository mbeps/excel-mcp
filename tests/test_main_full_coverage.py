import json
import os
from pathlib import Path

import pytest

os.environ["MCP_SERVER_DISABLE_TOOL_REGISTRATION"] = "1"
import mcp_server.main as main


def test_resources_and_metadata(sample_xlsx):
    sheet_list = json.loads(main.resource_list_sheets(sample_xlsx))
    assert any(sheet.get("name") == "Sheet1" for sheet in sheet_list)

    preview = json.loads(main.resource_sheet_preview(sample_xlsx, "Sheet1"))
    assert "rows" in preview and len(preview["rows"]) > 0

    meta = main.get_workbook_metadata(sample_xlsx)
    assert meta.active_sheet == "Sheet1"


def test_sheet_management_paths(monkeypatch):
    monkeypatch.setattr(main._workbook, "rename_sheet", lambda file_path, sheet_name, new_name: "RENAMED")
    monkeypatch.setattr(main._workbook, "delete_sheet", lambda file_path, sheet_name: "DELETED")
    monkeypatch.setattr(main._workbook, "copy_sheet", lambda file_path, sheet_name, new_name: "COPIED")
    monkeypatch.setattr(main._workbook, "hide_sheet", lambda file_path, sheet_name: "HIDDEN")
    monkeypatch.setattr(main._workbook, "unhide_sheet", lambda file_path, sheet_name: "UNHIDDEN")
    monkeypatch.setattr(main._workbook, "set_tab_color", lambda file_path, sheet_name, color: "TABCOLOR")
    monkeypatch.setattr(main._workbook, "move_sheet", lambda file_path, sheet_name, offset: "MOVED")

    assert main.sheet_management("rename", "foo.xlsx", "Sheet1", new_name="SheetX") == "RENAMED"
    assert main.sheet_management("delete", "foo.xlsx", "Sheet1") == "DELETED"
    assert main.sheet_management("copy", "foo.xlsx", "Sheet1", new_name="SheetX") == "COPIED"
    assert main.sheet_management("hide", "foo.xlsx", "Sheet1") == "HIDDEN"
    assert main.sheet_management("unhide", "foo.xlsx", "Sheet1") == "UNHIDDEN"
    assert main.sheet_management("tab_color", "foo.xlsx", "Sheet1", color="00FF00") == "TABCOLOR"
    assert main.sheet_management("move", "foo.xlsx", "Sheet1", offset=1) == "MOVED"

    with pytest.raises(ValueError, match="new_name is required for action='rename'"):
        main.sheet_management("rename", "foo.xlsx", "Sheet1")
    with pytest.raises(ValueError, match="offset is required for action='move'"):
        main.sheet_management("move", "foo.xlsx", "Sheet1")
    with pytest.raises(ValueError, match="Unknown action"):
        main.sheet_management("unknown", "foo.xlsx", "Sheet1")


def test_read_cells_actions(monkeypatch):
    monkeypatch.setattr(main._cell_ops, "read_cell", lambda *args, **kwargs: "CELL")
    monkeypatch.setattr(main._cell_ops, "read_range", lambda *args, **kwargs: "RANGE")
    monkeypatch.setattr(main._cell_ops, "read_file_chunked", lambda *args, **kwargs: "CHUNKED")

    assert main.read_cells("single", "foo.xlsx", "Sheet1", cell_ref="A1") == "CELL"
    assert main.read_cells("range", "foo.xlsx", "Sheet1", start_cell="A1", end_cell="A2") == "RANGE"
    assert main.read_cells("chunked", "foo.xlsx", "Sheet1") == "CHUNKED"

    with pytest.raises(ValueError, match="cell_ref is required for mode='single'"):
        main.read_cells("single", "foo.xlsx", "Sheet1")
    with pytest.raises(ValueError, match="Unknown mode"):
        main.read_cells("bad", "foo.xlsx", "Sheet1")


def test_write_cells_actions(monkeypatch):
    monkeypatch.setattr(main._cell_ops, "write_cell", lambda *args, **kwargs: "W-CELL")
    monkeypatch.setattr(main._cell_ops, "write_range", lambda *args, **kwargs: "W-RANGE")
    monkeypatch.setattr(main._cell_ops, "fill_series", lambda *args, **kwargs: "W-SERIES")
    monkeypatch.setattr(main._cell_ops, "merge_cells", lambda *args, **kwargs: "W-MERGE")
    monkeypatch.setattr(main._cell_ops, "unmerge_cells", lambda *args, **kwargs: "W-UNMERGE")

    assert main.write_cells("single", "foo.xlsx", "Sheet1", cell_ref="A1", value=123) == "W-CELL"
    assert main.write_cells("range", "foo.xlsx", "Sheet1", start_cell="A1", data=[[1, 2]]) == "W-RANGE"
    assert main.write_cells("series", "foo.xlsx", "Sheet1", start_cell="A1", count=2) == "W-SERIES"
    assert main.write_cells("merge", "foo.xlsx", "Sheet1", range_string="A1:B1") == "W-MERGE"
    assert main.write_cells("unmerge", "foo.xlsx", "Sheet1", range_string="A1:B1") == "W-UNMERGE"

    with pytest.raises(ValueError, match="cell_ref is required for mode='single'"):
        main.write_cells("single", "foo.xlsx", "Sheet1")

    with pytest.raises(ValueError, match="Unknown mode"):
        main.write_cells("bad", "foo.xlsx", "Sheet1")


def test_misc_standalone_ops(monkeypatch):
    monkeypatch.setattr(main._cell_ops, "clear_range", lambda *args, **kwargs: "CLEAR")
    monkeypatch.setattr(main._cell_ops, "copy_range", lambda *args, **kwargs: "COPY")
    monkeypatch.setattr(main._cell_ops, "find_replace", lambda *args, **kwargs: {"count": 1})
    monkeypatch.setattr(main._cell_ops, "transpose_range", lambda *args, **kwargs: {"status": "ok"})

    assert main.clear_range("foo.xlsx", "Sheet1", "A1", "B2") == "CLEAR"
    assert main.copy_range("foo.xlsx", "Sheet1", "A1:B2", "Sheet1", "C1:D2") == "COPY"
    assert main.find_replace("foo.xlsx", "Sheet1", "a", "b")["count"] == 1
    assert main.transpose_range("foo.xlsx", "Sheet1", "A1:B2", "C1")["status"] == "ok"


def test_formatting_wrappers(monkeypatch):
    monkeypatch.setattr(main._formatting, "format_cells", lambda *args, **kwargs: "FMT")
    monkeypatch.setattr(main._formatting, "auto_fit_columns", lambda *args, **kwargs: "AF")
    monkeypatch.setattr(main._formatting, "copy_cell_format", lambda *args, **kwargs: {"ok": True})
    monkeypatch.setattr(main._formatting, "clear_cell_format", lambda *args, **kwargs: {"ok": True})
    monkeypatch.setattr(main._formatting, "apply_named_style", lambda *args, **kwargs: {"ok": True})

    assert main.format_cells("foo.xlsx", "Sheet1", "A1").startswith("FMT")
    assert main.auto_fit_columns("foo.xlsx", "Sheet1") == "AF"
    assert main.copy_cell_format("foo.xlsx", "Sheet1", "A1", "A1")["ok"]
    assert main.clear_cell_format("foo.xlsx", "Sheet1", "A1")["ok"]
    assert main.apply_named_style("foo.xlsx", "Sheet1", "A1", "Good")["ok"]


def test_formula_write_and_audit(monkeypatch):
    monkeypatch.setattr(main._formulas, "set_formula", lambda *args, **kwargs: "FSET")
    monkeypatch.setattr(main._formulas, "set_formulas_batch", lambda *args, **kwargs: "FBATCH")
    monkeypatch.setattr(main._cell_ops, "fill_formula", lambda *args, **kwargs: {"filled": True})
    monkeypatch.setattr(main._cell_ops, "auto_sum", lambda *args, **kwargs: {"sum": 5})
    monkeypatch.setattr(main._formulas, "get_formula_value", lambda *args, **kwargs: "FVAL")
    monkeypatch.setattr(main._formulas, "get_formula_errors", lambda *args, **kwargs: [])
    monkeypatch.setattr(main._formulas, "get_formula_precedents", lambda *args, **kwargs: ["A1"])
    monkeypatch.setattr(main._formulas, "get_formula_dependents", lambda *args, **kwargs: ["B1"])
    monkeypatch.setattr(main._formulas, "list_formulas", lambda *args, **kwargs: ["=A1+B1"])

    assert main.formula_write("set", "foo.xlsx", "Sheet1", cell_ref="A1", formula="=1+1") == "FSET"
    assert main.formula_write("batch", "foo.xlsx", "Sheet1", formulas={"A1": "=1"}) == "FBATCH"
    assert main.formula_write("fill", "foo.xlsx", "Sheet1", cell_ref="A1", target_range="A1:A2")["filled"]
    assert main.formula_write("auto_sum", "foo.xlsx", "Sheet1", cell_ref="B1")["sum"] == 5

    assert main.formula_audit("value", "foo.xlsx", "Sheet1", cell_ref="A1") == "FVAL"
    assert main.formula_audit("errors", "foo.xlsx", "Sheet1") == []
    assert main.formula_audit("precedents", "foo.xlsx", "Sheet1", cell_ref="A1") == ["A1"]
    assert main.formula_audit("dependents", "foo.xlsx", "Sheet1", cell_ref="A1") == ["B1"]
    assert main.formula_audit("list", "foo.xlsx", "Sheet1") == ["=A1+B1"]

    with pytest.raises(ValueError, match="Unknown action"):
        main.formula_audit("x", "foo.xlsx", "Sheet1")


def test_csv_conditional_table(monkeypatch, sample_csv):
    monkeypatch.setattr(main._csv_ops, "read_csv_preview", lambda *args, **kwargs: {"rows": []})
    monkeypatch.setattr(main._csv_ops, "csv_to_xlsx", lambda *args, **kwargs: "C2X")
    monkeypatch.setattr(main._csv_ops, "xlsx_to_csv", lambda *args, **kwargs: "X2C")

    assert main.csv_ops("preview", file_path=sample_csv)["rows"] == []
    assert main.csv_ops("to_xlsx", csv_path=sample_csv, xlsx_path="out.xlsx") == "C2X"
    assert main.csv_ops("to_csv", file_path=sample_csv, output_path="out.csv") == "X2C"

    with pytest.raises(ValueError, match="Unknown action"):
        main.csv_ops("bad")

    # Conditional formatting actions execute internal calls and return JSON when required
    monkeypatch.setattr(main._cond_fmt, "apply_conditional_formatting", lambda *args, **kwargs: "APPLY")
    monkeypatch.setattr(main._cond_fmt, "add_highlight_rule", lambda *args, **kwargs: "HIGHLIGHT")
    monkeypatch.setattr(main._cond_fmt, "add_formula_rule", lambda *args, **kwargs: "FORMULA")
    monkeypatch.setattr(main._cond_fmt, "remove_conditional_formatting", lambda *args, **kwargs: "REMOVED")
    monkeypatch.setattr(main._cond_fmt, "add_top_bottom_rule", lambda *args, **kwargs: {"ok": True})
    monkeypatch.setattr(main._cond_fmt, "add_above_below_average_rule", lambda *args, **kwargs: {"ok": True})

    assert (
        main.conditional_format("apply", "foo.xlsx", "Sheet1", cell_range="A1:A2", format_type="2_color_scale")
        == "APPLY"
    )
    assert (
        main.conditional_format(
            "highlight", "foo.xlsx", "Sheet1", cell_range="A1:A2", operator="greaterThan", formula="A1>1"
        )
        == "HIGHLIGHT"
    )
    assert (
        main.conditional_format("formula_rule", "foo.xlsx", "Sheet1", cell_range="A1:A2", formula="A1>1") == "FORMULA"
    )
    assert main.conditional_format("remove", "foo.xlsx", "Sheet1") == "REMOVED"
    assert main.conditional_format("top_bottom", "foo.xlsx", "Sheet1", cell_range="A1:A2")["ok"]
    assert main.conditional_format("above_below_average", "foo.xlsx", "Sheet1", cell_range="A1:A2")["ok"]

    with pytest.raises(ValueError, match="Unknown action"):
        main.conditional_format("bad", "foo.xlsx", "Sheet1")

    # Table operations
    monkeypatch.setattr(main._tables, "create_table", lambda *args, **kwargs: "TCREATE")
    monkeypatch.setattr(main._tables, "list_tables", lambda *args, **kwargs: [])
    monkeypatch.setattr(main._tables, "resize_table", lambda *args, **kwargs: "TRESIZE")
    monkeypatch.setattr(main._tables, "set_table_totals_row", lambda *args, **kwargs: "TTOTALS")
    monkeypatch.setattr(main._tables, "get_table_data", lambda *args, **kwargs: {"rows": []})
    monkeypatch.setattr(main._tables, "convert_table_to_range", lambda *args, **kwargs: "TCONVERT")

    assert main.table("create", "foo.xlsx", "Sheet1", data_range="A1:B2", table_name="T1") == "TCREATE"
    assert main.table("list", "foo.xlsx", "Sheet1") == []
    assert main.table("resize", "foo.xlsx", "Sheet1", table_name="T1", new_range="A1:C3") == "TRESIZE"
    assert main.table("totals", "foo.xlsx", "Sheet1", table_name="T1", show_totals=True) == "TTOTALS"
    assert main.table("data", "foo.xlsx", "Sheet1", table_name="T1")["rows"] == []
    assert main.table("convert_to_range", "foo.xlsx", "Sheet1", table_name="T1") == "TCONVERT"


def test_data_validation_protection_chart(monkeypatch):
    monkeypatch.setattr(main._data_val, "add_dropdown_validation", lambda *args, **kwargs: "DD")
    monkeypatch.setattr(main._data_val, "add_numeric_validation", lambda *args, **kwargs: "DN")
    monkeypatch.setattr(main._data_val, "add_date_validation", lambda *args, **kwargs: "DDATE")
    monkeypatch.setattr(main._data_val, "remove_validation", lambda *args, **kwargs: "DREM")
    monkeypatch.setattr(main._data_val, "add_formula_validation", lambda *args, **kwargs: {"ok": True})

    assert main.data_validation("dropdown", "foo.xlsx", "Sheet1", "A1:A2", options=["a"]) == "DD"
    assert (
        main.data_validation("numeric", "foo.xlsx", "Sheet1", "A1:A2", operator="between", value1=1, value2=10) == "DN"
    )
    assert main.data_validation("date", "foo.xlsx", "Sheet1", "A1:A2") == "DDATE"
    assert main.data_validation("remove", "foo.xlsx", "Sheet1", "A1:A2") == "DREM"
    assert main.data_validation("formula", "foo.xlsx", "Sheet1", "A1:A2", formula="A1>0")["ok"]

    monkeypatch.setattr(main._protection, "protect_sheet", lambda *args, **kwargs: "PS")
    monkeypatch.setattr(main._protection, "unprotect_sheet", lambda *args, **kwargs: "UPS")
    monkeypatch.setattr(main._protection, "protect_cells", lambda *args, **kwargs: "PC")
    monkeypatch.setattr(main._doc_props, "protect_workbook", lambda *args, **kwargs: "PW")
    monkeypatch.setattr(main._doc_props, "unprotect_workbook", lambda *args, **kwargs: "UPW")

    assert main.protection("protect_sheet", "foo.xlsx", "Sheet1") == "PS"
    assert main.protection("unprotect_sheet", "foo.xlsx", "Sheet1") == "UPS"
    assert main.protection("protect_cells", "foo.xlsx", "Sheet1", locked_range="A1:A2") == "PC"
    assert main.protection("protect_workbook", "foo.xlsx") == "PW"
    assert main.protection("unprotect_workbook", "foo.xlsx") == "UPW"

    monkeypatch.setattr(main._charts, "create_chart", lambda *args, **kwargs: "C")
    monkeypatch.setattr(main._charts, "delete_chart", lambda *args, **kwargs: "D")
    monkeypatch.setattr(main._charts, "list_charts", lambda *args, **kwargs: [])
    monkeypatch.setattr(main._charts, "add_chart_series", lambda *args, **kwargs: "AS")
    monkeypatch.setattr(main._charts, "set_chart_axes", lambda *args, **kwargs: "AX")
    monkeypatch.setattr(main._charts, "add_chart_trendline", lambda *args, **kwargs: "T")
    monkeypatch.setattr(main._charts, "create_combo_chart", lambda *args, **kwargs: "CO")
    monkeypatch.setattr(main._charts, "set_chart_data_labels", lambda *args, **kwargs: "DL")
    monkeypatch.setattr(main._charts, "set_chart_legend", lambda *args, **kwargs: "LG")
    monkeypatch.setattr(main._charts, "update_chart", lambda *args, **kwargs: "U")

    assert main.chart("create", "foo.xlsx", "Sheet1", data_range="A1:B2") == "C"
    assert main.chart("delete", "foo.xlsx", "Sheet1") == "D"
    assert main.chart("list", "foo.xlsx", "Sheet1") == []
    assert main.chart("add_series", "foo.xlsx", "Sheet1", chart_index=0, data_range="A1:B2") == "AS"
    assert main.chart("set_axes", "foo.xlsx", "Sheet1") == "AX"
    assert main.chart("trendline", "foo.xlsx", "Sheet1") == "T"
    assert main.chart("combo", "foo.xlsx", "Sheet1", data_range="A1:B2", bar_columns=[0], line_columns=[1]) == "CO"
    assert main.chart("data_labels", "foo.xlsx", "Sheet1", chart_title="MyChart") == "DL"
    assert main.chart("legend", "foo.xlsx", "Sheet1", chart_title="MyChart") == "LG"
    assert main.chart("update", "foo.xlsx", "Sheet1") == "U"

    with pytest.raises(ValueError, match="Unknown action"):
        main.chart("bad", "foo.xlsx", "Sheet1")


def test_doc_named_comment_hyperlink_scenario(monkeypatch):
    monkeypatch.setattr(main._doc_props, "get_document_properties", lambda *args, **kwargs: {"ok": True})
    monkeypatch.setattr(main._doc_props, "set_calculation_mode", lambda *args, **kwargs: "CALC")
    monkeypatch.setattr(main._named_ranges, "list_named_ranges", lambda *args, **kwargs: [])
    monkeypatch.setattr(main._named_ranges, "create_named_range", lambda *args, **kwargs: "CREATED")
    monkeypatch.setattr(main._named_ranges, "delete_named_range", lambda *args, **kwargs: "DELETED")
    monkeypatch.setattr(main._named_ranges, "update_named_range", lambda *args, **kwargs: "UPDATED")
    monkeypatch.setattr(main._comments, "add_comment", lambda *args, **kwargs: "ADDED")
    monkeypatch.setattr(main._comments, "read_comment", lambda *args, **kwargs: "READ")
    monkeypatch.setattr(main._comments, "delete_comment", lambda *args, **kwargs: "DEL")
    monkeypatch.setattr(main._comments, "list_comments", lambda *args, **kwargs: [])
    monkeypatch.setattr(main._hyperlinks, "add_hyperlink", lambda *args, **kwargs: "HADD")
    monkeypatch.setattr(main._hyperlinks, "read_hyperlink", lambda *args, **kwargs: "HREAD")
    monkeypatch.setattr(main._hyperlinks, "delete_hyperlink", lambda *args, **kwargs: "HDEL")
    monkeypatch.setattr(main._hyperlinks, "list_hyperlinks", lambda *args, **kwargs: [])
    monkeypatch.setattr(main._scenarios, "add_scenario", lambda *args, **kwargs: "SADD")
    monkeypatch.setattr(main._scenarios, "list_scenarios", lambda *args, **kwargs: [])
    monkeypatch.setattr(main._scenarios, "apply_scenario", lambda *args, **kwargs: "SAPPLY")

    assert main.doc_properties("get", "foo.xlsx")["ok"]
    assert main.doc_properties("set_calc_mode", "foo.xlsx", calc_mode="manual") == "CALC"
    assert main.named_range("list", "foo.xlsx") == []
    assert main.named_range("create", "foo.xlsx", name="N", destination="Sheet1!A1") == "CREATED"
    assert main.named_range("delete", "foo.xlsx", name="N") == "DELETED"
    assert main.named_range("update", "foo.xlsx", name="N", new_destination="Sheet1!B1") == "UPDATED"
    assert main.comment("add", "foo.xlsx", "Sheet1", cell_ref="A1", text="v") == "ADDED"
    assert main.comment("read", "foo.xlsx", "Sheet1", cell_ref="A1") == "READ"
    assert main.comment("delete", "foo.xlsx", "Sheet1", cell_ref="A1") == "DEL"
    assert main.comment("list", "foo.xlsx", "Sheet1") == []
    assert main.hyperlink("add", "foo.xlsx", "Sheet1", cell_ref="A1", url="http://x") == "HADD"
    assert main.hyperlink("read", "foo.xlsx", "Sheet1", cell_ref="A1") == "HREAD"
    assert main.hyperlink("delete", "foo.xlsx", "Sheet1", cell_ref="A1") == "HDEL"
    assert main.hyperlink("list", "foo.xlsx", "Sheet1") == []
    assert main.scenario("add", "foo.xlsx", name="Sc", cell_values={"Sheet1": {"A1": 1}}) == "SADD"
    assert main.scenario("list", "foo.xlsx") == []
    assert main.scenario("apply", "foo.xlsx", name="Sc") == "SAPPLY"


def test_multi_file_and_worksheet_ops(monkeypatch):
    monkeypatch.setattr(main._multi_file, "bulk_aggregate_multi_files", lambda *args, **kwargs: {"agg": 1})
    monkeypatch.setattr(main._multi_file, "bulk_filter_multi_files", lambda *args, **kwargs: {"filtered": 1})
    monkeypatch.setattr(main._multi_file, "validate_data_consistency", lambda *args, **kwargs: {"valid": True})
    monkeypatch.setattr(main._multi_file, "compare_workbooks", lambda *args, **kwargs: {"diff": []})

    assert main.multi_file("aggregate", file_paths=["a.xlsx"], column="A")["agg"] == 1
    assert main.multi_file("filter", file_paths=["a.xlsx"], column="A", operator="=", value=1)["filtered"] == 1
    assert main.multi_file("validate", file_paths=["a.xlsx"], key_column="A")["valid"] is True
    assert main.multi_file("compare", file_a="a.xlsx", file_b="b.xlsx")["diff"] == []

    monkeypatch.setattr(main._ws_ops, "freeze_panes", lambda *args, **kwargs: "FROZEN")
    monkeypatch.setattr(main._ws_ops, "set_auto_filter", lambda *args, **kwargs: "FILTERED")
    monkeypatch.setattr(main._ws_ops, "copy_range_across_sheets", lambda *args, **kwargs: "CRANGE")
    monkeypatch.setattr(main._ws_ops, "copy_sheet_across_workbooks", lambda *args, **kwargs: "CSHEET")
    monkeypatch.setattr(main._ws_ops, "merge_workbooks", lambda *args, **kwargs: "MERGED")
    monkeypatch.setattr(main._ws_ops, "insert_rows", lambda *args, **kwargs: "IR")
    monkeypatch.setattr(main._ws_ops, "delete_rows", lambda *args, **kwargs: "DR")
    monkeypatch.setattr(main._ws_ops, "insert_cols", lambda *args, **kwargs: "IC")
    monkeypatch.setattr(main._ws_ops, "delete_cols", lambda *args, **kwargs: "DC")
    monkeypatch.setattr(main._ws_ops, "set_print_area", lambda *args, **kwargs: "SPA")
    monkeypatch.setattr(main._ws_ops, "set_page_setup", lambda *args, **kwargs: "SPS")
    monkeypatch.setattr(main._ws_ops, "group_rows", lambda *args, **kwargs: "GR")
    monkeypatch.setattr(main._ws_ops, "group_cols", lambda *args, **kwargs: "GC")
    monkeypatch.setattr(main._ws_ops, "ungroup_rows", lambda *args, **kwargs: "UGR")
    monkeypatch.setattr(main._ws_ops, "ungroup_cols", lambda *args, **kwargs: "UGC")
    monkeypatch.setattr(main._ws_ops, "set_print_titles", lambda *args, **kwargs: "SPT")
    monkeypatch.setattr(main._ws_ops, "set_row_height", lambda *args, **kwargs: "SRH")
    monkeypatch.setattr(main._ws_ops, "set_col_width", lambda *args, **kwargs: "SCW")
    monkeypatch.setattr(main._ws_ops, "set_gridlines", lambda *args, **kwargs: "SGL")
    monkeypatch.setattr(main._ws_ops, "stack_sheets", lambda *args, **kwargs: "STACK")
    monkeypatch.setattr(main._ws_ops, "add_page_break", lambda *args, **kwargs: "APB")
    monkeypatch.setattr(main._ws_ops, "remove_page_break", lambda *args, **kwargs: "RPB")

    assert main.worksheet_view("freeze", file_path="foo.xlsx", sheet_name="Sheet1") == "FROZEN"
    assert (
        main.worksheet_view("auto_filter", file_path="foo.xlsx", sheet_name="Sheet1", cell_range="A1:A2", remove=False)
        == "FILTERED"
    )
    assert (
        main.worksheet_transfer(
            "copy_range_across",
            file_path="foo.xlsx",
            source_sheet="Sheet1",
            source_range="A1:A2",
            target_sheet="Sheet2",
        )
        == "CRANGE"
    )
    assert (
        main.worksheet_transfer(
            "copy_sheet_across", source_file="foo.xlsx", source_sheet="Sheet1", dest_file="foo2.xlsx"
        )
        == "CSHEET"
    )
    assert main.worksheet_transfer("merge_workbooks", source_files=["foo.xlsx"], output_file="out.xlsx") == "MERGED"
    assert main.worksheet_structure("insert_rows", file_path="foo.xlsx", sheet_name="Sheet1", row=2) == "IR"
    assert main.worksheet_structure("delete_rows", file_path="foo.xlsx", sheet_name="Sheet1", row=2) == "DR"
    assert main.worksheet_structure("insert_cols", file_path="foo.xlsx", sheet_name="Sheet1", col=2) == "IC"
    assert main.worksheet_structure("delete_cols", file_path="foo.xlsx", sheet_name="Sheet1", col=2) == "DC"
    assert (
        main.worksheet_print("set_print_area", file_path="foo.xlsx", sheet_name="Sheet1", print_area="A1:B2") == "SPA"
    )
    assert main.worksheet_print("set_page_setup", file_path="foo.xlsx", sheet_name="Sheet1") == "SPS"
    assert (
        main.worksheet_structure("group_rows", file_path="foo.xlsx", sheet_name="Sheet1", start_row=1, end_row=2)
        == "GR"
    )
    assert (
        main.worksheet_structure("group_cols", file_path="foo.xlsx", sheet_name="Sheet1", start_col=1, end_col=2)
        == "GC"
    )
    assert (
        main.worksheet_structure("ungroup_rows", file_path="foo.xlsx", sheet_name="Sheet1", start_row=1, end_row=2)
        == "UGR"
    )
    assert (
        main.worksheet_structure("ungroup_cols", file_path="foo.xlsx", sheet_name="Sheet1", start_col=1, end_col=2)
        == "UGC"
    )
    assert (
        main.worksheet_print("set_print_titles", file_path="foo.xlsx", sheet_name="Sheet1", title_rows="1:1") == "SPT"
    )
    assert (
        main.worksheet_structure("set_row_height", file_path="foo.xlsx", sheet_name="Sheet1", rows=[1], height=20)
        == "SRH"
    )
    assert (
        main.worksheet_structure("set_col_width", file_path="foo.xlsx", sheet_name="Sheet1", cols_list=["A"], width=20)
        == "SCW"
    )
    assert main.worksheet_view("set_gridlines", file_path="foo.xlsx", sheet_name="Sheet1", show=True) == "SGL"
    assert (
        main.worksheet_transfer(
            "stack_sheets", file_path="foo.xlsx", sheet_names=["Sheet1"], dest_sheet="Stacked", output_path="out.xlsx"
        )
        == "STACK"
    )
    assert main.worksheet_print("add_page_break", file_path="foo.xlsx", sheet_name="Sheet1", row=2) == "APB"
    assert main.worksheet_print("remove_page_break", file_path="foo.xlsx", sheet_name="Sheet1", row=2) == "RPB"


def test_analysis_pivot_financial_shortcuts(monkeypatch):
    monkeypatch.setattr(main._analysis, "sort_data", lambda *args, **kwargs: "S")
    monkeypatch.setattr(main._analysis, "column_statistics", lambda *args, **kwargs: "CS")
    monkeypatch.setattr(main._analysis, "aggregate_data", lambda *args, **kwargs: "AD")
    monkeypatch.setattr(main._analysis, "find_duplicates", lambda *args, **kwargs: "FD")
    monkeypatch.setattr(main._analysis, "vlookup_helper", lambda *args, **kwargs: "VL")
    monkeypatch.setattr(main._analysis, "filter_data_advanced", lambda *args, **kwargs: "FDA")
    monkeypatch.setattr(main._analysis, "insert_subtotals", lambda *args, **kwargs: "IS")

    assert main.sort_data("foo.xlsx", "Sheet1") == "S"
    assert main.column_statistics("foo.xlsx", "Sheet1", "A") == "CS"
    assert main.aggregate_data("foo.xlsx", "Sheet1", "A", "B") == "AD"
    assert main.find_duplicates("foo.xlsx", "Sheet1", ["A"]) == "FD"
    assert main.vlookup_helper("a.xlsx", "b.xlsx", "A", "A", ["B"]) == "VL"
    assert main.filter_data_advanced("foo.xlsx", "Sheet1", []) == "FDA"
    assert main.insert_subtotals("foo.xlsx", "Sheet1", "A", "B") == "IS"

    monkeypatch.setattr(main._pivot_etl, "create_pivot_table", lambda *args, **kwargs: "CPT")
    monkeypatch.setattr(main._pivot_etl, "refresh_pivot_table", lambda *args, **kwargs: "RPT")
    monkeypatch.setattr(main._pivot_etl, "unpivot_data", lambda *args, **kwargs: "UP")
    monkeypatch.setattr(main._pivot_etl, "merge_datasets", lambda *args, **kwargs: "MD")
    monkeypatch.setattr(main._pivot_etl, "add_computed_column", lambda *args, **kwargs: "AC")
    monkeypatch.setattr(main._pivot_etl, "deduplicate_data", lambda *args, **kwargs: "DD")

    assert main.create_pivot_table("foo.xlsx", "Sheet1", ["A"], ["B"]) == "CPT"
    assert main.refresh_pivot_table("foo.xlsx", "OutSheet") == "RPT"
    assert main.unpivot_data("foo.xlsx", "Sheet1", ["A"], ["B"]) == "UP"
    assert main.merge_datasets("foo.xlsx", "Sheet1", "Sheet1") == "MD"
    assert main.add_computed_column("foo.xlsx", "Sheet1", "C", "A+B") == "AC"
    assert main.deduplicate_data("foo.xlsx", "Sheet1") == "DD"

    monkeypatch.setattr(main._financial, "goal_seek", lambda *args, **kwargs: "GS")
    monkeypatch.setattr(main._financial, "loan_amortization", lambda *args, **kwargs: "LA")
    monkeypatch.setattr(main._financial, "dcf_analysis", lambda *args, **kwargs: "DCF")
    monkeypatch.setattr(main._financial, "budget_variance_analysis", lambda *args, **kwargs: "BA")
    monkeypatch.setattr(main._financial, "financial_ratio_analysis", lambda *args, **kwargs: "FR")
    monkeypatch.setattr(main._financial, "break_even_analysis", lambda *args, **kwargs: "BE")

    assert main.goal_seek("foo.xlsx", "Sheet1", "A1", "A1", 1.0) == "GS"
    assert main.loan_amortization(100, 0.05, 5) == "LA"
    assert main.dcf_analysis([100], 0.1) == "DCF"
    assert main.budget_variance_analysis("foo.xlsx", "Sheet1") == "BA"
    assert main.financial_ratio_analysis({}) == "FR"
    assert main.break_even_analysis(100, 10, 5) == "BE"
