"""End-to-end pytest suite mimicking the 5 Excel-expert workflows across all 69 MCP tools."""

import base64
import json
from pathlib import Path
from typing import Any

import pytest
from PIL import Image

from mcp_server.main import mcp


async def _call(tool_name: str, args: dict[str, Any]) -> Any:
    """Invoke an MCP tool in-process and return parsed JSON or text, asserting no error."""
    result = await mcp.call_tool(tool_name, args)
    # FastMCP call_tool returns (content_list, is_error) or CallToolResult depending on version
    if isinstance(result, tuple):
        contents, meta = result
        is_err = bool(meta.get("isError")) if isinstance(meta, dict) else False
    elif hasattr(result, "content"):
        contents = result.content
        is_err = bool(getattr(result, "is_error", getattr(result, "isError", False)))
    else:
        contents = result
        is_err = False

    text = "\n".join(c.text for c in contents if hasattr(c, "text"))
    assert not is_err, f"Tool '{tool_name}' returned error: {text}"
    try:
        parsed = json.loads(text)
        if isinstance(parsed, dict):
            assert parsed.get("status") != "error", f"Tool '{tool_name}' failed: {parsed}"
            assert not parsed.get("error"), f"Tool '{tool_name}' error: {parsed}"
        return parsed
    except json.JSONDecodeError:
        return text


@pytest.mark.anyio
async def test_workflow_1_fpa_modeling(tmp_path: Path) -> None:
    """Workflow 1: Corporate FP&A & Financial Modeling (19 tools)."""
    wb = str(tmp_path / "wf1_fpa_model.xlsx")

    res = await _call("create_workbook", {"file_path": wb, "sheet_names": ["Assumptions", "BudgetVsActual", "Model"]})
    assert "Assumptions" in res["sheets"]

    await _call("write_multi_sheet", {
        "file_path": wb,
        "sheets": [
            {
                "name": "Assumptions",
                "headers": ["Parameter", "Value"],
                "data": [
                    ["Revenue_Y1", 1000000],
                    ["Growth_Rate", 0.12],
                    ["COGS_Pct", 0.45],
                    ["OpEx", 250000],
                    ["Tax_Rate", 0.21],
                    ["WACC", 0.09],
                    ["Terminal_Growth", 0.025],
                ],
            },
            {
                "name": "BudgetVsActual",
                "headers": ["Category", "Budget", "Actual"],
                "data": [
                    ["Cloud Infra", 120000, 138000],
                    ["Payroll", 500000, 495000],
                    ["Marketing", 85000, 110000],
                    ["R&D", 200000, 192000],
                    ["G&A", 65000, 68500],
                ],
            },
        ],
    })

    meta = await _call("get_workbook_metadata", {"file_path": wb})
    assert len(meta["sheets"]) == 2

    summary = await _call("get_sheet_summary", {"file_path": wb, "sheet_name": "Assumptions"})
    assert summary["headers"] == ["Parameter", "Value"]

    await _call("named_range", {"file_path": wb, "action": "create", "name": "WACC", "destination": "Assumptions!B7"})
    nr_list = await _call("named_range", {"file_path": wb, "action": "list"})
    assert nr_list["name"] == "WACC"

    await _call("table", {"file_path": wb, "sheet_name": "BudgetVsActual", "action": "create", "table_name": "TblBudget", "data_range": "A1:C6"})
    tbl_list = await _call("table", {"file_path": wb, "sheet_name": "BudgetVsActual", "action": "list"})
    assert tbl_list["name"] == "TblBudget"

    await _call("formula_write", {"file_path": wb, "sheet_name": "BudgetVsActual", "action": "batch", "formulas": {"D2": "=C2-B2", "D3": "=C3-B3", "D4": "=C4-B4"}})
    auto_sum = await _call("formula_write", {"file_path": wb, "sheet_name": "BudgetVsActual", "action": "auto_sum", "cell_ref": "B7", "source_range": "B2:B6"})
    assert auto_sum["formula"] == "=SUM(B2:B6)"

    await _call("formula_audit", {"file_path": wb, "sheet_name": "BudgetVsActual", "action": "list"})

    bva = await _call("budget_variance_analysis", {"file_path": wb, "sheet_name": "BudgetVsActual", "category_col": "Category", "budget_col": "Budget", "actual_col": "Actual"})
    assert len(bva["items"]) == 5

    loan = await _call("loan_amortization", {"file_path": wb, "output_sheet": "DebtSchedule", "principal": 500000, "annual_rate": 0.065, "years": 5, "payments_per_year": 12})
    assert loan["monthly_payment"] == pytest.approx(9783.07, abs=0.05)

    dcf = await _call("dcf_analysis", {"cash_flows": [150000, 180000, 215000, 250000, 290000], "discount_rate": 0.09, "terminal_growth_rate": 0.025, "file_path": wb, "output_sheet": "DCF"})
    assert dcf["enterprise_value"] > 0

    ratios = await _call("financial_ratio_analysis", {"financial_data": {"current_assets": 450000, "current_liabilities": 210000, "inventory": 60000, "total_debt": 500000, "total_equity": 800000, "net_income": 185000, "revenue": 1200000, "ebit": 260000, "interest_expense": 32500}})
    assert ratios["ratios"]["current_ratio"]["value"] == pytest.approx(2.1429, abs=1e-3)

    be = await _call("break_even_analysis", {"fixed_costs": 250000, "variable_cost_per_unit": 45, "price_per_unit": 95})
    assert be["break_even_units"] == pytest.approx(5000.0)

    irr = await _call("time_value_calc", {"operation": "irr", "cash_flows": [-500000, 150000, 180000, 215000, 250000]})
    assert irr["irr"] == pytest.approx(0.195148, abs=1e-4)

    sens = await _call("create_sensitivity_table", {"file_path": wb, "sheet_name": "Assumptions", "output_cell": "D2", "expression": "1000000 * (1 + growth) * margin", "var1_name": "growth", "var1_values": [0.06, 0.09, 0.12], "var2_name": "margin", "var2_values": [0.15, 0.20, 0.25]})
    assert sens["var1"] == "growth"

    gs = await _call("goal_seek", {"file_path": wb, "sheet_name": "Assumptions", "variable_cell": "B2", "expression": "x * (1 - 0.45) - 250000", "target_value": 400000})
    assert gs["converged"] is True
    assert gs["found_value"] == pytest.approx(1181818.18, abs=1.0)

    solver = await _call("run_solver", {"file_path": wb, "sheet_name": "Assumptions", "variable_cells": {"B2": {"lower": 500000, "upper": 2000000, "initial": 1000000}, "B4": {"lower": 150000, "upper": 400000, "initial": 250000}}, "objective_expression": "B2 * 0.55 - B4", "maximize": True})
    assert solver["converged"] is True
    assert solver["objective_value"] == pytest.approx(950000.0, abs=1.0)

    await _call("scenario", {"file_path": wb, "action": "add", "name": "UpsideCase", "cell_values": {"Assumptions": {"B2": 1500000, "B3": 0.18}}, "description": "High growth upside"})
    await _call("scenario", {"file_path": wb, "action": "list"})
    props = await _call("doc_properties", {"file_path": wb, "action": "get"})
    assert "creator" in props


@pytest.mark.anyio
async def test_workflow_2_data_engineering_etl(tmp_path: Path) -> None:
    """Workflow 2: Data Engineering, ETL & Cleaning Pipeline (17 tools)."""
    csv_in = tmp_path / "wf2_raw_sales.csv"
    csv_in.write_text(
        "OrderID,OrderDate,RegionSegment,Units,Price,Notes\n"
        "1001,2026-01-15,North-Enterprise,10,120.50,  urgent  \n"
        "1002,01/20/2026,South-SMB,5,80.00,\n"
        "1003,Jan 25 2026,North-Enterprise,12,120.50,  standard \n"
        "1003,Jan 25 2026,North-Enterprise,12,120.50,  standard \n"
        "1004,2026/02/02,West-MidMarket,8,95.00,ok\n"
        "1005,2026-02-10,East-Enterprise,15,150.00,  priority \n"
    )
    wb = str(tmp_path / "wf2_etl.xlsx")
    wb_lookup = str(tmp_path / "wf2_lookup.xlsx")

    preview = await _call("csv_ops", {"action": "preview", "file_path": str(csv_in), "rows": 5})
    assert len(preview["headers"]) == 6
    await _call("csv_ops", {"action": "to_xlsx", "csv_path": str(csv_in), "xlsx_path": wb, "sheet_name": "RawSales"})

    prof = await _call("profile_data", {"file_path": wb, "sheet": "RawSales"})
    assert prof["row_count"] == 6

    cleaned = await _call("data_cleaner", {"file_path": wb, "sheet_name": "RawSales", "operations": ["trim", "fill_missing"], "fill_missing_strategy": "value", "fill_value": "N/A"})
    assert cleaned["rows_after"] == 6

    parsed = await _call("parse_date_column", {"file_path": wb, "sheet_name": "RawSales", "column": "OrderDate", "output_column": "OrderDateISO", "output_format": "%Y-%m-%d"})
    assert parsed["parsed_count"] == 6

    split = await _call("split_column", {"file_path": wb, "sheet_name": "RawSales", "column": "RegionSegment", "delimiter": "-", "new_columns": ["Region", "Segment"]})
    assert split["new_columns"] == ["Region", "Segment"]

    dups = await _call("find_duplicates", {"file_path": wb, "sheet_name": "RawSales", "columns": ["OrderID"]})
    assert len(dups["duplicates"]) == 2
    await _call("deduplicate_data", {"file_path": wb, "sheet_name": "RawSales", "columns": ["OrderID"], "keep": "first"})

    await _call("add_computed_column", {"file_path": wb, "sheet_name": "RawSales", "new_column_name": "Revenue", "expression": "Units * Price"})
    await _call("sort_data", {"file_path": wb, "sheet_name": "RawSales", "column": "Region", "ascending": True})

    filtered = await _call("filter_data_advanced", {"file_path": wb, "sheet_name": "RawSales", "conditions": [{"column": "Units", "operator": ">=", "value": 10}], "logic": "AND", "output_sheet": "LargeOrders"})
    assert filtered["rows"] == 3

    sub = await _call("insert_subtotals", {"file_path": wb, "sheet_name": "RawSales", "group_col": "Region", "value_col": "Revenue", "subtotal_func": 9})
    assert sub["groups"] == 4

    vc = await _call("value_counts", {"file_path": wb, "sheet_name": "RawSales", "column": "Segment"})
    assert vc["total_rows"] == 5

    agg = await _call("aggregate_data", {"file_path": wb, "sheet_name": "RawSales", "group_by": "Region", "value_column": "Revenue", "operation": "sum"})
    assert len(agg["groups"]) >= 4

    stats = await _call("column_statistics", {"file_path": wb, "sheet_name": "RawSales", "column": "Revenue"})
    assert stats["sum_val"] == pytest.approx(6061.0)

    await _call("write_multi_sheet", {
        "file_path": wb,
        "sheets": [
            {"name": "QuarterlyWide", "headers": ["Region", "Q1", "Q2", "Q3"], "data": [["North", 100, 120, 140], ["South", 80, 90, 110]]},
            {"name": "RegionOwners", "headers": ["Region", "VP"], "data": [["North", "Alice"], ["South", "Bob"]]},
        ],
    })
    unpiv = await _call("unpivot_data", {"file_path": wb, "sheet_name": "QuarterlyWide", "id_vars": ["Region"], "value_vars": ["Q1", "Q2", "Q3"], "var_name": "Quarter", "value_name": "Sales"})
    assert len(unpiv["data"]) == 6

    merged = await _call("merge_datasets", {"file_path": wb, "sheet1": "QuarterlyWide", "sheet2": "RegionOwners", "join_key": "Region", "how": "left", "output_sheet": "MergedRegions"})
    assert len(merged["data"]) == 2

    await _call("write_multi_sheet", {"file_path": wb_lookup, "sheets": [{"name": "Lookup", "headers": ["Region", "TaxRate"], "data": [["North", 0.08], ["South", 0.06]]}]})
    vlook = await _call("vlookup_helper", {"lookup_file": wb, "data_file": wb_lookup, "lookup_sheet": "QuarterlyWide", "data_sheet": "Lookup", "lookup_column": "Region", "data_key_column": "Region", "data_return_columns": ["TaxRate"]})
    assert vlook["matched"] == 2


@pytest.mark.anyio
async def test_workflow_3_executive_reporting(tmp_path: Path) -> None:
    """Workflow 3: Executive Reporting, Formatting, Charts & Presentation (21 tools)."""
    wb = str(tmp_path / "wf3_exec_report.xlsx")
    img = str(tmp_path / "logo.png")
    Image.new("RGB", (60, 30), color=(30, 90, 160)).save(img)

    await _call("create_workbook", {"file_path": wb, "sheet_names": ["Dashboard", "Archive"]})
    await _call("write_cells", {
        "mode": "range",
        "file_path": wb,
        "sheet_name": "Dashboard",
        "start_cell": "A1",
        "data": [
            ["Quarter", "Revenue", "EBITDA", "Margin"],
            ["Q1 2026", 1200000, 300000, 0.25],
            ["Q2 2026", 1450000, 391500, 0.27],
            ["Q3 2026", 1600000, 448000, 0.28],
            ["Q4 2026", 1900000, 570000, 0.30],
        ],
    })
    read_res = await _call("read_cells", {"mode": "range", "file_path": wb, "sheet_name": "Dashboard", "start_cell": "A1", "end_cell": "D5"})
    assert len(read_res["rows"]) == 5

    await _call("format_cells", {"file_path": wb, "sheet_name": "Dashboard", "cell_range": "A1:D1", "bold": True, "font_color": "FFFFFF", "bg_color": "1F4E78", "horizontal_alignment": "center"})
    await _call("apply_named_style", {"file_path": wb, "sheet_name": "Dashboard", "range_str": "D5", "style_name": "Good"})
    await _call("copy_cell_format", {"file_path": wb, "sheet_name": "Dashboard", "source_cell": "B2", "target_range": "B6:C6"})
    await _call("clear_cell_format", {"file_path": wb, "sheet_name": "Dashboard", "range_str": "B6:C6"})
    await _call("auto_fit_columns", {"file_path": wb, "sheet_name": "Dashboard"})

    await _call("copy_range", {"file_path": wb, "source_sheet": "Dashboard", "source_range": "A1:D5", "dest_sheet": "Archive", "dest_range": "A1:D5"})
    trans = await _call("transpose_range", {"file_path": wb, "sheet_name": "Archive", "source_range": "A1:D5", "target_cell": "A8"})
    assert trans["target_shape"] == [4, 5]
    fr = await _call("find_replace", {"file_path": wb, "sheet_name": "Archive", "find_text": "2026", "replace_text": "FY26"})
    assert fr["replacements_made"] == 8
    await _call("clear_range", {"file_path": wb, "sheet_name": "Archive", "start_cell": "A8", "end_cell": "E11"})

    await _call("chart", {"action": "create", "file_path": wb, "sheet_name": "Dashboard", "data_range": "A1:C5", "chart_type": "bar", "target_cell": "F2", "title": "Quarterly Revenue & EBITDA"})
    await _call("conditional_format", {"action": "apply", "file_path": wb, "sheet_name": "Dashboard", "cell_range": "D2:D5", "format_type": "color_scale", "start_color": "F8696B", "end_color": "63BE7B"})
    await _call("worksheet_view", {"action": "freeze", "file_path": wb, "sheet_name": "Dashboard", "cell_ref": "A2"})
    await _call("worksheet_structure", {"action": "set_row_height", "file_path": wb, "sheet_name": "Dashboard", "rows": [1], "height": 28})
    await _call("worksheet_print", {"action": "set_page_setup", "file_path": wb, "sheet_name": "Dashboard", "orientation": "landscape", "fit_to_width": 1})
    await _call("comment", {"action": "add", "file_path": wb, "sheet_name": "Dashboard", "cell_ref": "B5", "text": "Includes enterprise renewals", "author": "CFO"})
    await _call("hyperlink", {"action": "add", "file_path": wb, "sheet_name": "Dashboard", "cell_ref": "A7", "url": "https://example.com/q4-board-deck", "display_text": "Board Deck"})
    await _call("data_validation", {"action": "dropdown", "file_path": wb, "sheet_name": "Dashboard", "cell_range": "E2:E5", "options": ["Approved", "Pending", "Review"]})
    await _call("insert_image", {"file_path": wb, "sheet": "Dashboard", "image_path": img, "cell": "A9"})
    await _call("protection", {"action": "protect_sheet", "file_path": wb, "sheet_name": "Archive", "password": "secret"})


@pytest.mark.anyio
async def test_workflow_4_advanced_analytics(tmp_path: Path) -> None:
    """Workflow 4: Advanced Analytics, Forecasting & Custom Code (6 tools)."""
    wb = str(tmp_path / "wf4_analytics.xlsx")
    await _call("write_multi_sheet", {
        "file_path": wb,
        "sheets": [{
            "name": "SalesData",
            "headers": ["Month", "Region", "AdSpend", "Headcount", "Revenue"],
            "data": [
                [1, "North", 10000, 5, 120000],
                [2, "North", 12000, 6, 138000],
                [3, "North", 14000, 6, 152000],
                [4, "North", 15000, 7, 165000],
                [5, "South", 18000, 8, 190000],
                [6, "South", 20000, 9, 212000],
                [7, "South", 22000, 9, 228000],
                [8, "South", 25000, 10, 255000],
            ],
        }],
    })

    piv = await _call("create_pivot_table", {"file_path": wb, "sheet_name": "SalesData", "index_cols": ["Region"], "value_cols": ["Revenue", "AdSpend"], "aggfunc": "sum", "output_sheet": "PivotSummary"})
    assert len(piv["data"]) == 2
    ref = await _call("refresh_pivot_table", {"file_path": wb, "output_sheet": "PivotSummary"})
    assert ref["refreshed"] == "PivotSummary"

    reg = await _call("run_regression", {"file_path": wb, "sheet_name": "SalesData", "y_column": "Revenue", "x_columns": ["AdSpend", "Headcount"], "output_sheet": "Regression"})
    assert reg["r_squared"] > 0.99

    ewm = await _call("run_exponential_smoothing", {"file_path": wb, "sheet_name": "SalesData", "column": "Revenue", "alpha": 0.3, "method": "simple", "forecast_steps": 3})
    assert ewm["forecast_steps"] == 3

    corr = await _call("correlation_matrix", {"file_path": wb, "sheet_name": "SalesData", "columns": ["AdSpend", "Headcount", "Revenue"], "output_sheet": "Correlations"})
    assert len(corr["matrix"]) == 3

    custom = await _call("execute_custom_code", {"file_path": wb, "sheet": "SalesData", "code": "result = {'roas': float((df['Revenue'] / df['AdSpend']).mean()), 'rows': len(df)}"})
    assert custom["result"]["roas"] > 10.0


@pytest.mark.anyio
async def test_workflow_5_multi_workbook_and_file_lifecycle(tmp_path: Path) -> None:
    """Workflow 5: Multi-Workbook Consolidation & Sheet/File Lifecycle (6 tools)."""
    wb_a = str(tmp_path / "wf5_region_a.xlsx")
    wb_b = str(tmp_path / "wf5_region_b.xlsx")
    wb_out = str(tmp_path / "wf5_consolidated.xlsx")

    await _call("write_multi_sheet", {"file_path": wb_a, "sheets": [{"name": "Data", "headers": ["ID", "Sales"], "data": [[1, 100], [2, 200]]}]})
    await _call("write_multi_sheet", {"file_path": wb_b, "sheets": [{"name": "Data", "headers": ["ID", "Sales"], "data": [[1, 100], [2, 250]]}]})

    await _call("sheet_management", {"action": "copy", "file_path": wb_a, "sheet_name": "Data", "new_name": "DataBackup"})

    merged = await _call("worksheet_transfer", {"action": "merge_workbooks", "source_files": [wb_a, wb_b], "output_file": wb_out})
    assert merged["total_sheets"] == 3

    comp = await _call("multi_file", {"action": "compare", "file_a": wb_a, "file_b": wb_b, "sheet_name": "Data"})
    assert comp["total_differences"] == 1

    b64 = base64.b64encode(Path(wb_a).read_bytes()).decode("utf-8")
    up = await _call("upload_file", {"file_content": b64, "filename": "uploaded_region_a.xlsx"})
    assert "session_id" in up

    dl = await _call("download_file", {"file_path": wb_a})
    assert len(dl["file_content"]) > 0

    rel = await _call("release_file", {"session_id": up["session_id"]})
    assert rel["success"] is True
