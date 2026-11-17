"""Phase 4 remediation tests & 100% coverage for financial, solver, charts, and tables."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import openpyxl
import pytest
from openpyxl import Workbook

import mcp_server.routes.charts as chart_routes
import mcp_server.routes.financial as fin_routes
from mcp_server.tools import charts, financial, solver, tables


def _make_wb(path: Path, sheets: dict[str, list[list[object]]] | None = None) -> str:
    wb = Workbook()
    if sheets:
        first = True
        for name, rows in sheets.items():
            ws = wb.active if first else wb.create_sheet(title=name)
            ws.title = name
            first = False
            for r in rows:
                ws.append(r)
    else:
        wb.active.title = "Sheet1"
    wb.save(path)
    wb.close()
    return str(path)


# ── Financial Tools & Routes ────────────────────────────────────────────────


def test_financial_goal_seek_non_convergence_and_error(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    fp = _make_wb(tmp_path / "gs.xlsx", {"Sheet1": [[42.0]]})

    # Non-convergent equation (x^2 + 10 = 0): must NOT overwrite A1=42.0
    res = financial.goal_seek(fp, "Sheet1", "A1", "x**2 + 10", 0.0, initial_value=1.0, max_iterations=5)
    assert res["converged"] is False
    wb = openpyxl.load_workbook(fp)
    assert wb["Sheet1"]["A1"].value == 42.0
    wb.close()

    # Non-finite root from root_scalar -> converged=False, found_value=None
    import scipy.optimize

    monkeypatch.setattr(
        scipy.optimize,
        "root_scalar",
        lambda *a, **kw: SimpleNamespace(root=float("nan"), converged=False, iterations=3),
    )
    res_nan = financial.goal_seek(fp, "Sheet1", "A1", "x * 2", 10.0)
    assert res_nan["converged"] is False
    assert res_nan["found_value"] is None

    # False convergence (converged=True but residual exceeds tolerance) -> converged=False
    monkeypatch.setattr(
        scipy.optimize,
        "root_scalar",
        lambda *a, **kw: SimpleNamespace(root=1.0, converged=True, iterations=2),
    )
    res_bad_res = financial.goal_seek(fp, "Sheet1", "A1", "x**2 + 10", 0.0)
    assert res_bad_res["converged"] is False

    # Exception inside root_scalar -> ValueError
    monkeypatch.setattr(scipy.optimize, "root_scalar", lambda *a, **kw: (_ for _ in ()).throw(RuntimeError("boom")))
    with pytest.raises(ValueError, match="Goal seek failed to converge"):
        financial.goal_seek(fp, "Sheet1", "A1", "x * 2", 10.0)


def test_financial_loan_amortization_validation_and_final_payment() -> None:
    with pytest.raises(ValueError, match="years must be positive"):
        financial.loan_amortization(10000, 0.05, 0)
    with pytest.raises(ValueError, match="payments_per_year must be positive"):
        financial.loan_amortization(10000, 0.05, 1, payments_per_year=0)
    with pytest.raises(ValueError, match="max_periods must be non-negative"):
        financial.loan_amortization(10000, 0.05, 1, max_periods=-1)

    res = financial.loan_amortization(10000, 0.06, 1, payments_per_year=12)
    assert res["schedule"][-1]["balance"] == 0.0


def test_financial_dcf_validation_and_nan_irr(monkeypatch: pytest.MonkeyPatch) -> None:
    with pytest.raises(ValueError, match="greater than -1.0"):
        financial.dcf_analysis([100.0], discount_rate=-1.0, terminal_growth_rate=-2.0)
    with pytest.raises(ValueError, match="At least one cash flow"):
        financial.dcf_analysis([], discount_rate=0.1)

    # All-positive cash flows -> NaN IRR sanitized to None
    res = financial.dcf_analysis([100.0, 200.0], discount_rate=0.1, initial_investment=0.0)
    assert res["irr"] is None

    # Exception in npf.irr -> None
    monkeypatch.setattr(financial.npf, "irr", lambda _cfs: (_ for _ in ()).throw(RuntimeError("irr err")))
    res2 = financial.dcf_analysis([100.0, 200.0], discount_rate=0.1, initial_investment=50.0)
    assert res2["irr"] is None


def test_financial_budget_variance_zero_budget_non_numeric_and_multisheet(tmp_path: Path) -> None:
    fp = _make_wb(
        tmp_path / "bva.xlsx",
        {
            "Sheet1": [
                ["Category", "Budget", "Actual"],
                ["ZeroOn", 0, 0],
                ["ZeroOver", 0, 150],
                ["ZeroUnder", 0, -50],
                [None, 10, 20],
                ["BoolSkip", True, 10],
                ["TextSkip", "N/A", 10],
            ]
        },
    )
    out = _make_wb(tmp_path / "out.xlsx", {"KeepMe": [["hello"]], "Variance Analysis": [["old"]]})

    res = financial.budget_variance_analysis(fp, output_file=out)
    items = {i["category"]: i for i in res["items"]}
    assert items["ZeroOn"]["status"] == "on_budget"
    assert items["ZeroOn"]["variance_pct"] is None
    assert items["ZeroOver"]["status"] == "over_budget"
    assert items["ZeroUnder"]["status"] == "under_budget"
    assert "BoolSkip" not in items
    assert "TextSkip" not in items

    wb = openpyxl.load_workbook(out)
    assert "KeepMe" in wb.sheetnames
    assert wb["KeepMe"]["A1"].value == "hello"
    va = wb["Variance Analysis"]
    total_row = [va.cell(row=va.max_row, column=c).value for c in range(1, 7)]
    assert total_row[0] == "TOTAL"
    assert total_row[5] == "over_budget"
    wb.close()

    # Same-file output_file when "Variance Analysis" does not yet exist
    res_same = financial.budget_variance_analysis(fp, output_file=fp)
    assert res_same["output_file"] == fp


def test_financial_tvm_depreciation_irr_and_sensitivity(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    with pytest.raises(ValueError, match="variable_cost_per_unit must be non-negative"):
        financial.break_even_analysis(100.0, 20.0, -1.0)
    with pytest.raises(ValueError, match="when must be 'end' or 'begin'"):
        financial.calculate_fv(0.05, 5, -100, when="middle")

    with pytest.raises(ValueError, match="did not converge"):
        financial.calculate_rate(10, 100.0, 1000.0, fv=1000.0)

    with pytest.raises(ValueError, match="life must be positive"):
        financial.calculate_depreciation(1000, 100, 0)
    with pytest.raises(ValueError, match="period.*required"):
        financial.calculate_depreciation(1000, 100, 5, method="syd")
    with pytest.raises(ValueError, match="period must be between"):
        financial.calculate_depreciation(1000, 100, 5, method="ddb", period=6)

    # DDB clamping to salvage value
    ddb_res = financial.calculate_depreciation(1000, 800, 5, method="ddb", period=3)
    assert ddb_res["depreciation"] >= 0.0

    with pytest.raises(ValueError, match="at least 2"):
        financial.calculate_irr([-100])
    with pytest.raises(ValueError, match="at least one negative"):
        financial.calculate_irr([100, 200])
    with pytest.raises(ValueError, match="at least one positive"):
        financial.calculate_irr([-100, -200])
    monkeypatch.setattr(financial.npf, "irr", lambda _cfs: float("nan"))
    with pytest.raises(ValueError, match="could not converge"):
        financial.calculate_irr([-100, 200])

    # Sensitivity table: 1-variable alignment + variable not shadowed by math funcs (e.g. 'e')
    fp = _make_wb(tmp_path / "sens.xlsx")
    with pytest.raises(ValueError, match="var1_values must not be empty"):
        financial.create_sensitivity_table(fp, "Sheet1", "A1", "x * 2", "x", [])
    with pytest.raises(ValueError, match="Both var2_name and var2_values"):
        financial.create_sensitivity_table(fp, "Sheet1", "A1", "x * 2", "x", [1.0, 2.0], var2_name="y")

    sens1 = financial.create_sensitivity_table(fp, "Sheet1", "A1", "e * 10", "e", [1.0, 2.0, 3.0])
    assert sens1["table"] == [[10.0, 20.0, 30.0]]
    wb = openpyxl.load_workbook(fp)
    ws = wb["Sheet1"]
    assert [ws["B1"].value, ws["C1"].value, ws["D1"].value] == [1.0, 2.0, 3.0]
    assert [ws["B2"].value, ws["C2"].value, ws["D2"].value] == [10.0, 20.0, 30.0]
    wb.close()


def test_financial_routes_coverage(tmp_path: Path) -> None:
    fp = _make_wb(tmp_path / "rfin.xlsx")
    res = fin_routes.create_sensitivity_table(fp, "Sheet1", "A1", "x + 1", "x", [1.0, 2.0])
    assert res["table"] == [[2.0, 3.0]]

    # time_value_calc validation branches
    for op, kwargs in [
        ("fv", {}),
        ("fv", {"rate": 0.05}),
        ("fv", {"rate": 0.05, "nper": 5}),
        ("pv", {}),
        ("pv", {"rate": 0.05}),
        ("pv", {"rate": 0.05, "nper": 5}),
        ("nper", {}),
        ("nper", {"rate": 0.05}),
        ("rate", {}),
        ("rate", {"nper": 5}),
        ("depreciation", {}),
        ("depreciation", {"cost": 100.0}),
        ("depreciation", {"cost": 100.0, "salvage": 10.0}),
        ("irr", {}),
        ("unknown", {}),
    ]:
        with pytest.raises(ValueError):
            fin_routes.time_value_calc(op, **kwargs)  # type: ignore[arg-type]

    assert fin_routes.time_value_calc("fv", rate=0.05, nper=2, pmt=-100)["fv"] > 0
    assert fin_routes.time_value_calc("pv", rate=0.05, nper=2, pmt=-100)["pv"] > 0
    assert fin_routes.time_value_calc("nper", rate=0.05, pmt=-100, pv=150)["nper"] > 0
    assert fin_routes.time_value_calc("rate", nper=10, pmt=-150, pv=1000)["rate"] > 0
    assert fin_routes.time_value_calc("depreciation", cost=1000, salvage=100, life=5)["depreciation"] == 180.0
    assert fin_routes.time_value_calc("irr", cash_flows=[-100, 60, 60])["irr"] > 0


# ── Solver ──────────────────────────────────────────────────────────────────


def test_solver_edge_cases_and_guardrails(tmp_path: Path) -> None:
    assert solver._sanitize_var_name("1A") == "_1A"
    assert solver._remap_vars("B1 + B10", {}) == "B1 + B10"
    assert solver._remap_vars("B1 + B10", {"B1": "_B1"}) == "_B1 + B10"

    fp = _make_wb(tmp_path / "solver.xlsx", {"Sheet1": [[True, 500.0, "bad"]]})

    with pytest.raises(ValueError, match="lower bound"):
        solver.run_solver(fp, "Sheet1", "A1", {"A1": (10.0, 5.0)})
    with pytest.raises(ValueError, match="Unsupported constraint type"):
        solver.run_solver(fp, "Sheet1", "A1", {"A1": (0.0, 10.0)}, constraints=[{"expression": "A1", "type": "bad"}])
    with pytest.raises(ValueError, match="Chained comparisons"):
        solver.run_solver(fp, "Sheet1", "A1", {"A1": (0.0, 10.0)}, constraints=[{"expression": "0 <= A1 <= 10"}])
    with pytest.raises(ValueError, match="'!=' is not supported"):
        solver.run_solver(fp, "Sheet1", "A1", {"A1": (0.0, 10.0)}, constraints=[{"expression": "A1 != 5"}])
    with pytest.raises(ValueError, match="Invalid cell reference"):
        solver.run_solver(fp, "Sheet1", "A", {"A": (0.0, 10.0)})

    # Solve with non-numeric initial cell C1="bad" (falls back to midpoint)
    res_ok = solver.run_solver(fp, "Sheet1", "(C1 - 2)**2", {"C1": (0.0, 10.0)})
    assert res_ok["converged"] is True
    assert res_ok["found_values"]["C1"] == pytest.approx(2.0, abs=0.01)

    # Infeasible constraint -> converged=False -> workbook must NOT be modified
    res = solver.run_solver(
        fp,
        "Sheet1",
        "B1",
        {"B1": (0.0, 10.0)},
        constraints=[{"expression": "B1 >= 20"}, {"expression": "B1 <= 5"}],
    )
    assert res["converged"] is False
    wb = openpyxl.load_workbook(fp)
    assert wb["Sheet1"]["B1"].value == 500.0
    wb.close()


# ── Charts Tools & Routes ───────────────────────────────────────────────────


def test_charts_all_types_refs_and_combo_labels(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    fp = _make_wb(
        tmp_path / "charts.xlsx",
        {
            "Sheet1": [
                ["Cat", "X", "Y", "Z"],
                ["A", 1, 10, 5],
                ["B", 2, 20, 10],
                ["C", 3, 30, 15],
            ],
            "Other's": [
                ["Label"],
                ["L1"],
                ["L2"],
                ["L3"],
            ],
        },
    )

    with pytest.raises(ValueError, match="Unsupported chart type"):
        charts._make_chart("invalid")
    with pytest.raises(ValueError, match="bounded cell range"):
        charts.create_chart(fp, "Sheet1", "A:C")
    with pytest.raises(ValueError, match="Scatter chart without categories_range"):
        charts.create_chart(fp, "Sheet1", "A1:A4", chart_type="scatter")
    with pytest.raises(ValueError, match="Bubble chart requires at least 3 columns"):
        charts.create_chart(fp, "Sheet1", "A1:B4", chart_type="bubble")

    # Scatter with sheet-qualified categories_range (escaped single quote)
    charts.create_chart(fp, "Sheet1", "B1:C4", chart_type="scatter", categories_range="'Other''s'!A2:A4", title="Sc")
    # Radar, doughnut, bubble, stock
    charts.create_chart(fp, "Sheet1", "A1:C4", chart_type="radar", title="Rad")
    charts.create_chart(fp, "Sheet1", "A1:C4", chart_type="doughnut", title="Dgh")
    charts.create_chart(fp, "Sheet1", "B1:D4", chart_type="bubble", title="Bub")
    charts.create_chart(fp, "Sheet1", "A1:D4", chart_type="stock", categories_range="Sheet1!A2:A4", title="Stk")
    charts.create_chart(fp, "Sheet1", "A1:D4", chart_type="stock", title="Stk2")

    # Combo chart bounds + data labels propagated to secondary series
    with pytest.raises(ValueError, match="bar_columns index 0"):
        charts.create_combo_chart(fp, "Sheet1", "A1:D4", [0], [3])
    with pytest.raises(ValueError, match="line_columns index 0"):
        charts.create_combo_chart(fp, "Sheet1", "A1:D4", [2], [0])
    with pytest.raises(ValueError, match="x_axis_column"):
        charts.create_combo_chart(fp, "Sheet1", "A1:D4", [2], [3], x_axis_column=5)
    with pytest.raises(ValueError, match="bar_columns index"):
        charts.create_combo_chart(fp, "Sheet1", "A1:D4", [10], [3])
    with pytest.raises(ValueError, match="line_columns index"):
        charts.create_combo_chart(fp, "Sheet1", "A1:D4", [2], [10])

    charts.create_combo_chart(fp, "Sheet1", "A1:D4", [2], [3], title="Combo1")
    charts.set_chart_data_labels(fp, "Sheet1", "Combo1", show_value=True, position="t")
    wb = openpyxl.load_workbook(fp)
    combo = next(c for c in wb["Sheet1"]._charts if charts._safe_chart_title(c.title) == "Combo1")
    assert len(combo._charts) == 2
    assert combo._charts[1].dLbls is not None
    wb.close()

    # Legend invalid position should return position=None
    leg = charts.set_chart_legend(fp, "Sheet1", "Combo1", show=True, position="invalid_pos")
    assert leg["position"] is None

    # Axes with all options + categories_range
    charts.set_chart_axes(
        fp,
        "Sheet1",
        chart_index=1,
        x_title="X",
        y_title="Y",
        x_min=0,
        x_max=10,
        y_min=0,
        y_max=100,
        y_number_format="0.0%",
        log_scale_y=True,
        categories_range="Sheet1!A2:A4",
    )

    # Trendline polynomial order & moving_average period + name/forward/backward
    charts.add_chart_trendline(
        fp,
        "Sheet1",
        chart_index=0,
        trendline_type="polynomial",
        order=3,
        name="Poly3",
        periods_forward=2,
        periods_backward=1,
    )
    charts.add_chart_trendline(fp, "Sheet1", chart_index=0, trendline_type="moving_average", period=2)
    with pytest.raises(ValueError, match="order must be between 2 and 6"):
        charts.add_chart_trendline(fp, "Sheet1", chart_index=0, trendline_type="polynomial", order=9)
    with pytest.raises(ValueError, match="period must be >= 2"):
        charts.add_chart_trendline(fp, "Sheet1", chart_index=0, trendline_type="moving_average", period=1)
    with pytest.raises(ValueError, match="Series index"):
        charts.add_chart_trendline(fp, "Sheet1", chart_index=0, series_index=99)

    # Update chart with default-like values (width=15, height=10, anchor_cell="F1") via route
    chart_routes.chart("update", fp, "Sheet1", chart_index=0, title="Updated", width=15, height=10, anchor_cell="F1")

    # Delete on empty sheet via title
    with pytest.raises(ValueError, match="No charts found"):
        charts.delete_chart(fp, "Other's", chart_title="Missing")

    # _safe_chart_title edge cases
    assert charts._safe_chart_title("Plain") == "Plain"
    assert charts._safe_chart_title(SimpleNamespace(tx=SimpleNamespace(strRef=SimpleNamespace(f="Sheet1!A1")))) == "Sheet1!A1"
    assert (
        charts._safe_chart_title(
            SimpleNamespace(
                tx=SimpleNamespace(
                    strRef=None,
                    rich=SimpleNamespace(paragraphs=[SimpleNamespace(r=[SimpleNamespace(t=""), SimpleNamespace(t="Hi")])]),
                )
            )
        )
        == "Hi"
    )
    assert (
        charts._safe_chart_title(
            SimpleNamespace(
                tx=SimpleNamespace(strRef=None, rich=SimpleNamespace(paragraphs=[SimpleNamespace(r=None, runs="not_list")]))
            )
        )
        == "(untitled)"
    )

    class _BoomRich:
        @property
        def paragraphs(self) -> list:
            raise RuntimeError("boom")

    assert charts._safe_chart_title(SimpleNamespace(tx=SimpleNamespace(strRef=None, rich=_BoomRich()))) == "(untitled)"
    assert charts._safe_chart_title(SimpleNamespace(tx=None, strRef=None, rich=None)) == "(untitled)"
    assert charts._safe_chart_title(12345) == "12345"

    # add_chart_series error wrapping
    from openpyxl.chart import BarChart

    monkeypatch.setattr(BarChart, "add_data", lambda *a, **kw: (_ for _ in ()).throw(RuntimeError("bad series")))
    with pytest.raises(ValueError, match="Cannot add series"):
        charts.add_chart_series(fp, "Sheet1", 6, "A1:B2")


def test_charts_routes_validation(tmp_path: Path) -> None:
    fp = _make_wb(tmp_path / "rc.xlsx", {"Sheet1": [["A", "B"], [1, 2], [3, 4]]})
    for kwargs in [
        {"action": "create"},
        {"action": "delete"},
        {"action": "add_series"},
        {"action": "add_series", "chart_index": 0},
        {"action": "combo"},
        {"action": "combo", "data_range": "A1:B3"},
        {"action": "combo", "data_range": "A1:B3", "bar_columns": [1]},
        {"action": "data_labels"},
        {"action": "legend"},
        {"action": "bogus"},
    ]:
        with pytest.raises(ValueError):
            chart_routes.chart(file_path=fp, sheet_name="Sheet1", **kwargs)  # type: ignore[arg-type]

    assert chart_routes.chart("create", fp, "Sheet1", data_range="A1:B3", title="T1").startswith("Created")
    assert len(chart_routes.chart("list", fp, "Sheet1")) == 1
    assert chart_routes.chart("set_axes", fp, "Sheet1", chart_index=0, x_title="X").startswith("Axes")
    assert chart_routes.chart("trendline", fp, "Sheet1", chart_index=0).startswith("Added")
    assert chart_routes.chart("data_labels", fp, "Sheet1", chart_title="T1")["status"] == "ok"
    assert chart_routes.chart("legend", fp, "Sheet1", chart_title="T1")["status"] == "ok"
    assert chart_routes.chart("combo", fp, "Sheet1", data_range="A1:B3", bar_columns=[2], line_columns=[2]).startswith(
        "Created combo"
    )
    assert chart_routes.chart("delete", fp, "Sheet1", chart_index=0).startswith("Deleted")


# ── Tables ──────────────────────────────────────────────────────────────────


def test_tables_validation_resize_totals_and_convert(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    fp = _make_wb(
        tmp_path / "tbl.xlsx",
        {
            "Sheet1": [
                ["Item", "Qty", "Price"],
                ["Apple", 2, 5.0],
                ["Banana", 3, 2.0],
                [None, None, None],
            ]
        },
    )

    with pytest.raises(ValueError, match="Invalid table name"):
        tables.create_table(fp, "Sheet1", "A1:C3", "Bad Name")
    with pytest.raises(ValueError, match="bounded A1 range"):
        tables.create_table(fp, "Sheet1", "A:C", "GoodName")
    with pytest.raises(ValueError, match="reversed bounds"):
        tables.create_table(fp, "Sheet1", "C3:A1", "GoodName")

    tables.create_table(fp, "Sheet1", "A1:C3", "SalesTbl")
    with pytest.raises(ValueError, match="already exists"):
        tables.create_table(fp, "Sheet1", "E1:F3", "salestbl")
    with pytest.raises(ValueError, match="overlaps existing table"):
        tables.create_table(fp, "Sheet1", "B2:D4", "OtherTbl")

    # Resize table (expand to 4 columns with blank header fallback & duplicate header)
    with pytest.raises(ValueError, match="bounded A1 range"):
        tables.resize_table(fp, "Sheet1", "SalesTbl", "A:D")
    with pytest.raises(ValueError, match="reversed bounds"):
        tables.resize_table(fp, "Sheet1", "SalesTbl", "D3:A1")

    tables.resize_table(fp, "Sheet1", "SalesTbl", "A1:D3")
    wb = openpyxl.load_workbook(fp)
    tbl = wb["Sheet1"].tables["SalesTbl"]
    assert len(tbl.tableColumns) == 4
    assert tbl.tableColumns[3].name == "Column4"
    # Resize again with duplicate header name and blank header preserving old_cols[i].name
    wb["Sheet1"]["B1"] = "Item"
    wb["Sheet1"]["D1"] = None
    wb.save(fp)
    wb.close()
    tables.resize_table(fp, "Sheet1", "SalesTbl", "A1:D3")
    wb = openpyxl.load_workbook(fp)
    tbl = wb["Sheet1"].tables["SalesTbl"]
    assert tbl.tableColumns[1].name == "Item2"
    assert tbl.tableColumns[3].name == "Column4"
    # _get_table matching by displayName when dict key differs
    tbl.displayName = "DisplayAlias"
    assert tables._get_table(wb["Sheet1"], "DisplayAlias")[1] is tbl
    wb.close()

    # Totals row: enable, update with 'none', invalid func, get_table_data excludes totals row
    tables.set_table_totals_row(fp, "Sheet1", "SalesTbl", True, {"Item2": "sum"})
    # Calling again when totalsRowCount already 1 + clearing with "none"
    tables.set_table_totals_row(fp, "Sheet1", "SalesTbl", True, {"Item2": "none"})
    with pytest.raises(ValueError, match="Unsupported totals function"):
        tables.set_table_totals_row(fp, "Sheet1", "SalesTbl", True, {"Item2": "median"})

    data = tables.get_table_data(fp, "Sheet1", "SalesTbl")
    assert data["row_count"] == 2  # excludes totals row!

    # Outside structured ref + uncached formula preservation + cached formula replacement on convert_table_to_range
    wb = openpyxl.load_workbook(fp)
    ws = wb["Sheet1"]
    ws["F1"] = "=SUM(SalesTbl[Qty])"  # uncached (cached is None) -> must be preserved!
    ws["F2"] = "=SUM(SalesTbl[Price])"  # will have mocked cached value 7.0 -> replaced!
    wb.save(fp)
    wb.close()

    orig_load = tables.load_workbook_safe

    def _mock_load(p: str, *args: object, **kwargs: object) -> Workbook:
        loaded = orig_load(p, *args, **kwargs)  # type: ignore[arg-type]
        if kwargs.get("data_only"):
            loaded["Sheet1"]["F2"] = 7.0
        return loaded

    monkeypatch.setattr(tables, "load_workbook_safe", _mock_load)
    conv = tables.convert_table_to_range(fp, "Sheet1", "SalesTbl")
    assert conv["status"] == "ok"
    assert conv["structured_refs_converted"] == ["F2"]
    wb = openpyxl.load_workbook(fp)
    assert wb["Sheet1"]["F1"].value == "=SUM(SalesTbl[Qty])"
    assert wb["Sheet1"]["F2"].value == 7.0
    wb.close()
