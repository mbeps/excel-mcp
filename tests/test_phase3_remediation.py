"""Targeted regression and coverage tests for Phase 3 remediations."""

from __future__ import annotations

import math
from pathlib import Path
from unittest.mock import patch

import openpyxl
import pytest
from openpyxl import Workbook, load_workbook

from mcp_server.routes.cleaning import csv_ops as route_csv_ops
from mcp_server.routes.cleaning import data_cleaner as route_data_cleaner
from mcp_server.routes.cleaning import parse_date_column as route_parse_date_column
from mcp_server.routes.cleaning import split_column as route_split_column
from mcp_server.routes.multi_file import multi_file as route_multi_file
from mcp_server.routes.statistical import (
    correlation_matrix as route_correlation_matrix,
)
from mcp_server.routes.statistical import (
    run_exponential_smoothing as route_run_exponential_smoothing,
)
from mcp_server.routes.statistical import (
    run_regression as route_run_regression,
)
from mcp_server.tools.analysis import (
    column_statistics,
    filter_data_advanced,
    insert_subtotals,
    profile_data,
    vlookup_helper,
)
from mcp_server.tools.cleaning import data_cleaner, parse_date_column, split_column
from mcp_server.tools.multi_file import (
    bulk_aggregate_multi_files,
    bulk_filter_multi_files,
    compare_workbooks,
    validate_data_consistency,
)
from mcp_server.tools.pivot_etl import (
    add_computed_column,
    create_pivot_table,
    deduplicate_data,
    refresh_pivot_table,
)
from mcp_server.tools.statistical import (
    correlation_matrix,
    run_exponential_smoothing,
    run_regression,
)


def _save_wb(path: Path, sheets: dict[str, list[list[object]]]) -> str:
    wb = Workbook()
    first = True
    for name, rows in sheets.items():
        ws = wb.active if first else wb.create_sheet(name)
        assert ws is not None
        ws.title = name
        first = False
        for row in rows:
            ws.append(row)
    wb.save(path)
    wb.close()
    return str(path)


# ── Cleaning ──────────────────────────────────────────────────────────────────


def test_cleaning_mixed_object_trim_and_normalize(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "mixed.xlsx",
        {"Sheet1": [["Val"], ["  Hello   World  "], [123], [None], ["  FOO  "]]},
    )
    res = data_cleaner(fp, "Sheet1", operations=["trim_whitespace", "normalize_text"])
    assert res["changes"]["trim_whitespace"] == 2
    wb = load_workbook(fp)
    ws = wb["Sheet1"]
    assert ws["A2"].value == "hello world"
    assert ws["A3"].value == 123
    assert ws["A4"].value is None
    assert ws["A5"].value == "foo"
    wb.close()


def test_cleaning_remove_empty_columns_mixed_nan_and_whitespace(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "empty_col.xlsx",
        {"Sheet1": [["Keep", "Empty"], ["a", None], ["b", "   "], ["c", ""]]},
    )
    res = data_cleaner(fp, "Sheet1", operations=["remove_empty_columns"])
    assert res["changes"]["remove_empty_columns"] == 1
    wb = load_workbook(fp)
    assert [c.value for c in wb["Sheet1"][1]] == ["Keep"]
    wb.close()


def test_cleaning_fix_numbers_scientific_and_european(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "nums.xlsx",
        {"Sheet1": [["Amt"], ["1e3"], [None], ["1.234,56"], ["$2,500.50"]]},
    )
    res = data_cleaner(fp, "Sheet1", operations=["fix_numbers", "fill_missing"], fill_missing_strategy="mean")
    assert res["changes"]["fix_numbers"] == 3
    assert res["changes"]["fill_missing"] == 1
    wb = load_workbook(fp)
    ws = wb["Sheet1"]
    assert ws["A2"].value == 1000.0
    assert isinstance(ws["A3"].value, float)
    assert ws["A4"].value == pytest.approx(1234.56)
    assert ws["A5"].value == pytest.approx(2500.50)
    wb.close()


def test_cleaning_fill_missing_all_nan_does_not_increment_count(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "all_nan.xlsx",
        {"Sheet1": [["Num", "Flag"], [None, True], [None, False]]},
    )
    res = data_cleaner(fp, "Sheet1", operations=["fill_missing"], fill_missing_strategy="mean", preview=True)
    assert res["changes"]["fill_missing"] == 0


def test_cleaning_preserves_banner_rows_and_other_sheets(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "banner.xlsx",
        {
            "Sheet1": [["Banner Title"], ["Subtitle"], ["Name"], ["  Alice  "]],
            "KeepSheet": [["X"], [42]],
        },
    )
    non_canonical = str(tmp_path / "." / "banner.xlsx")
    data_cleaner(non_canonical, "Sheet1", operations=["trim_whitespace"], header_row=3)
    wb = load_workbook(fp)
    assert "KeepSheet" in wb.sheetnames
    assert wb["Sheet1"]["A1"].value == "Banner Title"
    assert wb["Sheet1"]["A2"].value == "Subtitle"
    assert wb["Sheet1"]["A4"].value == "Alice"
    wb.close()


def test_split_column_preserves_nan(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "split.xlsx",
        {"Sheet1": [["Full"], ["a,b"], [None]]},
    )
    split_column(fp, "Sheet1", "Full", delimiter=",", new_columns=["First", "Second"])
    wb = load_workbook(fp)
    ws = wb["Sheet1"]
    assert ws["A2"].value == "a"
    assert ws["B2"].value == "b"
    assert ws["A3"].value is None
    assert ws["B3"].value is None
    wb.close()


def test_parse_date_column_guards_bare_numbers_and_preserves_invalid(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "dates.xlsx",
        {"Sheet1": [["Date"], ["2024-01-15"], ["5"], [None], ["2024"], [2024], ["not-a-date"]]},
    )
    res = parse_date_column(fp, "Sheet1", "Date")
    assert res["parsed_count"] == 1
    assert res["failed_count"] == 5
    wb = load_workbook(fp)
    ws = wb["Sheet1"]
    assert ws["A2"].value == "2024-01-15"
    assert ws["A3"].value == "5"
    assert ws["A4"].value is None
    assert ws["A5"].value == "2024"
    assert ws["A6"].value == 2024
    assert ws["A7"].value == "not-a-date"
    wb.close()


def test_cleaning_routes_dispatch(tmp_path: Path) -> None:
    fp = _save_wb(tmp_path / "route_clean.xlsx", {"Sheet1": [["Full", "D"], ["a-b", "2024-05-01"]]})
    csv_fp = str(tmp_path / "sample.csv")
    Path(csv_fp).write_text("A,B\n1,2\n", encoding="utf-8")

    assert route_split_column(fp, "Sheet1", "Full", delimiter="-")["rows_affected"] == 1
    assert route_data_cleaner(fp, "Sheet1", operations=["trim_whitespace"], preview=True)["preview_only"] is True
    assert route_parse_date_column(fp, "Sheet1", "D")["parsed_count"] == 1
    assert route_csv_ops("preview", file_path=csv_fp)["total_rows"] == 1

    xlsx_out = str(tmp_path / "from_csv.xlsx")
    route_csv_ops("to_xlsx", csv_path=csv_fp, xlsx_path=xlsx_out)
    assert Path(xlsx_out).exists()

    csv_out = str(tmp_path / "to_csv.csv")
    route_csv_ops("to_csv", file_path=xlsx_out, output_path=csv_out)
    assert Path(csv_out).exists()

    with pytest.raises(ValueError, match="file_path is required"):
        route_csv_ops("preview")
    with pytest.raises(ValueError, match="csv_path"):
        route_csv_ops("to_xlsx")
    with pytest.raises(ValueError, match="xlsx_path"):
        route_csv_ops("to_xlsx", csv_path=csv_fp)
    with pytest.raises(ValueError, match="file_path is required"):
        route_csv_ops("to_csv")
    with pytest.raises(ValueError, match="output_path is required"):
        route_csv_ops("to_csv", file_path=xlsx_out)
    with pytest.raises(ValueError, match="Unknown action"):
        route_csv_ops("bogus")  # type: ignore[arg-type]


# ── Analysis ──────────────────────────────────────────────────────────────────


def test_analysis_id_age_qty_not_hijacked_and_filter_regex(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "analysis.xlsx",
        {"Sheet1": [["ID", "Age", "Qty", "Note"], [1, 30, 5, "a+b"], [2, 25, 10, None]]},
    )
    stats_res = column_statistics(fp, "Sheet1", "ID")
    assert stats_res.column == "ID"
    assert stats_res.count == 2

    filt = filter_data_advanced(
        fp,
        "Sheet1",
        [{"column": "Note", "operator": "contains", "value": "a+b"}],
    )
    assert filt["rows"] == 1

    filt_sw = filter_data_advanced(
        fp,
        "Sheet1",
        [{"column": "Note", "operator": "startswith", "value": "n"}],
    )
    assert filt_sw["rows"] == 0

    with pytest.raises(ValueError, match="Cannot compare"):
        filter_data_advanced(fp, "Sheet1", [{"column": "Note", "operator": ">", "value": 5}])


def test_analysis_bool_excluded_and_nan_stats_none(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "stats_edge.xlsx",
        {"Sheet1": [["Flag", "Single", "Const"], [True, 5, 10], [False, None, 10], [True, None, 10]]},
    )
    flag_stats = column_statistics(fp, "Sheet1", "Flag")
    assert "not numeric" in (flag_stats.message or "")

    single_stats = column_statistics(fp, "Sheet1", "Single")
    assert single_stats.std is None

    const_stats = column_statistics(fp, "Sheet1", "Const")
    assert const_stats.skewness is None
    assert const_stats.kurtosis is None

    prof = profile_data(fp, "Sheet1", data_range="A:B")
    assert prof["column_count"] == 2
    assert "min" not in prof["columns"][0]  # bool not numeric
    assert prof["columns"][1]["std"] is None


def test_analysis_insert_subtotals_same_col_and_dropna_false(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "sub.xlsx",
        {"Sheet1": [["Val"], [10], [None], [20]]},
    )
    res = insert_subtotals(fp, "Sheet1", group_col="Val", value_col="Val")
    assert res["groups"] == 3
    wb = load_workbook(fp)
    ws = wb["Sheet1"]
    formulas = [ws.cell(r, 1).value for r in range(1, ws.max_row + 1)]
    assert any(isinstance(v, str) and v.startswith("=SUBTOTAL(9,") for v in formulas)
    wb.close()


def test_analysis_vlookup_bool_vs_int_and_data_only(tmp_path: Path) -> None:
    lookup_fp = _save_wb(tmp_path / "lookup.xlsx", {"Sheet1": [["Key"], [True], [1]]})
    data_fp = _save_wb(tmp_path / "data.xlsx", {"Sheet1": [["Key", "Val"], [1, "one"], [True, "yes"]]})
    out_fp = str(tmp_path / "vout.xlsx")

    res = vlookup_helper(lookup_fp, data_fp, "Key", "Key", ["Val"], output_file=out_fp)
    assert res["matched"] == 2
    assert res["results"][0]["return_values"]["Val"] == "yes"
    assert res["results"][1]["return_values"]["Val"] == "one"
    assert Path(out_fp).exists()


# ── Pivot & ETL ───────────────────────────────────────────────────────────────


def test_pivot_date_freq_refresh_and_output_file_preserves_sheets(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "pivot.xlsx",
        {"Sales": [["Date", "Amount"], ["2024-01-05", 10], ["2024-01-20", 20], ["2024-02-05", 30]]},
    )
    out_fp = _save_wb(tmp_path / "out_pivot.xlsx", {"Existing": [["Keep"], [1]]})

    create_pivot_table(
        fp,
        "Sales",
        index_cols=["Date"],
        value_cols=["Amount"],
        output_sheet="Monthly",
        output_file=out_fp,
        date_freq="ME",
    )
    out_wb = load_workbook(out_fp)
    assert "Existing" in out_wb.sheetnames
    assert "Monthly" in out_wb.sheetnames
    out_wb.close()

    # Also test persisted date_freq in same workbook + refresh
    create_pivot_table(
        fp,
        "Sales",
        index_cols=["Date"],
        value_cols=["Amount"],
        output_sheet="MonthlyLocal",
        date_freq="ME",
    )
    refreshed = refresh_pivot_table(fp, "MonthlyLocal")
    assert refreshed["rows"] == 2


def test_add_computed_column_ast_fallback_overwrite_and_rolling_validation(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "calc.xlsx",
        {"Sheet1": [["X", "Y"], [2, 10], [-3, 10]]},
    )
    # Overwrite existing column X using AST fallback (ternary IfExp not supported by numexpr)
    msg = add_computed_column(fp, "Sheet1", "X", "X * 2 if X > 0 else 0")
    assert "mode=formula" in str(msg)
    wb = load_workbook(fp)
    ws = wb["Sheet1"]
    assert ws["A2"].value == 4
    assert ws["A3"].value == 0
    assert ws["B2"].value == 10
    wb.close()

    with pytest.raises(ValueError, match="window must be an integer >= 1"):
        add_computed_column(fp, "Sheet1", "Roll", "", column_type="rolling", source_col="X", window=0)


def test_deduplicate_keep_none_and_false(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "dedup.xlsx",
        {"Sheet1": [["Name"], ["a"], ["a"], ["b"]]},
    )
    res = deduplicate_data(fp, "Sheet1", keep="none")
    assert "Removed 2" in res
    wb = load_workbook(fp)
    assert wb["Sheet1"]["A2"].value == "b"
    assert wb["Sheet1"]["A3"].value is None
    wb.close()

    with pytest.raises(ValueError, match="Invalid keep value"):
        deduplicate_data(fp, "Sheet1", keep=True)


# ── Multi-file ────────────────────────────────────────────────────────────────


def test_multi_file_empty_mean_bool_filter_and_compare_edge_cases(tmp_path: Path) -> None:
    f1 = _save_wb(
        tmp_path / "f1.xlsx",
        {"Sheet1": [["ID", "Val", "Tag"], [1, True, "a+b"], [1, False, None]]},
    )
    f2 = _save_wb(
        tmp_path / "f2.xlsx",
        {"Sheet1": [["ID", "Val", "Tag"], [1, 99, "other"]]},
    )

    agg = bulk_aggregate_multi_files([f1], "Val", operation="mean")
    assert agg["per_file"][0]["value"] is None
    assert agg["aggregate"] is None

    filt = bulk_filter_multi_files([f1], "Tag", "contains", "a+b")
    assert filt["total_matched"] == 1

    filt_eq = bulk_filter_multi_files([f2], "Val", "equals", "99.0")
    assert filt_eq["total_matched"] == 1

    val = validate_data_consistency([f1, f2], key_column="ID", check_columns=["Val"])
    assert val["consistent"] is False
    assert any(m.get("issue") == "dtype_mismatch" for m in val["schema_mismatches"])

    # Compare: True vs 1 should differ; NaN vs NaN (empty) should not; formulas compared via route
    wb_a = _save_wb(tmp_path / "ca.xlsx", {"S1": [["=1+1", True]]})
    wb_b = _save_wb(tmp_path / "cb.xlsx", {"S1": [["=2+0", 1]]})
    cmp_res = route_multi_file("compare", file_a=wb_a, file_b=wb_b, compare_values=False, compare_formulas=True)
    assert cmp_res["total_differences"] == 2
    assert math.isnan(float("nan"))  # sanity


def test_compare_workbooks_nan_equality(tmp_path: Path) -> None:
    wb_a = _save_wb(tmp_path / "na.xlsx", {"Sheet1": [[None, 1]]})
    wb_b = _save_wb(tmp_path / "nb.xlsx", {"Sheet1": [[None, 1]]})
    res = compare_workbooks(wb_a, wb_b)
    assert res["identical"] is True


# ── Statistical ───────────────────────────────────────────────────────────────


def test_statistical_smoothing_alignment_validation_and_corr_new_file(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "stat.xlsx",
        {"Data": [["X", "Y", "Flag"], [1, 10, True], [2, None, False], [3, 20, True], [4, 30, False]]},
    )
    out_fp = str(tmp_path / "stat_out.xlsx")
    res = run_exponential_smoothing(fp, "Data", "Y", alpha=0.5, method="holt", output_file=out_fp)
    assert res["rows_written"] == 4
    wb = load_workbook(out_fp)
    ws = wb["Data"]
    # Row 3 (index 1 in data) was None so smoothed value at row 3 must be None, row 4 must be populated
    assert ws["D3"].value is None
    assert ws["D4"].value is not None
    wb.close()

    with pytest.raises(ValueError, match="forecast_steps"):
        run_exponential_smoothing(fp, "Data", "Y", forecast_steps=-1)
    with pytest.raises(ValueError, match="seasonal_periods"):
        run_exponential_smoothing(fp, "Data", "Y", seasonal_periods=1)
    with pytest.raises(ValueError, match="smoothing_trend"):
        run_exponential_smoothing(fp, "Data", "Y", smoothing_trend=1.5)
    with pytest.raises(ValueError, match="smoothing_seasonal"):
        run_exponential_smoothing(fp, "Data", "Y", smoothing_seasonal=-0.1)

    corr_out = str(tmp_path / "corr_new.xlsx")
    corr_res = correlation_matrix(fp, "Data", output_sheet="Corr", output_file=corr_out)
    assert "Flag" not in corr_res["columns"]
    assert Path(corr_out).exists()


def test_statistical_constant_y_and_numpy_fallback(tmp_path: Path) -> None:
    const_fp = _save_wb(tmp_path / "const_y.xlsx", {"Data": [["X", "Y"], [1, 5], [2, 5], [3, 5], [4, 5]]})
    with pytest.raises(ValueError, match="zero variance"):
        run_regression(const_fp, "Data", "Y", ["X"])

    reg_fp = _save_wb(tmp_path / "reg.xlsx", {"Data": [["X", "Y"], [1, 3], [2, 5], [3, 7], [4, 9]]})
    with patch("mcp_server.tools.statistical._HAS_STATSMODELS", False):
        res = run_regression(reg_fp, "Data", "Y", ["X"])
        assert res["r_squared"] == pytest.approx(1.0)


def test_statistical_routes_not_marked_read_only() -> None:
    class _FakeMCP:
        def __init__(self) -> None:
            self.registered: dict[str, object] = {}

        def tool(self, annotations: object = None):
            def _dec(fn):
                self.registered[fn.__name__] = annotations
                return fn

            return _dec

    from mcp_server.routes.statistical import register

    fake = _FakeMCP()
    register(fake)
    assert fake.registered["run_regression"] is None
    assert fake.registered["run_exponential_smoothing"] is None
    assert fake.registered["correlation_matrix"] is None
    assert route_run_regression is not None
    assert route_run_exponential_smoothing is not None
    assert route_correlation_matrix is not None
