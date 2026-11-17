"""Coverage-completion tests for Phase 3 routes and tools."""

from __future__ import annotations

import ast
from pathlib import Path
from unittest.mock import patch

import pandas as pd
import pytest
from openpyxl import Workbook

from mcp_server.routes.analysis import value_counts as route_value_counts
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
    _resolve_column_index,
    aggregate_data,
    filter_data_advanced,
    find_duplicates,
    sort_data,
    value_counts,
    vlookup_helper,
)
from mcp_server.tools.cleaning import (
    _collect_sample_changes,
    _normalize_text,
    _parse_numeric_string,
    _trim_whitespace,
    data_cleaner,
    parse_date_column,
)
from mcp_server.tools.multi_file import (
    _cells_equal,
    _try_float,
    bulk_aggregate_multi_files,
    bulk_filter_multi_files,
    compare_workbooks,
    validate_data_consistency,
)
from mcp_server.tools.pivot_etl import (
    _eval_ast_node,
    add_computed_column,
    create_pivot_table,
    deduplicate_data,
    merge_datasets,
    unpivot_data,
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


# ── Routes ────────────────────────────────────────────────────────────────────


def test_phase3_routes_missing_branches(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "r.xlsx",
        {"Sheet1": [["X", "Y"], [1, 2], [2, 4], [3, 6], [4, 8]]},
    )
    assert route_value_counts(fp, "Sheet1", "X")["total_rows"] == 4
    assert route_run_regression(fp, "Sheet1", "Y", ["X"])["n_observations"] == 4
    assert route_run_exponential_smoothing(fp, "Sheet1", "Y")["rows_written"] == 4
    assert len(route_correlation_matrix(fp, "Sheet1")["columns"]) == 3

    with pytest.raises(ValueError, match="file_paths is required"):
        route_multi_file("aggregate")
    with pytest.raises(ValueError, match="column is required"):
        route_multi_file("aggregate", file_paths=[fp])
    with pytest.raises(ValueError, match="file_paths is required"):
        route_multi_file("filter")
    with pytest.raises(ValueError, match="column is required"):
        route_multi_file("filter", file_paths=[fp])
    with pytest.raises(ValueError, match="operator is required"):
        route_multi_file("filter", file_paths=[fp], column="X")
    with pytest.raises(ValueError, match="value is required"):
        route_multi_file("filter", file_paths=[fp], column="X", operator="==")
    with pytest.raises(ValueError, match="file_paths is required"):
        route_multi_file("validate")
    with pytest.raises(ValueError, match="key_column is required"):
        route_multi_file("validate", file_paths=[fp])
    with pytest.raises(ValueError, match="file_a is required"):
        route_multi_file("compare")
    with pytest.raises(ValueError, match="file_b is required"):
        route_multi_file("compare", file_a=fp)
    with pytest.raises(ValueError, match="Unknown action"):
        route_multi_file("bogus")  # type: ignore[arg-type]


# ── Cleaning ──────────────────────────────────────────────────────────────────


def test_cleaning_remaining_branches(tmp_path: Path) -> None:
    # Object column with no strings
    df_obj = pd.DataFrame({"A": pd.Series([1, None, 2], dtype=object)})
    _, c1 = _trim_whitespace(df_obj.copy(), ["A"])
    _, c2 = _normalize_text(df_obj.copy(), ["A"])
    assert c1 == 0 and c2 == 0

    assert _parse_numeric_string("$   ") is None
    assert _parse_numeric_string("12,34") == pytest.approx(12.34)

    # ffill, bfill, median, mode, unknown strategy
    fp = _save_wb(
        tmp_path / "fill.xlsx",
        {"Sheet1": [["A", "B", "C"], [1, None, "x"], [None, 10, None], [3, 20, "x"]]},
    )
    assert data_cleaner(fp, "Sheet1", operations=["fill_missing"], fill_missing_strategy="ffill", preview=True)[
        "changes"
    ]["fill_missing"] >= 1
    assert data_cleaner(fp, "Sheet1", operations=["fill_missing"], fill_missing_strategy="bfill", preview=True)[
        "changes"
    ]["fill_missing"] >= 1
    assert data_cleaner(fp, "Sheet1", operations=["fill_missing"], fill_missing_strategy="median", preview=True)[
        "changes"
    ]["fill_missing"] >= 1
    assert data_cleaner(fp, "Sheet1", operations=["fill_missing"], fill_missing_strategy="mode", preview=True)[
        "changes"
    ]["fill_missing"] >= 1
    with pytest.raises(ValueError, match="Unknown fill_missing_strategy"):
        data_cleaner(fp, "Sheet1", operations=["fill_missing"], fill_missing_strategy="invalid", preview=True)

    # _collect_sample_changes hitting max_samples limit
    orig = pd.DataFrame({"A": [f"x{i}" for i in range(15)], "B": [f"y{i}" for i in range(15)]})
    clean = pd.DataFrame({"A": [f"z{i}" for i in range(15)], "B": [f"w{i}" for i in range(15)]})
    assert len(_collect_sample_changes(orig, clean, max_samples=5)) == 5

    # target becoming empty after remove_empty_columns
    fp_empty = _save_wb(tmp_path / "drop_target.xlsx", {"Sheet1": [["A", "B"], [1, None], [2, None]]})
    res_drop = data_cleaner(fp_empty, "Sheet1", operations=["remove_empty_columns", "trim_whitespace"], columns=["B"])
    assert res_drop["changes"]["trim_whitespace"] == 0

    # parse_date_column with output_column and unparseable value
    fp_date = _save_wb(tmp_path / "d_out.xlsx", {"Sheet1": [["D"], ["2024-01-01"], ["bad"]]})
    res_d = parse_date_column(fp_date, "Sheet1", "D", output_column="Parsed")
    assert res_d["failed_count"] == 1


# ── Analysis ──────────────────────────────────────────────────────────────────


def test_analysis_remaining_branches(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "a.xlsx",
        {"Sheet1": [["Name", "Val"], ["Alice", 10], ["Bob", 20], ["Charlie", 30]]},
    )
    # filter_data_advanced with output_sheet
    res_f = filter_data_advanced(
        fp,
        "Sheet1",
        [{"column": "Val", "operator": ">", "value": 10}],
        output_sheet="Filtered",
    )
    assert res_f["rows"] == 2

    # sort_data with neither sort_by nor column
    with pytest.raises(ValueError, match="Provide either"):
        sort_data(fp, "Sheet1")

    # aggregate_data invalid op
    with pytest.raises(ValueError, match="Unsupported operation"):
        aggregate_data(fp, "Sheet1", "Name", "Val", operation="bogus")

    # find_duplicates missing col
    with pytest.raises(ValueError, match="not found"):
        find_duplicates(fp, "Sheet1", ["Missing"])

    # _resolve_column_index case-insensitive + out-of-bounds column letter
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.append(["Name", "Val"])
    assert _resolve_column_index(ws, 1, "name") == 1
    assert _resolve_column_index(ws, 1, "B") == 2
    with pytest.raises(ValueError, match="not found"):
        _resolve_column_index(ws, 1, "ZZ")
    with pytest.raises(ValueError, match="not found"):
        _resolve_column_index(ws, 1, "ZZZ")
    wb.close()

    # vlookup_helper with fuzzy=True and output_file
    lookup_fp = _save_wb(tmp_path / "l.xlsx", {"Sheet1": [["K"], ["a"], ["missing"]]})
    data_fp = _save_wb(tmp_path / "d.xlsx", {"Sheet1": [["K", None], ["a", 100], [None, 200]]})
    v_out = str(tmp_path / "v_fuzzy.xlsx")
    v_res = vlookup_helper(lookup_fp, data_fp, "K", "K", ["B"], fuzzy=True, fuzzy_threshold=0.5, output_file=v_out)
    assert v_res["matched"] == 1

    # value_counts on empty sheet
    empty_fp = _save_wb(tmp_path / "empty_vc.xlsx", {"Sheet1": [["A"]]})
    with pytest.raises(ValueError, match="empty"):
        value_counts(empty_fp, "Sheet1", "A")


# ── Pivot & ETL ───────────────────────────────────────────────────────────────


def test_pivot_etl_remaining_branches(tmp_path: Path) -> None:
    df_test = pd.DataFrame({"X": [1, 2], "Y": [3, 4]})
    # _eval_ast_node branches: UnaryOp, BinOp, BoolOp, Call, unknown name, unsupported node
    assert list(_eval_ast_node(ast.parse("-X + +Y", mode="eval"), df_test)) == [2, 2]
    assert list(_eval_ast_node(ast.parse("(X > 0) and (Y > 3)", mode="eval"), df_test)) == [False, True]
    assert list(_eval_ast_node(ast.parse("abs(-X)", mode="eval"), df_test)) == [1, 2]
    with pytest.raises(ValueError, match="Unknown column"):
        _eval_ast_node(ast.parse("Z + 1", mode="eval"), df_test)
    with pytest.raises(ValueError, match="Unsupported AST"):
        _eval_ast_node(ast.parse("[1, 2]", mode="eval"), df_test)

    fp = _save_wb(
        tmp_path / "p.xlsx",
        {"Sheet1": [["Cat", "D", "V"], ["A", "not-a-date", 10], ["B", "bad", 20]]},
    )
    with pytest.raises(ValueError, match="not found"):
        create_pivot_table(fp, "Sheet1", ["Missing"], ["V"])
    # date_freq with unparseable column falls back to column grouping
    res_p = create_pivot_table(fp, "Sheet1", ["D"], ["V"], date_freq="ME")
    assert len(res_p["data"]) == 0 or res_p["index_columns"] == ["D"]

    # date_freq with valid date column + unparseable non-date column in index
    fp_d = _save_wb(
        tmp_path / "pd.xlsx",
        {"Sheet1": [["D", "V"], ["2024-01-01", 10], ["2024-02-01", 20]]},
    )
    with patch("mcp_server.tools.pivot_etl.pd.to_datetime", side_effect=RuntimeError("boom")):
        res_fallback = create_pivot_table(fp_d, "Sheet1", ["D"], ["V"], date_freq="ME")
        assert len(res_fallback["data"]) == 2

    with pytest.raises(ValueError, match="not found"):
        unpivot_data(fp, "Sheet1", ["Missing"], ["V"])

    fp_m = _save_wb(
        tmp_path / "m.xlsx",
        {"S1": [["K", "A"], [1, 10]], "S2": [["K", "B"], [1, 20]]},
    )
    with pytest.raises(ValueError, match="not found"):
        merge_datasets(fp_m, "S1", "S2", left_on="Missing", right_on="K")
    with pytest.raises(ValueError, match="not found"):
        merge_datasets(fp_m, "S1", "S2", join_key="Missing")

    with pytest.raises(ValueError, match="Failed to evaluate"):
        add_computed_column(fp, "Sheet1", "Z", "UnknownCol + 1")

    with pytest.raises(ValueError, match="not found"):
        deduplicate_data(fp, "Sheet1", columns=["Missing"])
    assert "Removed 0" in deduplicate_data(fp, "Sheet1", keep=False)
    with pytest.raises(ValueError, match="Invalid keep value"):
        deduplicate_data(fp, "Sheet1", keep="bogus")

    # df.eval raising ValueError (non-NotImplementedError)
    with patch.object(pd.DataFrame, "eval", side_effect=RuntimeError("eval boom")):
        with pytest.raises(ValueError, match="Failed to evaluate"):
            add_computed_column(fp, "Sheet1", "Z", "V + 1")


# ── Multi-file ────────────────────────────────────────────────────────────────


def test_multi_file_remaining_branches(tmp_path: Path) -> None:
    f1 = _save_wb(
        tmp_path / "mf1.xlsx",
        {"Sheet1": [["ID", "Num", "Txt", "Dt"], [1, 10, "a", "2024-01-01"], [2, None, "b", "2024-01-02"]]},
    )
    f2 = _save_wb(
        tmp_path / "mf2.xlsx",
        {"Sheet1": [["ID", "Num", "Txt"], [1, 10, "a"], [2, 20, "c"]]},
    )
    f_empty = _save_wb(tmp_path / "mf_empty.xlsx", {"Sheet1": [["ID", "Num"], [1, "bad"]]})

    # empty min/max/count
    assert bulk_aggregate_multi_files([f_empty], "Num", operation="min")["aggregate"] is None
    assert bulk_aggregate_multi_files([f_empty], "Num", operation="max")["aggregate"] is None
    assert bulk_aggregate_multi_files([f_empty], "Num", operation="count")["aggregate"] == 0

    # filter operators: not_equals (numeric & string), less_than_or_equal
    assert bulk_filter_multi_files([f1], "Num", "not_equals", 10)["total_matched"] == 1
    assert bulk_filter_multi_files([f1], "Txt", "not_equals", "a")["total_matched"] == 1
    assert bulk_filter_multi_files([f1], "Num", "less_than_or_equal", 10)["total_matched"] == 1
    with pytest.raises(ValueError, match="not found"):
        bulk_filter_multi_files([f1], "Missing", "equals", 1)

    # validate_data_consistency: empty file_paths, missing key_column, missing check_columns,
    # missing_columns without check_columns, missing_keys, mismatched_values, bool/datetime dtypes
    with pytest.raises(ValueError, match="file_paths must not be empty"):
        validate_data_consistency([], "ID")
    f_nokey = _save_wb(tmp_path / "nokey.xlsx", {"Sheet1": [["Other"], [1]]})
    v1 = validate_data_consistency([f1, f_nokey], "ID", check_columns=["Dt"])
    assert any(m.get("issue") == "missing_key_column" for m in v1["schema_mismatches"])
    assert any(m.get("issue") == "missing_columns" for m in v1["schema_mismatches"])

    f_bool = _save_wb(tmp_path / "fbool.xlsx", {"Sheet1": [["ID", "Flag"], [1, True], [3, False]]})
    f_bool2 = _save_wb(tmp_path / "fbool2.xlsx", {"Sheet1": [["ID", "Flag"], [1, False]]})
    v2 = validate_data_consistency([f_bool, f_bool2], "ID", check_columns=["Flag"])
    assert v2["missing_keys"]
    assert v2["mismatched_values"]

    # compare_workbooks: sheet_name_b missing, both same sheet_name missing, diff_found with output_file
    with pytest.raises(ValueError, match="not found in workbook A"):
        compare_workbooks(f1, f2, sheet_name="Nope", sheet_name_b="Sheet1")
    with pytest.raises(ValueError, match="not found in workbook B"):
        compare_workbooks(f1, f2, sheet_name="Sheet1", sheet_name_b="Nope")
    with pytest.raises(ValueError, match="not found in both"):
        compare_workbooks(f1, f2, sheet_name="Nope")

    with pytest.raises(ValueError, match="file_paths must not be empty"):
        bulk_aggregate_multi_files([], "Num")
    with pytest.raises(ValueError, match="file_paths must not be empty"):
        bulk_filter_multi_files([], "Num", "equals", 1)
    assert _try_float(True) is None
    assert _cells_equal(float("nan"), float("nan")) is True

    cmp_out = str(tmp_path / "cmp_out.xlsx")
    cmp_res = compare_workbooks(f1, f2, output_file=cmp_out, compare_values=True, compare_formulas=True)
    assert cmp_res["total_differences"] >= 1
    assert Path(cmp_out).exists()

    # datetime + text coarse dtype and NaN key in validate_data_consistency
    from datetime import datetime

    f_dt1 = _save_wb(tmp_path / "dt1.xlsx", {"Sheet1": [["ID", "Dt", "Txt"], [1, datetime(2024, 1, 1), "a"], [None, datetime(2024, 1, 2), "b"]]})
    f_dt2 = _save_wb(tmp_path / "dt2.xlsx", {"Sheet1": [["ID", "Dt", "Txt"], [1, datetime(2024, 1, 1), "a"]]})
    v_dt = validate_data_consistency([f_dt1, f_dt2], "ID", check_columns=["Dt", "Txt"])
    assert v_dt["consistent"] is True


# ── Statistical ───────────────────────────────────────────────────────────────


def test_statistical_remaining_branches(tmp_path: Path) -> None:
    fp = _save_wb(
        tmp_path / "s.xlsx",
        {"Data": [["X", "Y"], [1, 2], [2, 4], [3, 6], [4, 8], [5, 10], [6, 12]]},
    )
    # run_regression with output_file == file_path (twice to hit existing output_sheet)
    res_same = run_regression(fp, "Data", "Y", ["X"], output_file=fp)
    assert res_same["n_observations"] == 6
    run_regression(fp, "Data", "Y", ["X"], output_file=fp)
    run_regression(fp, "Data", "Y", ["X"])
    run_regression(fp, "Data", "Y", ["X"])

    with pytest.raises(ValueError, match="method must be one of"):
        run_exponential_smoothing(fp, "Data", "Y", method="bad")

    fp_short = _save_wb(tmp_path / "short.xlsx", {"Data": [["Y"], [1], [2]]})
    with pytest.raises(ValueError, match="Not enough valid data points"):
        run_exponential_smoothing(fp_short, "Data", "Y")

    with pytest.raises(ValueError, match="Need at least"):
        run_exponential_smoothing(fp, "Data", "Y", method="holt_winters", seasonal_periods=4)

    res_hw = run_exponential_smoothing(
        fp,
        "Data",
        "Y",
        alpha=0.5,
        method="holt",
        smoothing_trend=0.2,
        output_file=fp,
    )
    assert res_hw["rows_written"] == 6
    # Run again with same output_column to hit existing header match
    run_exponential_smoothing(fp, "Data", "Y", alpha=0.5, method="holt", smoothing_trend=0.2, output_file=fp)

    # _HAS_STATSMODELS=False fallback for holt
    with patch("mcp_server.tools.statistical._HAS_STATSMODELS", False):
        res_fb = run_exponential_smoothing(fp, "Data", "Y", method="holt")
        assert res_fb["method"] == "simple"

    # holt_winters with smoothing_seasonal provided
    fp_long = _save_wb(
        tmp_path / "long.xlsx",
        {"Data": [["Y"], *[[i * 2 + (i % 3)] for i in range(1, 13)]]},
    )
    res_hw2 = run_exponential_smoothing(
        fp_long,
        "Data",
        "Y",
        alpha=0.4,
        method="holt_winters",
        seasonal_periods=3,
        smoothing_trend=0.2,
        smoothing_seasonal=0.2,
    )
    assert res_hw2["method"] == "holt_winters"
