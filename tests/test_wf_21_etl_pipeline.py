"""Workflow Test 21: Complex ETL Pipeline.

Validates a multi-step sequence: CSV import -> data cleaning -> deduplication -> sorting -> pivot table -> chart.
"""

from __future__ import annotations

import csv
from pathlib import Path

import pytest
from openpyxl import load_workbook

from mcp_server.tools.csv_ops import csv_to_xlsx
from mcp_server.tools.pivot_etl import deduplicate_data, create_pivot_table
from mcp_server.tools.cleaning import data_cleaner
from mcp_server.tools.analysis import sort_data
from mcp_server.tools.charts import create_chart
from mcp_server.tools.workbook import get_workbook_metadata


def test_wf_etl_pipeline(tmp_path: Path) -> None:
    # 1. Setup raw CSV data
    csv_path = tmp_path / "raw_data.csv"
    data = [
        ["ID", "Name", "Category", "Amount", "Date"],
        [1, "Alice", "Sales", 100, "2023-01-01"],
        [2, "Bob", "Marketing", 200, "2023-01-02"],
        [1, "Alice", "Sales", 100, "2023-01-01"],  # Duplicate
        [3, "Charlie", "Sales", 150, "2023-01-03"],
        [4, "Dave", "Marketing", 50, "2023-01-04"],
        [5, "Eve", "Sales", 300, "2023-01-05"],
    ]
    with open(csv_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(data)

    xlsx_path = str(tmp_path / "processed.xlsx")

    # 2. CSV to XLSX
    res_import = csv_to_xlsx(csv_path=str(csv_path), xlsx_path=xlsx_path, sheet_name="RawData")
    assert "converted" in str(res_import).lower()
    assert Path(xlsx_path).exists()

    # 3. Data Cleaning
    # data_cleaner(file_path, sheet_name, operations, ...)
    res_clean = data_cleaner(file_path=xlsx_path, sheet_name="RawData", operations=["trim_whitespace"])
    assert "operations_applied" in str(res_clean).lower()

    # 4. Deduplicate
    res_dedup = deduplicate_data(file_path=xlsx_path, sheet_name="RawData", columns=["ID", "Name"])
    assert "removed" in str(res_dedup).lower()
    
    # 5. Sort by Amount descending
    res_sort = sort_data(file_path=xlsx_path, sheet_name="RawData", column="Amount", ascending=False)
    assert "sorted" in str(res_sort).lower()

    # 6. Create Pivot Table
    # create_pivot_table(file_path, sheet_name, index_cols, value_cols, aggfunc, output_sheet, ...)
    res_pivot = create_pivot_table(
        file_path=xlsx_path,
        sheet_name="RawData",
        index_cols=["Category"],
        value_cols=["Amount"],
        aggfunc="sum",
        output_sheet="Pivot"
    )
    assert "data" in res_pivot

    # 7. Create Chart from Pivot
    res_chart = create_chart(
        xlsx_path,
        "Pivot",
        "A1:B3",
        "bar",
        "E1",
        "Sales by Category"
    )
    assert "created" in str(res_chart).lower()

    # 8. Final Verification
    meta = get_workbook_metadata(xlsx_path)
    sheet_names = [s.name for s in meta.sheets]
    assert "RawData" in sheet_names
    assert "Pivot" in sheet_names
