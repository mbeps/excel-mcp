"""Core route dispatch and tool integration tests.

Verifies that the high-level route entry points correctly dispatch to tools
and handle both success and error paths.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from mcp_server.models.analysis import FilterCondition
from mcp_server.routes.analysis import filter_data_advanced, sort_data, profile_data, aggregate_data
from mcp_server.routes.metadata import hyperlink, scenario
from mcp_server.routes.financial import time_value_calc
from mcp_server.routes.governance import doc_properties
from mcp_server.routes.worksheet_ops import worksheet_transfer
from mcp_server.routes.workbook import sheet_management, create_workbook
from mcp_server.routes.cell_ops import write_cells


@pytest.fixture
def sample_xlsx(tmp_path: Path) -> str:
    fp = str(tmp_path / "test.xlsx")
    create_workbook(file_path=fp, sheet_name="Sheet1")
    write_cells(
        mode="range",
        file_path=fp,
        sheet_name="Sheet1",
        start_cell="A1",
        data=[["Name", "Age", "City"], ["Alice", 30, "London"], ["Bob", 25, "Paris"], ["Charlie", 35, "Berlin"]]
    )
    return fp


# ---------------------------------------------------------------------------
# 1. Analysis Routes
# ---------------------------------------------------------------------------

class TestAnalysisRoutes:
    def test_filter_dispatch(self, sample_xlsx: str) -> None:
        res = filter_data_advanced(
            file_path=sample_xlsx,
            sheet_name="Sheet1",
            conditions=[FilterCondition(column="Age", operator=">", value=28)]
        )
        assert res["rows"] == 2
        assert len(res["data"]) == 2

    def test_sort_dispatch(self, sample_xlsx: str) -> None:
        res = sort_data(
            file_path=sample_xlsx,
            sheet_name="Sheet1",
            column="Age",
            ascending=True
        )
        assert "sorted" in str(res).lower()

    def test_profile_dispatch(self, sample_xlsx: str) -> None:
        res = profile_data(file_path=sample_xlsx, sheet="Sheet1")
        assert res["status"] == "success"
        assert res["row_count"] == 3

    def test_aggregate_dispatch(self, sample_xlsx: str) -> None:
        res = aggregate_data(
            file_path=sample_xlsx,
            sheet_name="Sheet1",
            group_by="City",
            value_column="Age",
            operation="mean"
        )
        assert len(res["groups"]) == 3


# ---------------------------------------------------------------------------
# 2. Metadata & Scenario Routes
# ---------------------------------------------------------------------------

class TestScenarioRoute:
    def test_scenario_lifecycle(self, sample_xlsx: str) -> None:
        fp = sample_xlsx
        # Create
        res_add = scenario(
            action="add",
            file_path=fp,
            name="TestScenario",
            description="Test",
            cell_values={"Sheet1": {"A2": "Modified"}}
        )
        assert "saved" in str(res_add).lower()

        # List
        res_list = scenario(action="list", file_path=fp)
        assert len(res_list) >= 1

        # Apply
        res_apply = scenario(action="apply", file_path=fp, name="TestScenario")
        assert res_apply["scenario"] == "TestScenario"
        assert res_apply["cells_updated"] >= 1


class TestHyperlinkRoute:
    def test_hyperlink_lifecycle(self, sample_xlsx: str) -> None:
        fp = sample_xlsx
        # Add
        res_add = hyperlink(
            action="add",
            file_path=fp,
            sheet_name="Sheet1",
            cell_ref="D1",
            url="https://google.com",
            tooltip="Search"
        )
        assert "added" in str(res_add).lower()

        # Read
        res_read = hyperlink(action="read", file_path=fp, sheet_name="Sheet1", cell_ref="D1")
        assert "google.com" in str(res_read)


# ---------------------------------------------------------------------------
# 3. Financial Routes
# ---------------------------------------------------------------------------

class TestFinancialRoute:
    def test_tvm_calc(self) -> None:
        res = time_value_calc(operation="fv", rate=0.05, nper=10, pmt=-100, pv=-1000)
        assert res["fv"] > 0


# ---------------------------------------------------------------------------
# 4. Governance Routes
# ---------------------------------------------------------------------------

class TestDocPropertiesRoute:
    def test_get_set_calc(self, sample_xlsx: str) -> None:
        fp = sample_xlsx
        # Set calc mode
        res_calc = doc_properties(action="set_calc_mode", file_path=fp, calc_mode="manual")
        assert "manual" in str(res_calc).lower()

        # Get core props
        res_props = doc_properties(action="get", file_path=fp)
        assert "creator" in str(res_props)


# ---------------------------------------------------------------------------
# 5. Worksheet Transfer Routes
# ---------------------------------------------------------------------------

class TestWorksheetTransferRoute:
    def test_copy_range(self, sample_xlsx: str) -> None:
        fp = sample_xlsx
        sheet_management(action="copy", file_path=fp, sheet_name="Sheet1", new_name="Sheet2")
        
        res = worksheet_transfer(
            action="copy_range_across",
            file_path=fp,
            source_sheet="Sheet1",
            target_sheet="Sheet2",
            source_range="A1:C2",
            target_start_cell="A1"
        )
        assert "copied" in str(res).lower()
