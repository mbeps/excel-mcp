"""Tests for MCP server-side prompts."""

from __future__ import annotations

import mcp_server.prompts as prompts_module
from mcp_server.prompts import (
    excel_chart_builder,
    excel_csv_workflow,
    excel_data_analysis,
    excel_data_cleaning,
    excel_data_governance,
    excel_financial_model,
    excel_formula_builder,
    excel_formula_diagnosis,
    excel_multi_file,
    excel_multi_file_reconciliation,
    excel_pivot_etl,
    excel_quickstart,
    excel_readonly_audit,
    excel_report_builder,
    excel_safe_transform,
    excel_search_repair,
    excel_statistical_analysis,
    excel_table_manager,
    excel_what_if_analysis,
    excel_workbook_maintenance,
)

EXPECTED_PROMPT_NAMES = [
    "excel-quickstart",
    "excel-data-analysis",
    "excel-data-cleaning",
    "excel-chart-builder",
    "excel-report-builder",
    "excel-financial-model",
    "excel-pivot-etl",
    "excel-multi-file",
    "excel-statistical-analysis",
    "excel-formula-builder",
    "excel-data-governance",
    "excel-csv-workflow",
    "excel-readonly-audit",
    "excel-formula-diagnosis",
    "excel-workbook-maintenance",
    "excel-table-manager",
    "excel-what-if-analysis",
    "excel-multi-file-reconciliation",
    "excel-search-repair",
    "excel-safe-transform",
]


class TestPromptReturnValues:
    """Prompt functions must return non-empty strings containing key context."""

    def test_quickstart_returns_str_with_filepath(self):
        result = excel_quickstart("test.xlsx")
        assert isinstance(result, str) and len(result) > 50
        assert "test.xlsx" in result

    def test_quickstart_sheet_name_in_output(self):
        result = excel_quickstart("test.xlsx", sheet_name="Sales")
        assert "Sales" in result

    def test_data_analysis_defaults(self):
        result = excel_data_analysis("data.xlsx")
        assert isinstance(result, str) and "data.xlsx" in result

    def test_data_analysis_sheet_and_type(self):
        result = excel_data_analysis("data.xlsx", sheet_name="Txn", analysis_type="profile")
        assert "Txn" in result and "profile" in result.lower()

    def test_data_cleaning_both_paths(self):
        result = excel_data_cleaning("in.xlsx", "out.xlsx")
        assert "in.xlsx" in result and "out.xlsx" in result
        assert "profile_data (preview=true)" not in result
        assert "profile_data to record null rates" in result

    def test_data_cleaning_fill_strategy(self):
        result = excel_data_cleaning("in.xlsx", "out.xlsx", fill_strategy="mean")
        assert "mean" in result

    def test_chart_builder_required_params(self):
        result = excel_chart_builder("chart.xlsx", "A1:D13")
        assert "A1:D13" in result

    def test_chart_builder_title_in_output(self):
        result = excel_chart_builder("chart.xlsx", "A1:D13", title="Revenue Trend")
        assert "Revenue Trend" in result

    def test_report_builder_paths(self):
        result = excel_report_builder("out.xlsx", "src.xlsx")
        assert "out.xlsx" in result and "src.xlsx" in result

    def test_report_builder_title(self):
        result = excel_report_builder("out.xlsx", "src.xlsx", report_title="Q1 Report")
        assert "Q1 Report" in result

    def test_financial_model_defaults(self):
        result = excel_financial_model("model.xlsx")
        assert isinstance(result, str) and "model.xlsx" in result

    def test_financial_model_params_in_output(self):
        result = excel_financial_model("model.xlsx", principal=500000.0, annual_rate=0.04, term_years=15)
        assert "500000" in result and "0.04" in result

    def test_pivot_etl_defaults(self):
        result = excel_pivot_etl("data.xlsx")
        assert isinstance(result, str) and "data.xlsx" in result

    def test_pivot_etl_transform_type(self):
        result = excel_pivot_etl("data.xlsx", transform_type="merge")
        assert "merge" in result.lower()

    def test_multi_file_required(self):
        result = excel_multi_file("a.xlsx,b.xlsx", "out.xlsx")
        assert "a.xlsx" in result and "out.xlsx" in result

    def test_multi_file_operation(self):
        result = excel_multi_file("a.xlsx,b.xlsx", "out.xlsx", operation="compare")
        assert "compare" in result.lower()

    def test_statistical_analysis_target_col(self):
        result = excel_statistical_analysis("data.xlsx", "Revenue")
        assert "Revenue" in result

    def test_statistical_analysis_forecast_steps(self):
        result = excel_statistical_analysis("data.xlsx", "Sales", forecast_steps=24)
        assert "24" in result

    def test_formula_builder_defaults(self):
        result = excel_formula_builder("wb.xlsx")
        assert isinstance(result, str) and "wb.xlsx" in result

    def test_formula_builder_type(self):
        result = excel_formula_builder("wb.xlsx", formula_type="audit")
        assert "audit" in result.lower()

    def test_data_governance_defaults(self):
        result = excel_data_governance("wb.xlsx")
        assert isinstance(result, str) and "wb.xlsx" in result

    def test_data_governance_type(self):
        result = excel_data_governance("wb.xlsx", governance_type="protection")
        assert "protection" in result.lower() or "protect" in result.lower()

    def test_csv_workflow_required(self):
        result = excel_csv_workflow("data.csv", "data.xlsx")
        assert "data.csv" in result and "data.xlsx" in result

    def test_csv_workflow_export_format(self):
        result = excel_csv_workflow("data.csv", "data.xlsx", export_format="csv")
        assert "csv" in result.lower()


class TestRegisterPrompts:
    """register_prompts must work with a _NoopMCP-like object."""

    class _RecordingMCP:
        def __init__(self):
            self.prompts: list[tuple[str | None, str | None, object]] = []

        def prompt(self, *args, **kwargs):
            name = kwargs.get("name")
            description = kwargs.get("description")

            def decorator(fn):
                self.prompts.append((name, description, fn))
                return fn

            return decorator

    def test_register_with_noop_mcp(self):
        """register_prompts should not raise with a noop-style MCP object."""

        class _Noop:
            def prompt(self, *args, **kwargs):
                def decorator(fn):
                    return fn

                return decorator

        prompts_module.register_prompts(_Noop())  # must not raise

    def test_register_catalogue_matches_expected_names(self):
        recorder = self._RecordingMCP()

        prompts_module.register_prompts(recorder)

        assert [name for name, _, _ in recorder.prompts] == EXPECTED_PROMPT_NAMES
        assert len(recorder.prompts) == len(EXPECTED_PROMPT_NAMES)
        assert all(description for _, description, _ in recorder.prompts)


def test_prompt_catalogue_names_and_count():
    assert len(prompts_module.PROMPT_CATALOGUE) == len(EXPECTED_PROMPT_NAMES)
    assert [spec.name for spec in prompts_module.PROMPT_CATALOGUE] == EXPECTED_PROMPT_NAMES


def test_prompt_catalogue_invariants():
    assert len({spec.name for spec in prompts_module.PROMPT_CATALOGUE}) == len(prompts_module.PROMPT_CATALOGUE)
    assert all(spec.name.startswith("excel-") for spec in prompts_module.PROMPT_CATALOGUE)
    assert all(spec.description.strip() for spec in prompts_module.PROMPT_CATALOGUE)


def test_readonly_audit_prompt_mentions_read_only_and_key_tools():
    result = excel_readonly_audit(
        "audit.xlsx",
        sheet_name="Overview",
        scope="range",
        focus="formulas",
        target_range="A1:D20",
    )

    assert isinstance(result, str) and len(result) > 50
    assert "audit.xlsx" in result and "Overview" in result and "A1:D20" in result
    assert "read-only" in result.lower()
    assert "profile_data" in result and "formula_audit" in result and "table" in result


def test_formula_diagnosis_prompt_mentions_audit_and_repair_tools():
    result = excel_formula_diagnosis(
        "wb.xlsx",
        sheet_name="Calc",
        target_cells=["D2", "D3"],
        formula_range="D2:D20",
        mode="repair",
    )

    assert isinstance(result, str) and len(result) > 50
    assert "wb.xlsx" in result and "Calc" in result and "D2:D20" in result
    assert "formula_audit" in result and "formula_write" in result and "named_range" in result


def test_workbook_maintenance_prompt_mentions_wiring_tools():
    result = excel_workbook_maintenance("wb.xlsx", sheet_name="Ops", action="print")

    assert isinstance(result, str) and len(result) > 50
    assert "wb.xlsx" in result and "Ops" in result and "print" in result
    assert "sheet_management" in result and "worksheet_ops" in result and "auto_fit_columns" in result


def test_table_manager_prompt_mentions_table_actions():
    result = excel_table_manager(
        "wb.xlsx",
        sheet_name="Data",
        table_name="SalesTable",
        table_range="A1:D20",
        action="resize",
        new_range="A1:E25",
        show_totals=True,
    )

    assert isinstance(result, str) and len(result) > 50
    assert "SalesTable" in result and "A1:D20" in result and "A1:E25" in result
    assert "table" in result and "convert_to_range" in result and "totals" in result


def test_what_if_analysis_prompt_mentions_solver_tools():
    result = excel_what_if_analysis(
        "model.xlsx",
        sheet_name="Assumptions",
        analysis_type="all",
        objective_expression="B2 * B3 - B4",
        target_value=250000,
        target_cell="B10",
        variable_cells=["B2", "B3"],
        constraints=["B2 + B3 <= 100", "B2 >= 0"],
        scenario_name="Base",
    )

    assert isinstance(result, str) and len(result) > 50
    assert "model.xlsx" in result and "Assumptions" in result and "Base" in result
    assert "goal_seek" in result and "run_solver" in result
    assert "create_sensitivity_table" in result and "scenario" in result


def test_multi_file_reconciliation_prompt_mentions_validation_and_compare():
    result = excel_multi_file_reconciliation(
        ["a.xlsx", "b.xlsx"],
        "out.xlsx",
        operation="compare",
        sheet_name="Sheet1",
        key_columns=["ID"],
    )

    assert isinstance(result, str) and len(result) > 50
    assert "a.xlsx" in result and "b.xlsx" in result and "out.xlsx" in result
    assert "multi_file" in result and "validate" in result and "compare" in result


def test_search_repair_prompt_mentions_find_replace_and_formula_audit():
    result = excel_search_repair(
        "wb.xlsx",
        sheet_name="Data",
        search_pattern="Total",
        replacement="Grand Total",
        scope="workbook",
        search_formulas=True,
    )

    assert isinstance(result, str) and len(result) > 50
    assert "wb.xlsx" in result and "Data" in result and "Total" in result and "Grand Total" in result
    assert "find_replace" in result and "search_formulas" in result and "formula_audit" in result


def test_safe_transform_prompt_mentions_execute_custom_code_and_preview_mode():
    result = excel_safe_transform(
        "input.xlsx",
        "output.xlsx",
        sheet_name="Raw",
        transform_goal="normalise and aggregate",
        preview_only=True,
    )

    assert isinstance(result, str) and len(result) > 50
    assert "input.xlsx" in result and "output.xlsx" in result and "Raw" in result
    assert "execute_custom_code" in result and "sandboxed" in result and "preview_only" in result
