"""MCP server-side prompts for the excel-mcp server."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable


@dataclass(frozen=True, slots=True)
class PromptSpec:
    """Definition of a runtime prompt exposed through MCP."""

    name: str
    description: str
    builder: Callable[..., str]


def register_prompts(mcp: object) -> None:
    """Register all 20 Excel MCP prompts with the FastMCP instance."""
    for spec in PROMPT_CATALOGUE:
        mcp.prompt(name=spec.name, description=spec.description)(spec.builder)


# ---------------------------------------------------------------------------
# Prompt functions
# ---------------------------------------------------------------------------


def excel_quickstart(file_path: str, sheet_name: str = "Sheet1") -> str:
    return f"""\
You are working with an Excel MCP server. Create a new formatted Excel workbook
at `{file_path}` with a sheet named `{sheet_name}`.

Follow these steps in order:

1. Call `create_workbook` with file_path="{file_path}" and sheet_name="{sheet_name}".
2. Call `write_cells` (mode="range") with start_cell="A1" and a 2-D `data` array
   where the first row contains the column headers.
3. Call `format_cells` on the header row: bold=true, bg_color="366092",
   font_color="FFFFFF", horizontal_alignment="center".
4. Call `format_cells` on numeric or date data columns with appropriate number
   formats (e.g. "#,##0.00" for currency).
5. Call `worksheet_view` (action="freeze") with cell_ref="A2" to lock the headers.
6. Call `auto_fit_columns` to size every column to its content.

Constraints:
- Do NOT prefix hex colours with "#" — use "366092" not "#366092".
- Do NOT omit the mode parameter on write_cells.
- Pass the full 2-D data array in one write_cells call — not row by row.
- Default to ".xlsx" extension if the user has not specified one.

On completion, report the file path and the final row × column dimensions.
"""


def excel_data_analysis(
    file_path: str,
    sheet_name: str = "Sheet1",
    analysis_type: str = "all",
) -> str:
    return f"""\
Perform exploratory data analysis on sheet "{sheet_name}" in "{file_path}".
analysis_type = "{analysis_type}".

Steps:
1. Call `get_sheet_summary` first to confirm row/column counts and exact header names.
2. If analysis_type is "profile" or "all":
   - Call `profile_data` for dtype, null %, unique count, min/max/mean per column.
   - Call `column_statistics` for each numeric column to get IQR, p25, p75, p90.
3. If analysis_type is "filter" or "all":
   - Call `filter_data_advanced` with the user-specified conditions.
   - Set output_sheet when the filtered result exceeds 1,000 rows.
4. If analysis_type is "aggregate" or "all":
   - Call `aggregate_data` with group_by columns and required aggregation functions.
5. If analysis_type is "duplicates" or "all":
   - Call `find_duplicates` on the relevant subset columns.
6. Call `sort_data` to order results logically (e.g. descending by aggregate value).
7. Summarise all findings in a concise report with concrete numbers.

Constraints:
- NEVER use `filter_data` — always use `filter_data_advanced` (filter_data has a
  known numeric comparison bug).
- Pass percentile params as decimals: 0.25 not 25.
- Do not load raw data with read_cells before analysing — use analysis tools directly.
- If analysis_type is "all", run phases in order: profile → filter → aggregate → sort → duplicates.
"""


def excel_data_cleaning(
    input_path: str,
    output_path: str,
    sheet_name: str = "Sheet1",
    fill_strategy: str = "ffill",
) -> str:
    return f"""\
Execute a full data cleaning pipeline on sheet "{sheet_name}" in "{input_path}".
Write cleaned output to "{output_path}". Fill strategy: "{fill_strategy}".

Steps:
1. If input is .csv: call csv_ops (action="preview") to inspect raw structure,
   then csv_ops (action="to_xlsx") to convert BEFORE any edits.
2. Call get_sheet_summary to confirm shape and headers.
3. Call profile_data to record null rates, dtypes, and outliers BEFORE any
   changes — this is your baseline.
4. Call data_cleaner (preview=true) to review planned operations without modifying.
5. Call data_cleaner (preview=false): operations=["trim_whitespace","normalize_text",
   "fill_missing"], fill_missing_strategy="{fill_strategy}".
6. For delimited multi-value columns: call split_column with the appropriate delimiter.
7. For string-represented date columns: call parse_date_column to standardise to ISO.
8. Call deduplicate_data; or call find_duplicates first if inspection is needed.
9. If column constraints needed: call data_validation (action="numeric" or "dropdown").
10. Export: if CSV output required call csv_ops (action="to_csv"); otherwise the
    cleaned .xlsx is already written.

Constraints:
- NEVER call data_cleaner(preview=false) without first reviewing preview=true output.
- NEVER overwrite input_path — always write to output_path.
- When fill_strategy is "mean", only numeric columns are filled; string columns stay null.
- NEVER pass fill_missing_strategy="forward" or "backward" — use "ffill" or "bfill" respectively.
- Run profile_data before AND after cleaning for a before/after comparison.
- When input is .csv, convert to .xlsx first — all editing tools require .xlsx.
"""


def excel_chart_builder(
    file_path: str,
    data_range: str,
    sheet_name: str = "Sheet1",
    chart_type: str = "bar",
    title: str = "Chart",
) -> str:
    return f"""\
Create a fully annotated chart in "{file_path}" using data from "{data_range}"
on sheet "{sheet_name}". Chart type: "{chart_type}". Title: "{title}".
All chart operations use the `chart` consolidated tool.

Steps:
1. Call get_sheet_summary to confirm the data range is valid and headers are correct.
2. If chart_type is "combo":
   - Call chart (action="combo") specifying bar_columns and line_columns as 0-based lists.
3. Otherwise:
   - Call chart (action="create") with data_range="{data_range}", chart_type="{chart_type}",
     title="{title}", target_cell (anchor clear of data), x_axis_title, y_axis_title,
     width=15, height=10.
4. Call chart (action="set_axes") to configure axis number format (e.g. "#,##0") and
   min/max bounds if needed.
5. Call chart (action="data_labels") with show_value=true and appropriate label_position.
6. Call chart (action="legend") with show=true, legend_position="b".
7. For line or scatter charts: call chart (action="trendline") with
   trendline_type="linear", series_index=0.
8. For additional data series: call chart (action="add_series") per series.

Constraints:
- data_range MUST include the header row — headers become series labels.
- Place target_cell on an empty cell clear of all data.
- Do NOT apply add_series to a pie chart.
- Do NOT anchor chart on top of data cells.
- To delete a chart, use chart_index (0-based, from chart(list)) not title.
"""


def excel_report_builder(
    output_path: str,
    source_path: str,
    source_sheet: str = "Sheet1",
    report_title: str = "Report",
) -> str:
    return f"""\
Build a polished multi-sheet report workbook at "{output_path}".
Source data: "{source_path}" / "{source_sheet}". Title: "{report_title}".

Steps:
1. Call aggregate_data and column_statistics on "{source_path}" to compute
   summary stats BEFORE building the report workbook.
2. Call create_workbook at "{output_path}" with sheet_names=["Cover","Summary","Data","Charts"].
3. Cover sheet: call write_cells (mode="range") with title "{report_title}",
   generation date, source file path, and sheet index list.
   Apply large bold + centred formatting with format_cells.
4. Summary sheet: write aggregated stats with write_cells; apply conditional_format
   (action="highlight") on KPI cells to flag threshold crossings.
5. Data sheet: populate raw rows using write_cells (mode="range") or write_multi_sheet
   for bulk load — prefer write_multi_sheet when initialising multiple sheets.
6. Charts sheet: call chart (action="create") per visualisation; add axis labels
   with chart (action="set_axes") and legend with chart (action="legend").
7. Apply consistent header styling across every sheet: bold=true, uniform bg_color
   (NO "#" prefix), font_color="FFFFFF".
8. Call worksheet_view (action="freeze") on each data and summary sheet to freeze row 1.
9. Call worksheet_print (action="set_print_area") and worksheet_print (action="set_page_setup")
   on all non-Cover sheets for print-ready output.
10. Call auto_fit_columns on all sheets.

Constraints:
- Mandatory sheets: Cover, Summary, Data, Charts. Additional sheets allowed.
- NEVER use sequential single write_cells for bulk data — use write_multi_sheet.
- Define one colour scheme and reuse it across all sheets.
- Do NOT prefix hex colours with "#".
- Insert image (logo/branding) via insert_image if an image path is provided.
"""


def excel_financial_model(
    file_path: str,
    model_type: str = "full",
    principal: float = 100000.0,
    annual_rate: float = 0.05,
    term_years: int = 30,
) -> str:
    return f"""\
Build a multi-sheet financial model workbook at "{file_path}".
model_type="{model_type}", principal={principal}, annual_rate={annual_rate},
term_years={term_years}.

All rates are decimals — NEVER pass percentages.

Steps:
1. Call create_workbook at "{file_path}" with sheet names matching model_type
   (e.g. ["Amortisation","DCF","Ratios","BreakEven"] for "full").

2. AMORTISATION (model_type is "amortisation" or "full"):
   - Call loan_amortization with principal={principal}, annual_rate={annual_rate},
     years={term_years}, payments_per_year=12.
   - Write the returned schedule dict to the "Amortisation" sheet with write_cells.
   - Call chart (action="create") for a balance curve visualisation.

3. DCF (model_type is "dcf" or "full"):
   - Call dcf_analysis with cash_flows (year 0 = negative initial investment),
     discount_rate={annual_rate}, terminal_growth_rate, initial_investment.
   - Write NPV, IRR, payback period, terminal value to "DCF" sheet with write_cells.

4. RATIOS (model_type is "ratios" or "full"):
   - Call financial_ratio_analysis with a financial_data dict containing keys:
     current_assets, current_liabilities, total_debt, total_equity, net_income,
     total_assets, revenue, gross_profit, ebitda, interest_expense.
   - Write returned ratios to "Ratios" sheet with write_cells.

5. BREAK-EVEN (model_type is "full"):
   - Call break_even_analysis with fixed_costs, price_per_unit, variable_cost_per_unit.
   - Write results to "BreakEven" sheet.

6. For ad-hoc TVM: call time_value_calc with operation ("fv","pv","nper","rate",
   "depreciation") and the corresponding numeric parameters.

7. Apply header formatting on each sheet: bold=true, bg_color (no "#"), font_color="FFFFFF".
8. Apply number formats: "#,##0.00" for currency, "0.00%" for rates and ratios.
9. Call auto_fit_columns on all sheets.

Constraints:
- NEVER build amortisation rows manually — always call loan_amortization.
- All rates are decimals (0.05 for 5%), never percentages.
- DCF cash_flows array: index 0 = year 0 (negative initial investment).
- Build sheets in order when model_type="full": Amortisation → DCF → Ratios → BreakEven.
"""


def excel_pivot_etl(
    file_path: str,
    source_sheet: str = "Sheet1",
    output_sheet: str = "Output",
    transform_type: str = "all",
) -> str:
    return f"""\
Run ETL transforms on sheet "{source_sheet}" in "{file_path}".
transform_type="{transform_type}". Output to sheet "{output_sheet}".

Steps:
1. Call get_sheet_summary to inspect shape and exact column headers BEFORE any transform.

2. PIVOT (transform_type is "pivot" or "all"):
   - Call create_pivot_table with index_cols, value_cols, aggfunc,
     output_sheet="{output_sheet}", include_margins=true.
   - Call refresh_pivot_table if source data has changed since last build.

3. UNPIVOT (transform_type is "unpivot" or "all"):
   - Call unpivot_data with id_vars (anchor columns), value_vars (wide columns to
     collapse), var_name, value_name.

4. MERGE (transform_type is "merge" or "all"):
   - Call merge_datasets with file_path, sheet1, sheet2, left_on, right_on,
     how ("inner"/"left"/"right"/"outer"), output_sheet="{output_sheet}".
   - Call deduplicate_data on the output if a many-to-many join is possible.

5. COMPUTED_COL (transform_type is "computed_col" or "all"):
   - Call add_computed_column with new_column_name, expression (arithmetic over
     header names, e.g. "Revenue - Cost"), column_type="formula".
   - For running totals: column_type="cumsum", supply source_col.

6. Call sheet_management (action="rename") if the output sheet needs a better name.
7. Call worksheet_view (action="freeze") on the output sheet to freeze row 1.

Constraints:
- Do NOT write pivot results manually with write_cells — use create_pivot_table with output_sheet.
- Reference columns by EXACT header name, not position.
- Pivot state persists in hidden _mcp_pivots sheet across sessions.
- merge_datasets only merges within the same workbook — both sheets must be in "{file_path}".
- If transform_type="all", execute in order: pivot → unpivot → merge → computed_col.
"""


def excel_multi_file(
    file_paths: str,
    output_path: str,
    operation: str = "aggregate",
    sheet_name: str = "Sheet1",
) -> str:
    return f"""\
Consolidate or compare data from files: {file_paths}.
operation="{operation}", target sheet="{sheet_name}", output="{output_path}".

Steps:
1. Parse file_paths as a comma-separated list of file paths.
2. Call get_workbook_metadata on EACH source file to confirm sheet names, headers,
   and row counts before proceeding.
3. If any source file is .csv, first call csv_ops (action="to_xlsx") to convert it.
4. Call multi_file with action="{operation}" and the appropriate parameters:

   VALIDATE:
   - Call multi_file (action="validate") on all files to detect schema mismatches
     (missing/extra columns, type conflicts). Run this first for aggregate and filter ops.

   AGGREGATE:
   - Call multi_file (action="validate") first to confirm schema compatibility.
   - Call multi_file (action="aggregate") — the result includes a source-file column.

   FILTER:
   - Confirm schemas with validate first.
   - Call multi_file (action="filter") with a common filter condition valid across all files.

   COMPARE:
   - Use with exactly 2 files only.
   - Call multi_file (action="compare") to diff row by row — identify added, removed,
     and changed rows.

5. Write output to "{output_path}": call format_cells on the header row (bold, bg_color,
   no "#" prefix), then auto_fit_columns, then worksheet_view (action="freeze") row 1.

Constraints:
- NEVER use merge_datasets for cross-file operations — multi_file is purpose-built.
- compare accepts exactly 2 files; aggregate/filter/validate accept unlimited.
- ALWAYS run validate before aggregate to surface schema differences early.
- Source files are NEVER modified.
"""


def excel_statistical_analysis(
    file_path: str,
    target_col: str,
    sheet_name: str = "Sheet1",
    analysis_type: str = "both",
    feature_cols: str = "",
    forecast_steps: int = 12,
) -> str:
    return f"""\
Perform statistical modelling on sheet "{sheet_name}" in "{file_path}".
target_col="{target_col}", analysis_type="{analysis_type}",
feature_cols="{feature_cols}", forecast_steps={forecast_steps}.

Steps:
1. Call profile_data to verify column types and null rates. Confirm that
   "{target_col}" (and any feature columns) are numeric with no missing values.
   If nulls are found in the target column, STOP and ask user to run
   excel-data-cleaning first.
2. Call column_statistics on target and feature columns for distributions
   (mean, std, min, max, IQR, p25, p75, p90).

3. REGRESSION (analysis_type is "regression" or "both"):
   - Parse feature_cols "{feature_cols}" as a comma-separated list.
   - Call run_regression with y_column="{target_col}", x_columns=[list],
     output_sheet="Regression".
   - Results include: intercept, coefficients, standard errors, p-values,
     R², adjusted-R², F-statistic, and residuals.
   - Call chart (action="create") to plot actual vs. fitted values.
   - Call chart (action="trendline") with trendline_type="linear" for scatter plots.
   - Call write_cells to append a plain-language interpretation below the results.

4. SMOOTHING (analysis_type is "smoothing" or "both"):
   - Call run_exponential_smoothing with column="{target_col}",
     method="holt_winters" (or "holt" for trend-only; "simple" for flat series),
     forecast_steps={forecast_steps}.
   - For holt_winters, supply seasonal_periods (e.g. 12 for monthly data).
   - Call chart (action="create") to plot actual vs. smoothed + forecast horizon.
   - Call write_cells to append interpretation below results.

5. Call format_cells on header rows (bold, bg_color — no "#" prefix).
6. Call auto_fit_columns.

Constraints:
- x_columns MUST be a list even for single predictor: ["ColumnName"], not "ColumnName".
- Do NOT use holt_winters on series with <2 full seasonal cycles — fall back to "holt".
- Report p-values alongside R² — never interpret R² in isolation.
- If analysis_type="both", run regression FIRST, then smoothing.
- Verify output_sheet does not already exist before calling either modelling tool.
"""


def excel_formula_builder(
    file_path: str,
    sheet_name: str = "Sheet1",
    formula_type: str = "batch",
) -> str:
    return f"""\
Write, fill, or audit formulas on sheet "{sheet_name}" in "{file_path}".
formula_type="{formula_type}".

Steps:
1. Call get_sheet_summary to understand sheet layout BEFORE writing any formulas.

2. SET (formula_type is "set"):
   - Call formula_write (action="set") with cell_ref and formula (raw Excel syntax).
   - Example: formula="=SUM(A1:A10)"

3. BATCH (formula_type is "batch"):
   - Call formula_write (action="batch") with a formulas dict keyed by cell ref.
   - Example: {{"D2": "=SUM(B2:C2)", "D10": "=AVERAGE(D2:D9)"}}

4. FILL (formula_type is "fill"):
   - Confirm the origin cell_ref already contains a formula via read_cells.
   - Call formula_write (action="fill") with cell_ref and target_range.
   - The Translator adjusts relative refs per row/column; $-prefixed refs are unchanged.

5. AUTO_SUM (formula_type is "auto_sum"):
   - Call formula_write (action="auto_sum") with the data range.
   - SUM formulas are placed automatically at end of each column.

6. AUDIT (formula_type is "audit"):
   - Call formula_audit (action="errors") to list all error cells.
   - Call formula_audit (action="precedents") and/or (action="dependents") on
     specific cells as requested.
   - Call formula_audit (action="list") for a full formula inventory.

7. For frequently referenced ranges or inputs: call named_range (action="create")
   to improve formula readability. Check existing names first with named_range (action="list").

Constraints:
- Pass raw Excel formula syntax — no string escaping: "=SUM(A1:A10)" not "=SUM(A1\\:A10)".
- For batch, key the formulas dict by cell reference string ("D2"), not row index.
- formula_audit(value) reads the cached result — it does NOT force recalculation.
- Named range names must be unique across the entire workbook.
"""


def excel_data_governance(
    file_path: str,
    sheet_name: str = "Sheet1",
    governance_type: str = "all",
) -> str:
    return f"""\
Apply governance controls to sheet "{sheet_name}" in "{file_path}".
governance_type="{governance_type}".

Steps:
1. Call get_workbook_metadata to inspect existing named ranges, sheet list, and
   current protection status BEFORE making any changes.
2. Call get_sheet_summary to understand column layout.

3. VALIDATION (governance_type is "validation" or "all"):
   - Call data_validation (action="dropdown") for list-constrained cells:
     use options for static lists, source_range for dynamic lists.
   - Call data_validation (action="numeric") for numeric bounds: specify
     operator, value1, and optionally value2.
   - Call data_validation (action="date") for date range constraints.

4. PROTECTION — CRITICAL ORDER (governance_type is "protection" or "all"):
   a. Call protection (action="protect_cells") with unlocked_ranges listing
      all ranges users MUST be able to edit — DO THIS FIRST.
   b. THEN call protection (action="protect_sheet") with optional password.
   c. Call protection (action="protect_workbook") only if sheet structure
      (add/remove/rename) should be locked.

5. NAMED_RANGES (governance_type is "named_ranges" or "all"):
   - Call named_range (action="list") to see existing names.
   - Call named_range (action="create") for each key input cell or range;
     use descriptive business-purpose names.

6. SCENARIOS (governance_type is "scenarios" or "all"):
   - Call scenario (action="add") for base, optimistic, pessimistic variants
     with cell refs and values.
   - Call scenario (action="list") to verify.
   - Call scenario (action="apply") to switch the workbook to a target scenario.

7. Add explanatory comments to key cells via comment (action="add").
8. Call doc_properties (action="get") to verify workbook metadata.

CRITICAL CONSTRAINT — Protection order must be:
  unlock editable cells → protect_sheet (NEVER the other way round).
If protect_sheet is called before unlocking, call protection (action="unprotect_sheet")
immediately to recover, then retry in the correct order.

Additional constraints:
- Named range names must start with a letter or underscore and contain no spaces.
- Scenario cell refs must use Excel notation ("B5"), not header names.
- Scenarios persist in hidden _mcp_scenarios sheet across sessions.
"""


def excel_csv_workflow(
    csv_path: str,
    xlsx_path: str,
    analysis_goal: str = "clean",
    export_format: str = "xlsx",
) -> str:
    return f"""\
Full CSV-to-Excel workflow.
csv_path="{csv_path}" (read-only), output="{xlsx_path}",
analysis_goal="{analysis_goal}", export_format="{export_format}".

Steps:
1. Call csv_ops (action="preview") on "{csv_path}" to inspect the first 20 rows.
   Confirm column headers, delimiter, and encoding BEFORE conversion.
   If all data appears in a single column, the delimiter was not detected —
   do NOT proceed until confirmed.

2. Call csv_ops (action="to_xlsx") to convert "{csv_path}" to "{xlsx_path}".
   After this step, work EXCLUSIVELY on "{xlsx_path}" — do not read from "{csv_path}" again.

3. Call get_sheet_summary on "{xlsx_path}" to verify row and column counts.

4. Call profile_data to assess data quality: null rates, types, unique counts,
   numeric distributions.

5. Call data_cleaner (preview=true) to plan cleaning transforms — review before applying.

6. Call data_cleaner (preview=false):
   operations=["trim_whitespace","normalize_text","fill_missing"],
   fill_missing_strategy chosen based on profile output (use "mean" for numeric
   columns only; use "ffill" or "bfill" for string columns).

7. Apply analysis per analysis_goal="{analysis_goal}":
   - clean: steps 4–6 constitute the analysis.
   - filter: call filter_data_advanced with conditions; set output_sheet for large results.
   - aggregate: call aggregate_data with group_by and aggregation functions.
   - sort: call sort_data with the appropriate columns and direction.

8. Apply formatting: call format_cells for bold headers and any number formatting,
   then auto_fit_columns.

9. If export_format is "csv":
   - Call csv_ops (action="to_csv") to export the final sheet back to CSV.
   - Note: to_csv drops all formatting; only data values are preserved.

10. Report: row counts before and after each transform, nulls filled, columns
    trimmed, and the confirmed export path.

Constraints:
- NEVER edit "{csv_path}" directly with write_cells — convert to .xlsx first.
- NEVER skip csv_ops(preview) — delimiter and encoding issues must be caught early.
- Cleaning order is mandatory: profile → clean → analyse.
- ALWAYS call data_cleaner(preview=true) before data_cleaner(preview=false).
- NEVER pass fill_missing_strategy="forward" or "backward" — use "ffill" or "bfill" respectively.
- Use filter_data_advanced — NOT filter_data (deprecated, has a numeric comparison bug).
"""


def _as_text(value: object | None, fallback: str = "<not supplied>") -> str:
    """Render optional prompt inputs without leaking ``None`` into the instructions."""
    if value is None:
        return fallback
    if isinstance(value, (list, tuple, set)):
        text = ", ".join(str(item) for item in value)
        return text or fallback
    text = str(value)
    return text if text else fallback


def excel_readonly_audit(
    file_path: str,
    sheet_name: str = "Sheet1",
    scope: str = "workbook",
    focus: str = "all",
    target_range: str | None = None,
) -> str:
    target_range_text = _as_text(target_range)
    return f"""\
Inspect the Excel workbook at `{file_path}`. This prompt is read-only and must not mutate the file.

Inputs:
- sheet_name = `{sheet_name}`
- scope = `{scope}` (`workbook`, `sheet`, or `range`)
- focus = `{focus}` (`metadata`, `data_quality`, `formulas`, or `all`)
- target_range = `{target_range_text}`

Steps:
1. Call `get_workbook_metadata` first. If workbook-level settings matter, also call `doc_properties` (action="get").
2. If scope is `sheet` or `range`, call `get_sheet_summary` for `{sheet_name}`.
3. If scope is `range` and `target_range` is supplied, use `read_cells`
   (mode="range") on that range; otherwise keep raw reads minimal.
4. If focus includes `data_quality`, call `profile_data` on the relevant
   sheet/range and summarise nulls, dtypes, and suspicious values.
5. If focus includes `formulas`, call `formula_audit` (action="list"), then
   inspect `errors`, `precedents`, and `dependents` for suspicious cells.
6. If tables are present, call `table` (action="list"). If named ranges matter, call `named_range` (action="list").
7. Report workbook structure, risks, and likely follow-up actions, but do not edit anything.

Constraints:
- Never call `write_cells`, `formula_write`, `sheet_management`,
  `worksheet_view`, `worksheet_structure`, `worksheet_print`, `worksheet_transfer`, `protection`, or any other mutating tool.
- Prefer metadata and summaries before raw cell reads.
- If `scope` is `range` without a `target_range`, ask for the range rather than guessing.
"""


def excel_formula_diagnosis(
    file_path: str,
    sheet_name: str = "Sheet1",
    target_cells: str | list[str] | None = None,
    formula_range: str | None = None,
    mode: str = "diagnose",
) -> str:
    target_cells_text = _as_text(target_cells)
    formula_range_text = _as_text(formula_range)
    return f"""\
Diagnose formulas in `{file_path}` on sheet `{sheet_name}`.
mode = `{mode}`
target_cells = `{target_cells_text}`
formula_range = `{formula_range_text}`

Steps:
1. Call `get_sheet_summary` and `formula_audit` (action="list") to map the formula surface.
2. If target cells are provided, inspect each with `formula_audit` (action="value"),
   `formula_audit` (action="precedents"), and `formula_audit` (action="dependents").
   Use `read_cells` only when you need the displayed formula or cached value.
3. If `formula_range` is provided, call `formula_audit` (action="errors")
   limited to that range and use `read_cells` (mode="range", show_formula=true)
   to compare neighbouring formulas.
4. For `mode="audit"` or `mode="diagnose"`, explain the root cause, impacted
   cells, and smallest safe fix without changing the workbook.
5. For `mode="repair"`, call `formula_write` only after the defect is
   localised; prefer the narrowest possible edit, then re-run `formula_audit`
   to verify.
6. If repeated inputs or lookup ranges are part of the issue, consider
   `named_range` (action="create") to make the workbook easier to reason about.

Constraints:
- Do not rewrite formulas until the defect is localised.
- Do not change unrelated cells.
- Preserve relative and absolute references exactly; only translate where intended.
- If multiple failures have different causes, diagnose them separately.
"""


def excel_workbook_maintenance(
    file_path: str,
    sheet_name: str | None = None,
    action: str = "all",
) -> str:
    sheet_name_text = _as_text(sheet_name)
    return f"""\
Perform workbook maintenance on `{file_path}`.
sheet_name = `{sheet_name_text}`
action = `{action}` (`sheets`, `layout`, `print`, `rows_cols`, `freeze`, or `all`)

Steps:
1. Call `get_workbook_metadata` first, then `get_sheet_summary` for the target sheet(s).
2. If action includes `sheets`, use `sheet_management` for rename/copy/hide/
   unhide only after confirming the structural change.
3. If action includes `layout` or `rows_cols`, use `worksheet_structure` for
   insert/delete rows or columns, row heights, column widths, grouping;
   use `worksheet_view` for gridlines and freeze panes as needed.
4. If action includes `print`, use `worksheet_print` for `set_print_area`, `set_page_setup`, and `set_print_titles`.
5. If visible sheets need presentation polish, call `format_cells` on headers
   and `auto_fit_columns` on the relevant sheets.
6. Summarise every change and call out any destructive operation explicitly.

Constraints:
- Never delete, hide, or rename sheets without an explicit request or a clearly requested maintenance action.
- Confirm the target sheet list before applying structural changes.
- Keep changes conservative; prefer layout fixes over data edits.
"""


def excel_table_manager(
    file_path: str,
    sheet_name: str = "Sheet1",
    table_name: str | None = None,
    table_range: str | None = None,
    action: str = "list",
    new_range: str | None = None,
    show_totals: bool | None = None,
    column_totals: dict[str, str] | None = None,
    style_name: str = "TableStyleMedium9",
) -> str:
    table_name_text = _as_text(table_name)
    table_range_text = _as_text(table_range)
    new_range_text = _as_text(new_range)
    show_totals_text = _as_text(show_totals)
    column_totals_text = _as_text(column_totals)
    return f"""\
Manage native Excel tables in `{file_path}` on sheet `{sheet_name}`.
action = `{action}` (`create`, `list`, `data`, `resize`, `totals`, `convert`)
table_name = `{table_name_text}`
table_range = `{table_range_text}`
new_range = `{new_range_text}`
show_totals = `{show_totals_text}`
column_totals = `{column_totals_text}`
style_name = `{style_name}`

Steps:
1. Call `get_sheet_summary` and `table` (action="list") first to confirm existing tables and exact headers.
2. For `action="create"`, use `table` (action="create") with `table_name`
   and `table_range`, then optionally apply totals or header formatting.
3. For `action="list"` or `action="data"`, inspect table metadata and
   contents before any structural edits.
4. For `action="resize"`, confirm the new range still includes the header row
   and the full data block.
5. For `action="totals"`, enable or disable the totals row and set per-column
   functions only after confirming the schema.
6. For `action="convert"`, call `table` (action="convert_to_range") while preserving data and formatting.
7. Use `format_cells` and `auto_fit_columns` only after the table structure is final.

Constraints:
- Never create a table on top of blank or partial headers.
- Keep structured references intact; resize only when the new range is contiguous and valid.
- Do not convert or resize a table unless the user explicitly requests it.
"""


def excel_what_if_analysis(
    file_path: str,
    sheet_name: str = "Sheet1",
    analysis_type: str = "goal_seek",
    objective_expression: str | None = None,
    target_value: float | None = None,
    target_cell: str | None = None,
    variable_cells: str | list[str] | None = None,
    constraints: str | list[str] | None = None,
    scenario_name: str | None = None,
) -> str:
    objective_expression_text = _as_text(objective_expression)
    target_value_text = _as_text(target_value)
    target_cell_text = _as_text(target_cell)
    variable_cells_text = _as_text(variable_cells)
    constraints_text = _as_text(constraints)
    scenario_name_text = _as_text(scenario_name)
    return f"""\
Run what-if analysis on `{file_path}` using sheet `{sheet_name}`.
analysis_type = `{analysis_type}` (`goal_seek`, `solver`, `sensitivity`, `scenario`, or `all`)
objective_expression = `{objective_expression_text}`
target_value = `{target_value_text}`
target_cell = `{target_cell_text}`
variable_cells = `{variable_cells_text}`
constraints = `{constraints_text}`
scenario_name = `{scenario_name_text}`

Steps:
1. Call `get_sheet_summary` and `profile_data` first to confirm the inputs
   are numeric and the sheet structure is stable.
2. If analysis_type is `goal_seek` or `all`, use `goal_seek` for a single variable and a single target.
3. If analysis_type is `solver` or `all`, use `run_solver` for constrained
   multi-variable optimisation; keep bounds explicit.
4. If analysis_type is `sensitivity` or `all`, use `create_sensitivity_table`
   to sweep the selected variables and write the result grid.
5. If analysis_type is `scenario` or `all`, use `scenario` to add, list, or apply named assumption sets.
6. Explain the recommendation, the trade-offs, and which assumptions were most influential.

Constraints:
- Do not use `run_solver` when a single-variable `goal_seek` is sufficient.
- Keep constraints explicit; if they are missing, ask for them rather than guessing bounds.
- Write results back only when the analysis type requires it and the target cells are clear.
"""


def excel_multi_file_reconciliation(
    file_paths: list[str] | str,
    output_path: str,
    operation: str = "validate",
    sheet_name: str = "Sheet1",
    key_columns: str | list[str] | None = None,
) -> str:
    file_paths_text = _as_text(file_paths)
    key_columns_text = _as_text(key_columns)
    return f"""\
Reconcile data across multiple Excel files.
file_paths = `{file_paths_text}`
output_path = `{output_path}`
operation = `{operation}` (`validate`, `aggregate`, `filter`, or `compare`)
sheet_name = `{sheet_name}`
key_columns = `{key_columns_text}`

Steps:
1. Parse `file_paths` into a source list and call `get_workbook_metadata` for
   each file to confirm sheets, headers, and row counts.
2. If any source is CSV, convert it first with `csv_ops` (action="to_xlsx") and continue on the XLSX copy.
3. If operation is `validate`, call `multi_file` (action="validate") first and surface schema mismatches early.
4. If operation is `aggregate` or `filter`, validate first, then run the requested action on the common schema.
5. If operation is `compare`, require exactly two files and call `multi_file` (action="compare") with those files.
6. Write the result to `output_path`, then format the header row and auto-fit columns if a workbook was written.

Constraints:
- Never mutate the source files.
- Always validate before aggregate or filter.
- Only compare exactly two files.
- If multiple key columns are supplied but the tool needs a single key, choose
   the most stable unique key or ask for clarification.
"""


def excel_search_repair(
    file_path: str,
    sheet_name: str = "Sheet1",
    search_pattern: str = "",
    replacement: str = "",
    scope: str = "sheet",
    match_case: bool = False,
    match_entire_cell: bool = False,
    search_formulas: bool = False,
) -> str:
    search_pattern_text = _as_text(search_pattern)
    replacement_text = _as_text(replacement)
    return f"""\
Find and repair content in `{file_path}` on sheet `{sheet_name}`.
search_pattern = `{search_pattern_text}`
replacement = `{replacement_text}`
scope = `{scope}` (`sheet`, `workbook`, or `formulas`)
match_case = `{match_case}`
match_entire_cell = `{match_entire_cell}`
search_formulas = `{search_formulas}`

Steps:
1. Call `get_sheet_summary` and inspect representative matches with `read_cells` before changing anything.
2. Use `find_replace` for plain-text or formula-string repairs; set
   `search_formulas=true` when formula text itself should be searched.
3. If the repair touches formulas, re-run `formula_audit` on the affected
   cells or range to confirm the fix and check for collateral errors.
4. If the pattern is broad, test the search on a small scope first, then expand only after the preview looks correct.
5. Summarise the matches, replacements, and any cells that need a manual follow-up.

Constraints:
- Never bulk-replace without seeing at least one sample match.
- Keep the scope as narrow as possible.
- Treat regex-like repairs as out of scope for this prompt; use the safe
   transform prompt instead if the search must be pattern-based.
"""


def excel_safe_transform(
    file_path: str,
    output_path: str,
    sheet_name: str = "Sheet1",
    transform_goal: str = "",
    preview_only: bool = True,
) -> str:
    transform_goal_text = _as_text(transform_goal)
    return f"""\
Apply a safe custom transform to `{file_path}` and write to `{output_path}`.
sheet_name = `{sheet_name}`
transform_goal = `{transform_goal_text}`
preview_only = `{preview_only}`

Steps:
1. Inspect the input with `get_workbook_metadata` and `get_sheet_summary`; if
   the source is CSV, preview or convert it before editing.
2. Prefer built-in tools first; only use `execute_custom_code` when the
   requested transform cannot be done safely with existing MCP tools.
3. Keep any custom code minimal, sandboxed, and limited to pandas/numpy
   operations over the selected sheet or DataFrame.
4. If `preview_only` is true, describe the intended transform and validate the approach without writing the output file.
5. If writing, send the result to `output_path` and summarise the before/after shape and any column changes.

Constraints:
- Never overwrite the source file.
- Never use custom code when a built-in tool can do the job.
- Keep the transformation deterministic and easy to audit.
"""


PROMPT_CATALOGUE: tuple[PromptSpec, ...] = (
    PromptSpec(
        name="excel-quickstart",
        description="Create a new Excel workbook, populate it with formatted data, and auto-fit columns",
        builder=excel_quickstart,
    ),
    PromptSpec(
        name="excel-data-analysis",
        description="Profile, filter, aggregate, sort, and find duplicates in an Excel dataset",
        builder=excel_data_analysis,
    ),
    PromptSpec(
        name="excel-data-cleaning",
        description="Full data cleaning pipeline: profile, clean, validate, deduplicate, and export",
        builder=excel_data_cleaning,
    ),
    PromptSpec(
        name="excel-chart-builder",
        description="Create, style, and annotate charts with trendlines, axis labels, data labels, and legends",
        builder=excel_chart_builder,
    ),
    PromptSpec(
        name="excel-report-builder",
        description="Assemble a multi-sheet formatted report with data, charts, statistics, and print setup",
        builder=excel_report_builder,
    ),
    PromptSpec(
        name="excel-financial-model",
        description="Build a financial model workbook with loan amortisation, DCF analysis, and financial ratios",
        builder=excel_financial_model,
    ),
    PromptSpec(
        name="excel-pivot-etl",
        description="Create pivot tables and run ETL transforms: merge, unpivot, add computed columns",
        builder=excel_pivot_etl,
    ),
    PromptSpec(
        name="excel-multi-file",
        description="Aggregate, filter, compare, and validate schema consistency across multiple Excel files",
        builder=excel_multi_file,
    ),
    PromptSpec(
        name="excel-statistical-analysis",
        description="Run OLS regression and exponential smoothing with forecasting on Excel data",
        builder=excel_statistical_analysis,
    ),
    PromptSpec(
        name="excel-formula-builder",
        description="Write, fill, auto-sum, and audit Excel formulas across a sheet",
        builder=excel_formula_builder,
    ),
    PromptSpec(
        name="excel-data-governance",
        description="Apply data validation, protection, named ranges, and scenario management to a workbook",
        builder=excel_data_governance,
    ),
    PromptSpec(
        name="excel-csv-workflow",
        description="Preview a CSV, convert to Excel, clean, analyse, and export results",
        builder=excel_csv_workflow,
    ),
    PromptSpec(
        name="excel-readonly-audit",
        description="Inspect workbook structure, formulas, tables, and data quality without mutating the file",
        builder=excel_readonly_audit,
    ),
    PromptSpec(
        name="excel-formula-diagnosis",
        description="Diagnose formula errors, trace precedents and dependents, and apply the smallest safe fix",
        builder=excel_formula_diagnosis,
    ),
    PromptSpec(
        name="excel-workbook-maintenance",
        description="Perform conservative workbook housekeeping: sheets, layout, print setup, and sizing",
        builder=excel_workbook_maintenance,
    ),
    PromptSpec(
        name="excel-table-manager",
        description="Create, inspect, resize, total, and convert native Excel tables",
        builder=excel_table_manager,
    ),
    PromptSpec(
        name="excel-what-if-analysis",
        description="Run goal seek, solver, sensitivity, and scenario analysis",
        builder=excel_what_if_analysis,
    ),
    PromptSpec(
        name="excel-multi-file-reconciliation",
        description="Validate, compare, aggregate, and filter multiple files with schema checks",
        builder=excel_multi_file_reconciliation,
    ),
    PromptSpec(
        name="excel-search-repair",
        description="Find and repair text or formula content with a read-before-write workflow",
        builder=excel_search_repair,
    ),
    PromptSpec(
        name="excel-safe-transform",
        description="Apply a sandboxed custom transform only when built-in tools are not enough",
        builder=excel_safe_transform,
    ),
)
