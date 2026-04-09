# Excel MCP Server

A Python-based Model Context Protocol (MCP) server for Excel automation and data manipulation. Enables LLMs to interact with Excel files (`.xlsx`, `.xls`, `.csv`, `.xlsm`) through 66 structured tools, supporting workflows from basic spreadsheet operations to advanced financial modelling and data analysis.

# Features

> NOTE: Custom actions are not supported as first-class tools. Use `execute_custom_code` as the sanctioned, sandboxed path for custom data transformations. This uses AST validation and blocks dangerous operations.

## Workbook & Sheet Management
- Create workbooks and list, rename, delete, copy, and move sheets.
- Retrieve workbook metadata and per-sheet summaries.
- Write data to multiple sheets in a single call.

## Cell & Range Operations
- Read and write individual cells and contiguous ranges.
- Read detailed cell metadata (type, style, formula, value).
- Copy and delete ranges.
- Read large files in chunks using the python-calamine engine.
- Transpose data ranges (rows ↔ columns).
- Search and replace values with regex support.
- Retrieve workbook-level information (sheet count, metadata, etc.).

## Row & Column Operations
- Insert and delete rows and columns by index or letter.

## Formatting & Styling
- Apply fonts, colours, alignment, borders, and number formats to cells.
- Apply named Excel styles (Normal, Good, Bad, Heading 1, etc.) to a range.
- Use built-in number format presets (currency, percentage, date, accounting, etc.) via `format_cells`.
- Clear all formatting from a range without affecting cell values.
- Set column widths and row heights; auto-fit columns to content.
- Merge and unmerge cells; list all merged ranges.
- Apply per-cell format arrays across a range.
- Set gradient fills.
- Inspect existing cell formatting.
- Copy formatting from one range to another.

## Conditional Formatting
- Apply colour scales, data bars, and icon sets.
- Add highlight, formula, top/bottom, above/below-average, and duplicate rules.
- List and remove conditional formatting rules.

## Formulas
- Set single, array, and batch formulas.
- Drag-fill a formula across a range (relative-reference translation via openpyxl Translator).
- Insert AutoSum formulas for one or more ranges in a single call.
- Validate formula syntax.
- List all formulas in a sheet.
- Convert formulas to static values.

## Tables
- Create, list, rename, resize, and delete native Excel tables.
- Toggle the totals row and read table data.
- Convert a native Excel table back to a plain range.

## Data Validation
- Add dropdown, numeric, date, text-length, and formula-based validation rules.
- List and remove validation rules.

## Protection
- Protect and unprotect sheets.
- Lock individual cells.

## Charts
- Create 10 chart types: column, bar, line, pie, scatter, area, radar, doughnut, bubble, and stock.
- Configure axes, trendlines, and combo (dual-axis) charts.
- Add and configure data labels on chart series.
- Show, hide, and position chart legends.
- Add and remove chart series.
- List and delete charts.

## Data Analysis
- Filter data by single or multiple conditions (==, !=, >, <, contains, startswith, etc.).
- Sort by one or more columns.
- Compute column statistics (mean, median, min, max, std, sum).
- Aggregate and group data.
- Find and remove duplicate rows.
- Profile a dataset comprehensively, including per-column percentile stats (p25, p75, p90, IQR).
- Insert subtotal formula rows after each group in a dataset (SUM, AVG, COUNT, MAX, MIN), with optional grand total.
- Count value frequencies for a column, with normalize, top-n, and dropna options.
- Calculate a Pearson correlation matrix across numeric columns.
- VLOOKUP-style lookup helper across ranges.
- Export analysis results to a new file.

## CSV Operations
- Preview CSV content.
- Convert between CSV and XLSX formats.

## Pivot & ETL
- Create pivot tables.
- Refresh an existing pivot table from updated source data.
- Unpivot (melt) data from wide to long format.
- Merge datasets using SQL-style joins.
- Add computed columns with safe expression evaluation, or compute a running cumulative sum column.
- Deduplicate rows.
- Append datasets from multiple sources.
- Find differences between two datasets.

## Financial Calculations
- NPV, IRR, FV, PV, NPER, and RATE via time-value-of-money operations.
- DCF (discounted cash flow) analysis.
- Loan amortisation schedules.
- Goal seek with AST-validated expressions.
- Budget variance analysis.
- Financial ratio analysis.
- Scenario analysis.
- Break-even analysis.
- Sensitivity tables.

## Data Cleaning
- Configurable cleaning pipeline: trim whitespace, remove empty rows/columns, fix number formats, deduplicate, and fill missing values.
- Preview mode (dry run without saving).
- Split a column into multiple columns.
- Parse and normalise date columns into a standard format.

## Comments
- Add, read, update, and delete comments.
- List all comments in a sheet.

## Hyperlinks
- Add external and internal (intra-workbook) hyperlinks.
- Read, delete, and list hyperlinks.

## Images
- Insert images into worksheets (`.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`).

## Named Ranges
- Create, list, rename, update, and delete named ranges with scope preservation.

## Worksheet Operations
- Freeze and unfreeze panes.
- Set and remove auto-filters.
- Group and ungroup rows and columns; set row heights and column widths.
- Set sheet tab colour; hide, unhide, and move sheets.
- Toggle gridlines.
- Configure print area, page setup, margins, header/footer, and print titles.
- Add and remove page breaks.
- Copy ranges and sheets across workbooks; stack sheets into a consolidated sheet.

## Document Properties
- Read and write workbook metadata (author, title, etc.).
- Set calculation mode.

## Cross-file Operations
- Aggregate or filter data across multiple files in bulk.
- Validate data consistency across multiple workbooks.

**66 Tools | 2 Resources | 20 Prompts**


# Prerequisites
- Python 3.12+
- [uv](https://github.com/astral-sh/uv) for dependency management and execution.

# Setup and run
```sh
uv sync
uv run src/mcp_server/main.py   # or: uv run excel-mcp
```

## Try it with the MCP Inspector
Inspect tools and resources without an LLM in the loop:
```sh
npx @modelcontextprotocol/inspector uv run --directory . src/mcp_server/main.py
```

## Architecture
- **Modular design**: 25 independent domain modules under `src/mcp_server/tools/`; 15 route modules under `src/mcp_server/routes/` handle parameter validation and action dispatch.
- **Entry point**: `src/mcp_server/main.py` calls `register_all_routes(mcp)` to register 66 tools and 2 resources (`excel://workbook/{file_path}/sheets`, `excel://workbook/{file_path}/sheet/{sheet_name}/preview`). It contains no `@mcp.tool()` decorators directly.
- **Data flow**: MCP Client → FastMCP (stdio/JSON-RPC) → `main.py` → `routes/*.py` (action dispatch, validation) → `tools/*.py` (pure domain logic) → Excel/CSV file operations → Pydantic response models → JSON response.
- **Utilities**: `src/mcp_server/utils/excel_helpers.py` provides safe file handling; `logger.py` routes all logs to stderr only (MCP stdio-safe).
- **Safety**: File paths validated against an extension whitelist (`.xlsx`, `.xls`, `.csv`, `.xlsm`); `goal_seek` and `scenario_analysis` use AST whitelist validation; `add_computed_column` uses a blocklist to block dangerous patterns; image paths are resolved and extension-checked.
- **Workbook lifecycle**: All tools use `load_workbook_safe()` and `save_workbook_safe()` with `try/finally wb.close()` for reliable resource cleanup.

## Extending
- Add new tool functions under `src/mcp_server/tools/` (pure functions, no decorators).
- Add a route wrapper in the appropriate `src/mcp_server/routes/` module, or create a new route module and register it in `src/mcp_server/routes/__init__.py`.
- Add tests in `tests/` using pytest and `tmp_path` fixture.
- Run `uv run ruff check src/ tests/` to lint and `uv run pytest -v` to test.

## Tests and linting
```sh
uv run pytest -v
uv run ruff check src/ tests/
```

## Prompts
The server ships 20 MCP prompts (`prompts.py`) for common workflows:

| Prompt                            | Description                                                     |
| --------------------------------- | --------------------------------------------------------------- |
| `excel-quickstart`                | Create a new formatted workbook and auto-fit columns            |
| `excel-data-analysis`             | Profile, filter, aggregate, sort, and find duplicates           |
| `excel-data-cleaning`             | Full cleaning pipeline: profile, clean, deduplicate, export     |
| `excel-chart-builder`             | Create and annotate charts with trendlines, labels, and legends |
| `excel-report-builder`            | Multi-sheet formatted report with data, charts, and print setup |
| `excel-financial-model`           | Loan amortisation, DCF, and financial ratio workbook            |
| `excel-pivot-etl`                 | Pivot tables and ETL: merge, unpivot, computed columns          |
| `excel-multi-file`                | Aggregate, filter, compare, and validate across multiple files  |
| `excel-statistical-analysis`      | OLS regression and exponential smoothing with forecasting       |
| `excel-formula-builder`           | Write, fill, auto-sum, and audit formulas                       |
| `excel-data-governance`           | Data validation, protection, named ranges, and scenarios        |
| `excel-csv-workflow`              | Preview CSV, convert to Excel, clean, analyse, export           |
| `excel-readonly-audit`            | Inspect structure, formulas, tables, and data quality read-only |
| `excel-formula-diagnosis`         | Diagnose formula errors and trace precedents/dependents         |
| `excel-workbook-maintenance`      | Sheet layout, print setup, and sizing housekeeping              |
| `excel-table-manager`             | Create, inspect, resize, total, and convert Excel tables        |
| `excel-what-if-analysis`          | Goal seek, solver, sensitivity, and scenario analysis           |
| `excel-multi-file-reconciliation` | Validate, compare, aggregate, and filter multiple files         |
| `excel-search-repair`             | Find and repair text or formula content                         |
| `excel-safe-transform`            | Sandboxed custom transform via `execute_custom_code`            |


## Tools

This section merges the previous "Tool Overview" and the full grouped tool list. Each group includes a short description of its purpose and the individual tools provided by the server.

- Workbook management (5) — Manage workbooks and sheets (create, inspect, rename, delete, copy, move)
  - `get_workbook_metadata` — Return workbook metadata (sheets, active sheet, named ranges).
  - `create_workbook` — Create a new workbook file with optional initial sheets.
  - `get_sheet_summary` — Summarise a sheet (header detection, used range, dimensions).
  - `write_multi_sheet` — Create/overwrite a workbook from multiple sheet definitions.
  - `sheet_management` — Rename, delete, copy, hide/unhide, tab color, and move sheets.

- Cell & range operations (6) — Read and write cells/ranges and perform range transforms
  - `read_cells` — Read a single cell, a rectangular range, or stream large sheets in chunks.
  - `write_cells` — Write a single cell, a 2D range, generate series, merge/unmerge ranges.
  - `clear_range` — Clear values from a rectangular range.
  - `copy_range` — Copy a range between sheets (values and/or styles, optional paste-values-only).
  - `find_replace` — Find and replace text (optionally in formulas) across a sheet.
  - `transpose_range` — Transpose rows↔columns and write the result at a target cell.

- Formatting & styling (5) — Apply and manage cell formatting at scale
  - `format_cells` — Apply fonts, fills, alignment, borders and number formats to a range.
  - `auto_fit_columns` — Auto-fit column widths to their contents.
  - `copy_cell_format` — Copy formatting from a source cell to every cell in a target range.
  - `clear_cell_format` — Remove formatting from a range without changing values.
  - `apply_named_style` — Apply built-in Excel named styles (e.g. Heading, Good, Bad).

- Formulas (2) — Write and audit formulas
  - `formula_write` — Set single formulas, batch-set, drag-fill, or insert AutoSum formulas.
  - `formula_audit` — Inspect formula values, list errors, find precedents/dependents, or list formulas.

- Charts (1) — Create and manage chart lifecycle and series
  - `chart` — Create, list, delete, add series, configure axes/trendlines/legends/data labels.

- Worksheet operations (4) — UI, structure, print and cross-workbook transfers
  - `worksheet_view` — Freeze panes, set/remove auto-filter, toggle gridlines.
  - `worksheet_structure` — Insert/delete rows/cols, group/ungroup, set sizes.
  - `worksheet_print` — Set print area, page setup, print titles, and page breaks.
  - `worksheet_transfer` — Copy ranges/sheets across workbooks, merge workbooks, stack sheets.

- Data analysis (9) — Filtering, aggregation, profiling and helper utilities
  - `sort_data` — Sort worksheet rows by one or more columns.
  - `column_statistics` — Compute descriptive statistics for a numeric column.
  - `aggregate_data` — Group rows and aggregate values using common operations.
  - `find_duplicates` — Identify duplicate rows based on a set of columns.
  - `vlookup_helper` — Cross-file lookup helper (exact or fuzzy matching) to enrich data.
  - `filter_data_advanced` — Filter rows using multiple conditions combined with AND/OR logic.
  - `insert_subtotals` — Insert SUBTOTAL formula rows after groups in a sorted sheet.
  - `profile_data` — Produce a per-column data profile (types, nulls, unique counts, samples).
  - `value_counts` — Frequency counts (optionally normalized, top-n) for a column.

- Pivot & ETL (6) — Pivot tables and ETL-style transforms
  - `create_pivot_table` — Build a pivot table and optionally write it to a sheet/file.
  - `refresh_pivot_table` — Re-run a stored pivot definition to refresh output.
  - `unpivot_data` — Melt wide-form data into long-form (id_vars/value_vars).
  - `merge_datasets` — Join two sheets using SQL-style join semantics.
  - `add_computed_column` — Add a computed column via safe expressions or cumsum/rolling operations.
  - `deduplicate_data` — Remove duplicate rows (with keep strategy) from a sheet.

- Financial (8) — Time-value calculations, DCF, goal-seek and scenario tools
  - `goal_seek` — Solve for a variable cell value that makes an expression equal a target.
  - `loan_amortization` — Generate an amortization schedule for a loan.
  - `dcf_analysis` — Discounted cash flow valuation with terminal value calculation.
  - `budget_variance_analysis` — Compare budget vs actual by category and report variances.
  - `financial_ratio_analysis` — Compute common financial ratios and compare to benchmarks.
  - `break_even_analysis` — Compute break-even units and revenue from cost structure.
  - `create_sensitivity_table` — Build 1- or 2-variable sensitivity tables in the workbook.
  - `time_value_calc` — FV/PV/NPER/RATE/depreciation/IRR operations and helpers.

- Cleaning & CSV (4) — Data cleaning primitives and CSV helpers
  - `split_column` — Split a delimited text column into multiple columns.
  - `data_cleaner` — Run a configurable cleaning pipeline (trim, dedupe, fill, normalize).
  - `parse_date_column` — Parse varied date formats and normalise output formatting.
  - `csv_ops` — Preview CSVs and convert between CSV and XLSX.

- Statistical & solver (4) — Regression, smoothing and optimisation
  - `run_regression` — Run OLS regression and return coefficients and diagnostics.
  - `run_exponential_smoothing` — Apply simple/Holt/Holt-Winters smoothing and optional forecasting.
  - `run_solver` — Constrained optimisation using scipy for minimisation/maximisation objectives.
  - `correlation_matrix` — Compute Pearson correlation matrix for numeric columns.

- Governance (4) — Protection, validation, document properties and conditional formats
  - `protection` — Protect/unprotect sheets or workbooks and lock cell ranges.
  - `data_validation` — Add/remove dropdown, numeric, date, and formula-based validation rules.
  - `doc_properties` — Read document properties or set calculation mode.
  - `conditional_format` — Apply, list, or remove conditional formatting rules.

- Metadata & tables (5) — Comments, hyperlinks, scenarios, named ranges and tables
  - `comment` — Add, read, delete or list comments on cells.
  - `hyperlink` — Add, read, delete or list hyperlinks attached to cells.
  - `scenario` — Save, list and apply what-if scenarios persisted in a hidden sheet.
  - `named_range` — List, create, delete or update named ranges.
  - `table` — Create, list, resize, toggle totals, read, or convert Excel tables.

- Multi-file operations (1) — Bulk and cross-workbook operations
  - `multi_file` — Aggregate, filter, validate, or compare across multiple workbooks.

- Custom code & images (2) — Sandboxed code execution and images
  - `insert_image` — Insert an image into a worksheet anchored at a target cell.
  - `execute_custom_code` — Run sandboxed Python/pandas code against a workbook and return results.

 