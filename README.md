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

## Tool Overview
- **Workbook Management** (8 tools): Create workbooks; list, copy, rename, and delete sheets; extract metadata and summaries; write to multiple sheets.
- **Cell Operations** (11 tools): Read and write individual cells and ranges; clear, copy, and delete ranges; chunked reading for large files; batch range reads; detailed cell metadata.
- **Row/Column Operations** (6 tools): Insert and delete rows and columns by index or letter.
- **Formatting** (5 tools): Font, fill, alignment, borders, and number formats; apply named Excel styles; number format presets; clear formatting; merge/unmerge cells; auto-fit columns; gradient fills; per-cell format arrays; copy formatting between ranges.
- **Conditional Formatting** (8 tools): Colour scales, data bars, icon sets, highlight rules, formula rules, top/bottom rules, above/below average rules, duplicate highlighting.
- **Formulas** (9 actions): Set single, array, and batch formulas; drag-fill formulas across a range; auto-sum; validate syntax; list formulas; convert formulas to values.
- **Tables** (6 actions): Create, list, delete, rename, and resize native Excel tables; manage totals rows; read table data; convert tables to plain ranges.
- **Data Validation** (7 tools): Dropdown lists, numeric constraints, date rules, text-length rules, and formula-based validation; list and remove rules.
- **Protection** (3 tools): Protect and unprotect sheets; lock individual cells.
- **Charts** (9 actions): Create charts (10 types), delete, list, configure axes, trendlines and combo charts, add and configure data labels, manage legends, and add or remove series.
- **Data Analysis** (20 tools): Simple and advanced filtering, sorting, statistics, aggregation, duplicate detection, profiling (with percentiles and IQR), subtotals insertion, correlation, ranking, percentiles, sampling, histograms, transposition, unique values, VLOOKUP helper, format-based cell search, normalisation, search/replace (regex-capable), and analysis export.
- **CSV Operations** (5 tools): Preview CSV files, detect dialect, validate CSV structure, and convert between CSV and XLSX.
- **Pivot & ETL** (6 tools): Create and refresh pivot tables, unpivot data, merge datasets (SQL-style joins), add computed columns (expression-based or cumulative sum), deduplicate, append datasets, and find differences.
- **Financial** (14 tools): NPV, IRR, PMT, XNPV, XIRR, DCF analysis, budget variance, financial ratios, scenario analysis, trend analysis, CAGR, break-even analysis, goal seek (AST-safe), and loan amortisation.
- **Data Cleaning** (3 tools): Configurable cleaning pipeline (trim, remove empties, normalise, fix numbers, deduplication, fill missing values); split and combine columns; parse and normalise date columns; detect outliers.
- **Comments** (7 tools): Add, update, read, delete, and list cell comments; bulk add and delete.
- **Document Properties** (5 tools): Read and write workbook metadata; protect and unprotect workbooks; set calculation mode.
- **Hyperlinks** (5 tools): Add, read, delete, and list external hyperlinks; add internal (intra-workbook) hyperlinks.
- **Images** (3 tools): Insert, list, and delete images embedded in worksheets.
- **Cross-file Operations** (3 tools): Bulk aggregate and filter across multiple files; validate data consistency across workbooks.
- **Named Ranges** (5 tools): List, create, delete, update, and rename named ranges with scope preservation.
- **Worksheet Operations** (19 actions): Freeze/unfreeze panes, auto-filter, hide/unhide rows and columns, grouping/ungrouping, set row heights and column widths, tab colours, zoom, gridlines, print area, page setup, margins, header/footer, page breaks, print titles, and stack sheets.
