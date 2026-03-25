# Excel MCP Server

A Python-based Model Context Protocol (MCP) server for Excel automation and data manipulation. Enables LLMs to interact with Excel files (`.xlsx`, `.xls`, `.csv`, `.xlsm`) through 177 structured tools, supporting workflows from basic spreadsheet operations to advanced financial modelling and data analysis.

# Features

## Workbook & Sheet Management
- Create workbooks and list, rename, delete, copy, and move sheets.
- Retrieve workbook metadata and per-sheet summaries.
- Write data to multiple sheets in a single call.

## Cell & Range Operations
- Read and write individual cells and contiguous ranges.
- Read detailed cell metadata (type, style, formula, value).
- Batch-read multiple ranges in one call.
- Copy and delete ranges.
- Read large files in chunks using the python-calamine engine.
- Retrieve file-level information (size, sheet count, etc.).

## Row & Column Operations
- Insert and delete rows and columns by index or letter.

## Formatting & Styling
- Apply fonts, colours, alignment, borders, and number formats to cells.
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
- Validate formula syntax.
- List all formulas in a sheet.
- Convert formulas to static values.

## Tables
- Create, list, rename, resize, and delete native Excel tables.
- Toggle the totals row and read table data.

## Data Validation
- Add dropdown, numeric, date, text-length, and formula-based validation rules.
- List and remove validation rules.

## Protection
- Protect and unprotect sheets.
- Lock individual cells.

## Charts
- Create 10 chart types: column, bar, line, pie, scatter, area, radar, doughnut, bubble, and stock.
- Update chart properties.
- Add and remove chart series.
- List and delete charts.

## Data Analysis
- Filter data by single or multiple conditions (==, !=, >, <, contains, startswith, etc.).
- Sort by one or more columns.
- Compute column statistics (mean, median, min, max, std, sum).
- Aggregate and group data.
- Find and remove duplicate rows.
- Profile a dataset comprehensively.
- Search and replace values, including regex-based replacement.
- Calculate correlation between columns.
- Rank and calculate percentiles for data.
- Sample rows randomly or systematically.
- Create frequency histograms.
- Transpose data ranges.
- Extract unique values from a column.
- VLOOKUP-style helper across ranges.
- Find cells matching specific formatting criteria.
- Normalise numeric data.
- Export analysis results to a new file.

## CSV Operations
- Preview CSV content.
- Detect CSV dialect automatically.
- Validate CSV structure.
- Convert between CSV and XLSX formats.

## Pivot & ETL
- Create pivot tables.
- Unpivot (melt) data from wide to long format.
- Merge datasets using SQL-style joins.
- Add computed columns with safe expression evaluation.
- Deduplicate rows.
- Append datasets from multiple sources.
- Find differences between two datasets.

## Financial Calculations
- NPV, IRR, XNPV, and XIRR.
- PMT (periodic payment).
- DCF (discounted cash flow) analysis.
- Loan amortisation schedules.
- Goal seek with AST-validated expressions.
- Budget variance analysis.
- Financial ratio analysis.
- Scenario analysis.
- Trend analysis and CAGR.
- Break-even analysis.

## Data Cleaning
- Configurable cleaning pipeline: trim whitespace, remove empty rows/columns, normalise values, fix number formats, deduplicate, and fill missing values.
- Preview mode (dry run without saving).
- Split a column into multiple columns.
- Combine multiple columns into one.
- Detect outliers using IQR or z-score methods.

## Comments
- Add, read, update, and delete comments.
- List all comments in a sheet.
- Add and delete comments in bulk.

## Hyperlinks
- Add external and internal (intra-workbook) hyperlinks.
- Read, delete, and list hyperlinks.

## Images
- Insert, list, and delete images (`.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`).

## Named Ranges
- Create, list, rename, update, and delete named ranges with scope preservation.

## Worksheet Operations
- Freeze and unfreeze panes.
- Set and remove auto-filters.
- Hide and unhide rows and columns.
- Group and ungroup rows and columns.
- Set sheet tab colour.
- Hide, unhide, and move sheets.
- Set zoom level and toggle gridlines.
- Configure print area, page setup, margins, header/footer, and print titles.
- Insert, delete, and list page breaks.

## Document Properties
- Read and write workbook metadata (author, title, etc.).
- Set calculation mode.

## Cross-file Operations
- Aggregate or filter data across multiple files in bulk.
- Validate data consistency across multiple workbooks.

**177 Tools | 2 Resources**


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
- **Modular design**: 22 independent tool modules under `src/mcp_server/tools/`; pure functions with thin `@mcp.tool()` wrappers in `main.py`.
- **Entry point**: `src/mcp_server/main.py` registers 177 tools and 2 resources (`excel://workbook/{file_path}/sheets`, `excel://workbook/{file_path}/sheet/{sheet_name}/preview`).
- **Data flow**: MCP Client → FastMCP (stdio/JSON-RPC) → `main.py` (tool dispatch) → `tools/` modules → Excel/CSV file operations → Pydantic response models → JSON response.
- **Utilities**: `src/mcp_server/utils/excel_helpers.py` provides safe file handling; `logger.py` routes all logs to stderr only (MCP stdio-safe).
- **Safety**: File paths validated against an extension whitelist (`.xlsx`, `.xls`, `.csv`, `.xlsm`); `goal_seek` and `scenario_analysis` use AST whitelist validation; `add_computed_column` uses a blocklist to block dangerous patterns; image paths are resolved and extension-checked.
- **Workbook lifecycle**: All tools use `load_workbook_safe()` and `save_workbook_safe()` with `try/finally wb.close()` for reliable resource cleanup.

## Extending
- Add new tool functions under `src/mcp_server/tools/` (pure functions, no decorators).
- Register wrappers in `src/mcp_server/main.py` using `@mcp.tool()` decorator.
- Add tests in `tests/` using pytest and `tmp_path` fixture.
- Run `uv run ruff check src/ tests/` to lint and `uv run pytest -v` to test.

## Tests and linting
```sh
uv run pytest -v          # 254 tests across 17 test files
uv run ruff check src/ tests/
```

## Tool Overview
- **Workbook Management** (8 tools): Create workbooks; list, copy, rename, and delete sheets; extract metadata and summaries; write to multiple sheets.
- **Cell Operations** (11 tools): Read and write individual cells and ranges; clear, copy, and delete ranges; chunked reading for large files; batch range reads; detailed cell metadata.
- **Row/Column Operations** (6 tools): Insert and delete rows and columns by index or letter.
- **Formatting** (11 tools): Font, fill, alignment, borders, and number formats; merge/unmerge cells; auto-fit columns; gradient fills; per-cell format arrays; copy formatting between ranges.
- **Conditional Formatting** (8 tools): Colour scales, data bars, icon sets, highlight rules, formula rules, top/bottom rules, above/below average rules, duplicate highlighting.
- **Formulas** (6 tools): Set single, array, and batch formulas; validate syntax; list formulas; convert formulas to values.
- **Tables** (7 tools): Create, list, delete, rename, and resize native Excel tables; manage totals rows; read table data.
- **Data Validation** (7 tools): Dropdown lists, numeric constraints, date rules, text-length rules, and formula-based validation; list and remove rules.
- **Protection** (3 tools): Protect and unprotect sheets; lock individual cells.
- **Charts** (6 tools): Create charts (10 types), delete, list, update properties, and add or remove series.
- **Data Analysis** (19 tools): Simple and advanced filtering, sorting, statistics, aggregation, duplicate detection, profiling, correlation, ranking, percentiles, sampling, histograms, transposition, unique values, VLOOKUP helper, format-based cell search, normalisation, search/replace (regex-capable), and analysis export.
- **CSV Operations** (5 tools): Preview CSV files, detect dialect, validate CSV structure, and convert between CSV and XLSX.
- **Pivot & ETL** (7 tools): Create pivot tables, unpivot data, merge datasets (SQL-style joins), add computed columns, deduplicate, append datasets, and find differences.
- **Financial** (14 tools): NPV, IRR, PMT, XNPV, XIRR, DCF analysis, budget variance, financial ratios, scenario analysis, trend analysis, CAGR, break-even analysis, goal seek (AST-safe), and loan amortisation.
- **Data Cleaning** (4 tools): Configurable cleaning pipeline (trim, remove empties, normalise, fix numbers, deduplication, fill missing values); split and combine columns; detect outliers.
- **Comments** (7 tools): Add, update, read, delete, and list cell comments; bulk add and delete.
- **Document Properties** (5 tools): Read and write workbook metadata; protect and unprotect workbooks; set calculation mode.
- **Hyperlinks** (5 tools): Add, read, delete, and list external hyperlinks; add internal (intra-workbook) hyperlinks.
- **Images** (3 tools): Insert, list, and delete images embedded in worksheets.
- **Cross-file Operations** (3 tools): Bulk aggregate and filter across multiple files; validate data consistency across workbooks.
- **Named Ranges** (5 tools): List, create, delete, update, and rename named ranges with scope preservation.
- **Worksheet Operations** (26 tools): Freeze/unfreeze panes, auto-filter, hide/unhide rows and columns, grouping/ungrouping, tab colours, zoom, gridlines, print area, page setup, margins, header/footer, page breaks, and print titles.
