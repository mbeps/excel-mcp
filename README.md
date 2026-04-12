# Excel MCP Server

A Python-based Model Context Protocol (MCP) server for Excel and CSV automation. It enables LLMs to work with `.xlsx`, `.xlsm`, `.xls`, and `.csv` files through 66 structured tools, 2 JSON resources, and 20 prompts spanning workbook operations, formatting, charts, ETL, analysis, and financial/statistical workflows.

# Features

> **NOTE:** Custom actions are not supported as first-class tools. Use `execute_custom_code` as the sanctioned, sandboxed path for custom data transformations. This uses AST validation and blocks dangerous operations.

## Workbook & Sheet Management
- Create workbooks; inspect workbook metadata and per-sheet summaries.
- Rename, delete, copy, move, hide/unhide, and recolour sheet tabs.
- Write data to multiple sheets in a single call.

## Cell & Range Operations
- Read and write individual cells, contiguous ranges, and large sheets in chunks.
- Include formulas and cell metadata when reading ranges.
- Copy ranges, clear range values, transpose data, and search/replace with optional regex support.

## Row & Column Operations
- Insert and delete rows and columns by index.

## Formatting & Styling
- Apply fonts, colours, alignment, borders, and number formats to cells.
- Use built-in number format presets and named Excel styles.
- Merge and unmerge cells, auto-fit columns, clear formatting, and copy cell formatting.

## Conditional Formatting
- Apply colour scales, data bars, icon sets, and targeted highlight rules.
- Add formula, top/bottom, and above/below-average rules; list and remove rules.

## Formulas
- Set single or batch formulas and drag-fill them across ranges with relative-reference translation.
- Insert AutoSum formulas and audit formula values, errors, precedents, dependents, and sheet formulas.

## Tables
- Create, list, resize, and read native Excel tables.
- Toggle totals rows and convert tables back to plain ranges.

## Data Validation
- Add dropdown, numeric, date, and formula-based validation rules.
- Remove validation rules.

## Protection
- Protect and unprotect sheets and workbooks.
- Lock individual cells.

## Charts
- Create, list, delete, and update charts across 10 chart types: column, bar, line, pie, scatter, area, radar, doughnut, bubble, and stock.
- Add series and configure axes, trendlines, combo charts, data labels, legends, and category ranges.

## Data Analysis
- Filter and sort data by multiple conditions or columns.
- Compute column statistics, grouped aggregates, value counts, and correlation matrices.
- Find duplicates, profile datasets, insert subtotals, and perform VLOOKUP-style enrichment.
- Run OLS regression and exponential smoothing with forecasting support.

## CSV Operations
- Preview CSV content.
- Convert between CSV and XLSX formats.

## Pivot & ETL
- Create pivot tables and refresh them from stored definitions.
- Unpivot wide data, merge datasets with SQL-style joins, and deduplicate rows.
- Add computed columns with safe expressions, cumulative sums, or rolling calculations.

## Financial Calculations
- Run IRR, FV, PV, NPER, RATE, and depreciation time-value calculations.
- Build DCF models, loan amortisation schedules, budget variance reports, financial ratio analysis, and break-even analysis.
- Perform goal seek, constrained solver optimisation, sensitivity tables, and scenario analysis.

## Data Cleaning
- Run a configurable cleaning pipeline to trim whitespace, remove empty rows/columns, normalise text, fix number formats, deduplicate, and fill missing values.
- Use preview mode before saving changes.
- Split columns and parse/normalise date columns.

## Comments
- Add, read, delete, and list comments.

## Hyperlinks
- Add internal or external hyperlinks.
- Read, delete, and list hyperlinks.

## Images
- Insert images into worksheets.

## Named Ranges
- Create, list, update, and delete named ranges.

## Worksheet Operations
- Freeze and unfreeze panes.
- Set/remove auto-filters, and toggle gridlines.
- Group/ungroup rows and columns; set row heights and column widths.
- Configure print area, page setup, print titles, and manual page breaks.
- Copy ranges across sheets, copy sheets across workbooks, merge workbooks, and stack sheets into a consolidated sheet.

## Document Properties
- Read workbook metadata and set calculation mode.

## Cross-file Operations
- Aggregate, filter, validate, and compare data across multiple files.

**66 Tools | 2 Resources | 20 Prompts**


# Prerequisites
Below are the requirements for running this MCP:
- Python 3.12+
- [uv](https://github.com/astral-sh/uv) for dependency management and execution
- Optional: Node.js (for running the MCP Inspector via `npx`)

# Setup & Run
Installing all project dependencies:
```sh
uv sync
```

Running the MCP server:
```sh
uv run src/mcp_server/main.py   # or: uv run excel-mcp
```

## Using with MCP clients:
- ### GitHub Copilot (Visual Studio Code)
In the `.vscode/mcp.json` add:
```json
"excel-mcp": {
  "command": "uv",
  "args": [
    "run",
    "--with",
    "mcp[cli]",
    "excel-mcp"
  ],
  "env": {
    "PYTHONPATH": "src"
  }
}
```

- ### Anthropic Claude
Expose the FastMCP server via the Streamable HTTP transport and register it with Claude/Claude Code for direct tool use. For development you can also use the MCP Inspector or run a small HTTP adapter that forwards Claude Messages to the MCP server.

Example (from research spec):
```sh
claude mcp add --transport http my-excel-mcp http://localhost:8000/mcp
```

- ### Anthropic Claude Code
Claude Code supports stdio, HTTP/SSE, and plugin-bundled MCP servers. For local development add a project-scoped `.mcp.json` or register a stdio command; for sharing use a plugin that bundles a `.mcp.json` entry.

Example (project `.mcp.json` from research spec):
```json
{
  "mcpServers": {
    "excel-mcp": {
      "type": "stdio",
      "command": "uv",
      "args": ["run", "src/mcp_server/main.py"],
      "env": { "PYTHONPATH": "src" }
    }
  }
}
```

- ### OpenAI / ChatGPT
OpenAI supports an `mcp` tool via the Responses API (remote MCP/SSE or streamable HTTP), ChatGPT plugins (OpenAPI + `ai-plugin.json`), or a local function-calling wrapper that mediates calls to the local MCP server. Protect public endpoints with TLS/auth or prefer the local wrapper for private data.

Example (Responses API `mcp` tool from research spec):
```py
from openai import OpenAI
client = OpenAI()

resp = client.responses.create(
    model="gpt-5",
    tools=[{
        "type": "mcp",
        "server_label": "excel-mcp",
        "server_description": "Excel MCP server",
        "server_url": "https://your-public-host.example.com/sse",
        "authorization": "Bearer YOUR_MCP_TOKEN",
        "require_approval": "always",
    }],
    input="Read first 20 rows from /tmp/demo.xlsx"
)
print(resp.output_text)
```

## Try it with the MCP Inspector
Inspect tools, resources, and prompts without an LLM in the loop:
```sh
npx @modelcontextprotocol/inspector uv run --directory . src/mcp_server/main.py
```

## Verifications
Run tests with:
```sh
uv run pytest -v
```

Run linting with:
```sh
uv run ruff check src/ tests/
```

Run MyPy type checks with:
```sh
uv run mypy src/mcp_server
```

# Prompts
The server ships 20 MCP prompts (`src/mcp_server/prompts.py`) for common workflows:

| Prompt                            | Description                                                                                 |
| --------------------------------- | ------------------------------------------------------------------------------------------- |
| `excel-quickstart`                | Create a new Excel workbook, populate it with formatted data, and auto-fit columns          |
| `excel-data-analysis`             | Profile, filter, aggregate, sort, and find duplicates in an Excel dataset                   |
| `excel-data-cleaning`             | Full data cleaning pipeline: profile, clean, validate, deduplicate, and export              |
| `excel-chart-builder`             | Create, style, and annotate charts with trendlines, axis labels, data labels, and legends   |
| `excel-report-builder`            | Assemble a multi-sheet formatted report with data, charts, statistics, and print setup      |
| `excel-financial-model`           | Build a financial model workbook with loan amortisation, DCF analysis, and financial ratios |
| `excel-pivot-etl`                 | Create pivot tables and run ETL transforms: merge, unpivot, add computed columns            |
| `excel-multi-file`                | Aggregate, filter, compare, and validate schema consistency across multiple Excel files     |
| `excel-statistical-analysis`      | Run OLS regression and exponential smoothing with forecasting on Excel data                 |
| `excel-formula-builder`           | Write, fill, auto-sum, and audit Excel formulas across a sheet                              |
| `excel-data-governance`           | Apply data validation, protection, named ranges, and scenario management to a workbook      |
| `excel-csv-workflow`              | Preview a CSV, convert to Excel, clean, analyse, and export results                         |
| `excel-readonly-audit`            | Inspect workbook structure, formulas, tables, and data quality without mutating the file    |
| `excel-formula-diagnosis`         | Diagnose formula errors, trace precedents and dependents, and apply the smallest safe fix   |
| `excel-workbook-maintenance`      | Perform conservative workbook housekeeping: sheets, layout, print setup, and sizing         |
| `excel-table-manager`             | Create, inspect, resize, total, and convert native Excel tables                             |
| `excel-what-if-analysis`          | Run goal seek, solver, sensitivity, and scenario analysis                                   |
| `excel-multi-file-reconciliation` | Validate, compare, aggregate, and filter multiple files with schema checks                  |
| `excel-search-repair`             | Find and repair text or formula content with a read-before-write workflow                   |
| `excel-safe-transform`            | Apply a sandboxed custom transform only when built-in tools are not enough                  |


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
  - `chart` — Create, list, delete, update, add series, and configure axes, trendlines, combo charts, legends, and data labels.

- Worksheet operations (4) — UI, structure, print and cross-workbook transfers
  - `worksheet_view` — Freeze panes, set/remove auto-filter, toggle gridlines.
  - `worksheet_structure` — Insert/delete rows/cols, group/ungroup, set sizes.
  - `worksheet_print` — Set print area, page setup, print titles, and page breaks.
  - `worksheet_transfer` — Copy ranges across sheets within a workbook, copy sheets across workbooks, merge workbooks, and stack sheets.

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

- Financial (8) — Time-value calculations, DCF, goal-seek and financial modelling tools
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

 
# Architecture
- **Modular design**: `src/mcp_server/` is split into `tools/` (25 pure domain modules), `routes/` (15 registration/dispatch modules), `models/` (Pydantic response schemas), and `utils/` (shared workbook, logging, and expression-safety helpers).
- **Entry point**: `src/mcp_server/main.py` creates the FastMCP server, registers 2 JSON resources directly (`excel://workbook/{file_path}/sheets`, `excel://workbook/{file_path}/sheet/{sheet_name}/preview`), calls `register_all_routes(mcp)` to register the tool surface, and calls `_register_prompts(mcp)` from `src/mcp_server/prompts.py` to register 20 prompts.
- **Data flow**: MCP client → FastMCP (stdio/JSON-RPC) → `main.py` → `routes/*.py` (registration/dispatch) → `tools/*.py` (domain logic) → `openpyxl` / `pandas` / `scipy` and related libraries → Pydantic models → JSON-RPC response.
- **Utilities**: `src/mcp_server/utils/excel_helpers.py` centralises safe workbook access and workbook path validation; `logger.py` keeps logs on stderr; `expression_validator.py` provides shared AST validation for user-supplied expressions.
- **Safety**: Workbook paths are checked against an extension whitelist and optional `EXCEL_MCP_ALLOWED_DIRS` sandbox; AST validation is reused by `goal_seek`, `create_sensitivity_table`, and computed-column expressions; `execute_custom_code` uses a separate sandboxed validation path.
- **Workbook lifecycle**: Openpyxl-backed workbook tools generally use `load_workbook_safe()` / `save_workbook_safe()` with explicit close handling, while pandas/CSV flows and hidden-sheet state (`_mcp_pivots`, `_mcp_scenarios`) follow separate storage paths.

## Extending
- Add new tool functions under `src/mcp_server/tools/` (pure functions, no decorators).
- Expose that logic through an existing `src/mcp_server/routes/` module, or add a new route module with a `register(mcp)` function and include it in `register_all_routes()` in `src/mcp_server/routes/__init__.py`.
- Add tests in `tests/` using pytest and the `tmp_path` fixture where appropriate.
- Run `uv run ruff check src/ tests/` to lint and `uv run pytest -v` to test.

## References
- [Model Context Protocol specification](https://modelcontextprotocol.io/specification)
- [FastMCP Python SDK](https://github.com/modelcontextprotocol/python-sdk)
- [uv documentation](https://docs.astral.sh/uv/)
- [openpyxl documentation](https://openpyxl.readthedocs.io/)
- [pandas documentation](https://pandas.pydata.org/docs/)
- [python-calamine](https://github.com/dimastbk/python-calamine)
- [numpy-financial](https://numpy.org/numpy-financial/)
- [SciPy documentation](https://docs.scipy.org/)
- [statsmodels documentation](https://www.statsmodels.org/)
- [Pydantic documentation](https://docs.pydantic.dev/)
- [Pillow documentation](https://pillow.readthedocs.io/)