# Excel MCP Server

A Python-based Model Context Protocol (MCP) server for Excel automation and data manipulation. Enables LLMs to interact with Excel files (`.xlsx`, `.xls`, `.csv`) through 50 structured tools, supporting workflows from basic spreadsheet operations to advanced financial modelling and data analysis.

# Features

## Core Operations
- Create, list, rename, delete, and copy worksheets
- Read/write individual cells and ranges
- Insert/delete rows and columns
- Chunked reading for large files (calamine engine)

## Formatting & Styling
- Apply fonts, colors, alignment, borders, number formats
- Merge/unmerge cells
- Set column widths and row heights
- Auto-fit columns based on content

## Advanced Features
- **Conditional Formatting**: Color scales, data bars, icon sets, highlight rules
- **Tables**: Create and list native Excel tables
- **Data Validation**: Dropdown lists, numeric constraints, formula-based rules
- **Protection**: Sheet protection with granular cell locking
- **Formulas**: Single, array, and batch formula operations with syntax validation

## Data Analysis
- Filter by multiple operators (==, !=, >, <, contains, startswith, etc.)
- Sort by single or multiple columns
- Column statistics (mean, median, min, max, std, sum)
- Group and aggregate data
- Find duplicates
- Profile data comprehensively
- Search and replace

## Charting & Visualization
- Create charts (bar, column, line, pie, scatter, area)
- List and delete charts

## ETL & Transformation
- Pivot tables (static via pandas)
- Unpivot/melt data
- Merge datasets (SQL-style joins)
- Add computed columns with safe expression evaluation
- Deduplicate rows

## Financial Calculations
- NPV (Net Present Value)
- IRR (Internal Rate of Return)
- PMT (Periodic Payment)
- Goal Seek with AST-safe expression evaluation
- Loan amortization schedules

## CSV Integration
- Preview CSV files
- Convert CSV ↔ XLSX

**62 Tools | 2 Resources | 63 Tests Passing**
## Prerequisites
- Python 3.12+
- [uv](https://github.com/astral-sh/uv) for dependency management and execution.

## Setup and run
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
- **Modular design**: 14 independent tool modules, pure functions with thin MCP wrappers.
- **Entry point**: `src/mcp_server/main.py` registers all 50 tools and 2 resources.
- **Utilities**: `src/mcp_server/utils/excel_helpers.py` provides safe file handling; `logger.py` ensures stderr-only logging.
- **Safety**: Path validation (extension whitelist), expression safety (AST/blocklist validation), proper workbook lifecycle management.

## Extending
- Add new tool functions under `src/mcp_server/tools/` (pure functions, no decorators).
- Register wrappers in `src/mcp_server/main.py` using `@mcp.tool()` decorator.
- Add tests in `tests/` using pytest and `tmp_path` fixture.
- Run `uv run ruff check src/ tests/` to lint and `uv run pytest -v` to test.

## Tests and linting
```sh
uv run pytest -v          # 63 tests across 11 test files
uv run ruff check src/ tests/
```

## Tool Overview
- **Workbook Management** (7 tools): Create, list, copy, rename, delete sheets; metadata extraction.
- **Cell Operations** (6 tools): Read/write cells and ranges; chunked reading for large files (via python-calamine).
- **Row/Column Operations** (4 tools): Insert and delete rows and columns.
- **Formatting** (6 tools): Font, fill, alignment, borders, number formats, merge/unmerge, auto-fit.
- **Conditional Formatting** (3 tools): Color scales, data bars, icon sets, highlight rules.
- **Formulas** (4 tools): Single, array, and batch formula operations; syntax validation.
- **Tables** (2 tools): Create and list native Excel tables (ListObjects).
- **Data Validation** (4 tools): Dropdown lists, numeric constraints, formula-based validation.
- **Protection** (3 tools): Protect/unprotect sheets, lock/unlock cells.
- **Charts** (3 tools): Create (6 types), list, and delete native Excel charts.
- **Data Analysis** (7 tools): Filter, sort, aggregate, statistics, find duplicates, profile data, search/replace.
- **CSV Operations** (3 tools): CSV preview, CSV↔XLSX conversion.
- **Pivot & ETL** (5 tools): Pivot tables, unpivot, merge datasets, computed columns, deduplication.
- **Financial** (5 tools): NPV, IRR, PMT, goal seek (with AST-safe expressions), loan amortization.
