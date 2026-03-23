from __future__ import annotations

import csv
from pathlib import Path

from mcp_server.tools.csv_ops import csv_to_xlsx, read_csv_preview, xlsx_to_csv


def test_read_csv_preview(sample_csv: str) -> None:
    result = read_csv_preview(sample_csv, rows=3)
    assert result["total_rows"] == 5
    assert len(result["rows"]) == 3
    assert result["headers"] == ["Name", "Age", "City", "Salary"]


def test_csv_to_xlsx(sample_csv: str, tmp_path: Path) -> None:
    xlsx_path = str(tmp_path / "converted.xlsx")
    csv_to_xlsx(sample_csv, xlsx_path, sheet_name="Imported")

    from mcp_server.tools.workbook import list_sheets

    sheets = list_sheets(xlsx_path)
    assert sheets[0]["name"] == "Imported"


def test_xlsx_to_csv(sample_xlsx: str, tmp_path: Path) -> None:
    csv_path = str(tmp_path / "exported.csv")
    xlsx_to_csv(sample_xlsx, "Sheet1", csv_path)

    with open(csv_path) as f:
        reader = csv.reader(f)
        rows = list(reader)
    assert rows[0] == ["Name", "Age", "City", "Salary"]
    assert len(rows) == 6  # header + 5 data rows
