from __future__ import annotations

import csv
from pathlib import Path

from mcp_server.tools.csv_ops import csv_to_xlsx, detect_csv_dialect, read_csv_preview, validate_csv, xlsx_to_csv


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


def test_read_csv_preview_with_encoding(sample_csv: str) -> None:
    result = read_csv_preview(sample_csv, rows=5, encoding="utf-8")
    assert result["total_rows"] == 5
    assert result["headers"] == ["Name", "Age", "City", "Salary"]


def test_csv_to_xlsx_with_encoding(sample_csv: str, tmp_path: Path) -> None:
    xlsx_path = str(tmp_path / "enc_converted.xlsx")
    csv_to_xlsx(sample_csv, xlsx_path, sheet_name="Data", encoding="utf-8")
    from mcp_server.tools.workbook import list_sheets

    sheets = list_sheets(xlsx_path)
    assert sheets[0]["name"] == "Data"


def test_xlsx_to_csv_with_encoding(sample_xlsx: str, tmp_path: Path) -> None:
    csv_path = str(tmp_path / "enc_exported.csv")
    xlsx_to_csv(sample_xlsx, "Sheet1", csv_path, encoding="utf-8")
    with open(csv_path, encoding="utf-8") as f:
        reader = csv.reader(f)
        rows = list(reader)
    assert rows[0] == ["Name", "Age", "City", "Salary"]


def test_detect_csv_dialect(tmp_path: Path) -> None:
    """Test auto-detection of CSV dialect with semicolon delimiter."""
    csv_path = str(tmp_path / "semi.csv")
    with open(csv_path, "w", encoding="utf-8") as f:
        f.write("Name;Age;City\nAlice;30;NYC\nBob;25;LA\n")
    result = detect_csv_dialect(csv_path)
    assert result["delimiter"] == ";"
    assert "encoding" in result
    assert "has_header" in result
    assert "quotechar" in result


def test_validate_csv_valid(sample_csv: str) -> None:
    """Test CSV validation with matching expected columns."""
    result = validate_csv(sample_csv, expected_columns=["Name", "Age", "City", "Salary"])
    assert result["valid"] is True
    assert result["missing_columns"] == []
    assert result["column_count"] == 4
    assert result["row_count"] == 5


def test_validate_csv_missing_column(sample_csv: str) -> None:
    """Test CSV validation when expected columns include a missing one."""
    result = validate_csv(sample_csv, expected_columns=["Name", "Age", "City", "Salary", "Email"])
    assert result["valid"] is False
    assert "Email" in result["missing_columns"]
