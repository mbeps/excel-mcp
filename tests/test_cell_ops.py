from __future__ import annotations

from mcp_server.tools.cell_ops import (
    clear_range,
    read_cell,
    read_file_chunked,
    read_range,
    read_ranges_batch,
    write_cell,
    write_range,
)


def test_read_cell(sample_xlsx: str) -> None:
    result = read_cell(sample_xlsx, "Sheet1", "A1")
    assert result["value"] == "Name"
    assert result["cell_ref"] == "A1"


def test_write_cell(sample_xlsx: str) -> None:
    write_cell(sample_xlsx, "Sheet1", "E1", "Status")
    result = read_cell(sample_xlsx, "Sheet1", "E1")
    assert result["value"] == "Status"


def test_read_range(sample_xlsx: str) -> None:
    result = read_range(sample_xlsx, "Sheet1", "A1", "D2")
    assert result["row_count"] == 2
    assert result["col_count"] == 4
    assert result["rows"][0] == ["Name", "Age", "City", "Salary"]
    assert result["rows"][1][0] == "Alice"


def test_write_range(sample_xlsx: str) -> None:
    data = [["X", "Y"], [1, 2], [3, 4]]
    write_range(sample_xlsx, "Sheet1", "F1", data)
    result = read_range(sample_xlsx, "Sheet1", "F1", "G3")
    assert result["rows"][0] == ["X", "Y"]
    assert result["rows"][1] == [1, 2]
    assert result["rows"][2] == [3, 4]


def test_clear_range(sample_xlsx: str) -> None:
    clear_range(sample_xlsx, "Sheet1", "A2", "D2")
    result = read_range(sample_xlsx, "Sheet1", "A2", "D2")
    assert all(v is None for v in result["rows"][0])


def test_read_file_chunked(sample_xlsx: str) -> None:
    result = read_file_chunked(sample_xlsx, "Sheet1", start_row=0, chunk_size=3)
    assert result["chunk_size"] == 3
    assert result["chunk_start"] == 0
    assert result["has_more"] is True
    assert len(result["rows"]) == 3


def test_copy_range(sample_xlsx: str) -> None:
    """Test copying a range within same sheet."""
    from mcp_server.tools.cell_ops import copy_range

    result = copy_range(sample_xlsx, "Sheet1", "A1:B2", "Sheet1", "D1")
    assert "Copied" in result or "copied" in result


def test_delete_range(sample_xlsx: str) -> None:
    """Test deleting a range with shift up."""
    from mcp_server.tools.cell_ops import delete_range

    result = delete_range(sample_xlsx, "Sheet1", "A1:B1", "up")
    assert "Deleted" in result or "deleted" in result


def test_get_file_info(sample_xlsx: str) -> None:
    """Test getting file info."""
    from mcp_server.tools.cell_ops import get_file_info

    result = get_file_info(sample_xlsx)
    assert "size_bytes" in result
    assert "sheets" in result
    assert result["sheet_count"] == 1
    assert isinstance(result, dict)


def test_read_range_show_formula(sample_xlsx: str) -> None:
    """Test reading range with formulas."""
    from mcp_server.tools.formulas import set_formula

    set_formula(sample_xlsx, "Sheet1", "D1", "=SUM(A1:C1)")
    result = read_range(sample_xlsx, "Sheet1", "D1", "D1", show_formula=True)
    data = result.get("data", result.get("rows", []))
    assert len(data) > 0


def test_read_range_show_style(sample_xlsx: str) -> None:
    """Test reading range with style info."""
    result = read_range(sample_xlsx, "Sheet1", "A1", "B2", show_style=True)
    assert "styles" in result


def test_read_range_html_format(sample_xlsx: str) -> None:
    """Test reading range as HTML table."""
    result = read_range(sample_xlsx, "Sheet1", "A1", "B2", output_format="html")
    html = result.get("html", result.get("data", ""))
    assert "<table" in str(html).lower() or "<tr" in str(html).lower()


def test_validate_excel_range() -> None:
    """Test Excel range validation."""
    from mcp_server.utils.excel_helpers import validate_excel_range

    result = validate_excel_range("A1:C10")
    assert result["valid"] is True

    result = validate_excel_range("B5")
    assert result["valid"] is True

    result = validate_excel_range("ZZZ1")
    assert result["valid"] is False


def test_read_ranges_batch(sample_xlsx: str) -> None:
    """Test reading multiple non-contiguous ranges in a single workbook load."""
    result = read_ranges_batch(sample_xlsx, "Sheet1", ["A1:B2", "C1:D1"])
    assert "results" in result
    assert "A1:B2" in result["results"]
    assert "C1:D1" in result["results"]
    # A1:B2 should have 2 rows, 2 cols
    a1b2 = result["results"]["A1:B2"]
    assert len(a1b2) == 2
    assert len(a1b2[0]) == 2
    assert a1b2[0][0] == "Name"
    assert a1b2[0][1] == "Age"
    # C1:D1 should have 1 row, 2 cols
    c1d1 = result["results"]["C1:D1"]
    assert len(c1d1) == 1
    assert c1d1[0][0] == "City"
    assert c1d1[0][1] == "Salary"
