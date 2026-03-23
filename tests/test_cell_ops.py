from __future__ import annotations

from mcp_server.tools.cell_ops import (
    clear_range,
    read_cell,
    read_file_chunked,
    read_range,
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
