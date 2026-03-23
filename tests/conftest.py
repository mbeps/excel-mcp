from __future__ import annotations

import csv
from pathlib import Path

import pytest
from openpyxl import Workbook

HEADERS = ["Name", "Age", "City", "Salary"]
ROWS = [
    ["Alice", 30, "New York", 70000],
    ["Bob", 25, "Chicago", 55000],
    ["Charlie", 35, "New York", 90000],
    ["Diana", 28, "Chicago", 62000],
    ["Eve", 32, "Boston", 80000],
]


@pytest.fixture()
def sample_xlsx(tmp_path: Path) -> str:
    path = str(tmp_path / "sample.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(HEADERS)
    for row in ROWS:
        ws.append(row)
    wb.save(path)
    return path


@pytest.fixture()
def sample_csv(tmp_path: Path) -> str:
    path = str(tmp_path / "sample.csv")
    with open(path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS)
        writer.writerows(ROWS)
    return path


@pytest.fixture()
def empty_xlsx(tmp_path: Path) -> str:
    path = str(tmp_path / "empty.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    wb.save(path)
    return path
