"""Tests for images.py — image insertion operations."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook
from PIL import Image as PILImage

from mcp_server.tools.images import insert_image


@pytest.fixture()
def sample_image(tmp_path: Path) -> Path:
    """Create a small test image."""
    img_path = tmp_path / "test_image.png"
    img = PILImage.new("RGB", (100, 100), color="red")
    img.save(str(img_path))
    return img_path


@pytest.fixture()
def sample_xlsx(tmp_path: Path) -> str:
    """Create a sample xlsx for image tests."""
    fp = tmp_path / "test.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Test"
    wb.save(str(fp))
    wb.close()
    return str(fp)


def test_insert_image(sample_xlsx: str, sample_image: Path) -> None:
    result = insert_image(sample_xlsx, "Sheet1", str(sample_image), "B2")
    assert result["status"] == "success"


def test_insert_image_with_dimensions(sample_xlsx: str, sample_image: Path) -> None:
    result = insert_image(sample_xlsx, "Sheet1", str(sample_image), "C3", width=200, height=150)
    assert result["status"] == "success"


def test_insert_image_invalid_extension(sample_xlsx: str, tmp_path: Path) -> None:
    bad_file = tmp_path / "test.txt"
    bad_file.write_text("not an image")
    with pytest.raises(ValueError):
        insert_image(sample_xlsx, "Sheet1", str(bad_file), "A1")


def test_insert_image_missing_file(sample_xlsx: str) -> None:
    with pytest.raises((ValueError, FileNotFoundError)):
        insert_image(sample_xlsx, "Sheet1", "/nonexistent/image.png", "A1")
