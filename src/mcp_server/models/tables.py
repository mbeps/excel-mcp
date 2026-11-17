"""Schemas for table metadata."""

from __future__ import annotations

from typing import TypedDict


class TableInfo(TypedDict):
    """An Excel table defined on a worksheet.

    Keys:
        name (str): Table name. Required.
        ref (str): A1-style reference for the table. Required.
        style (str | None): Named table style, if set.
    """

    name: str
    ref: str
    style: str | None
