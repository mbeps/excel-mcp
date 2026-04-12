"""TypedDict for representing comments attached to cells."""

from __future__ import annotations

from typing import TypedDict


class CommentInfo(TypedDict):
    """A comment attached to a cell.

    Keys:
        cell_ref (str): Cell reference in A1 notation. Required.
        text (str): Comment text. Required.
        author (str): Author name. Required.
    """

    cell_ref: str
    text: str
    author: str
