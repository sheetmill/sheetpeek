"""The lines drawn with text in ``excel`` mode: underlines and strikethrough (§4.13)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, TypeAlias

TextRuleKind: TypeAlias = Literal["underline", "strike"]
TextDirection: TypeAlias = Literal["ltr", "rtl"]


@dataclass(frozen=True, slots=True)
class TextRule:
    """A horizontal line drawn with a text run: an underline or a strikethrough.

    Positions are in points, in sheet coordinates, as a :class:`~sheetpeek.TextRun`'s. The
    thickness is not: Excel draws these lines 1 screen point thick at any zoom, so a drawing
    output makes each one device-pixel thin line (one CSS pixel in HTML), and puts a second line
    2 screen points below the first.
    """

    kind: TextRuleKind
    x0: float
    """Where the line starts."""
    x1: float
    """Where the line ends."""
    y: float
    """The top of the (first) line."""
    lines: Literal[1, 2] = 1
    """1, or 2 for a double underline."""
