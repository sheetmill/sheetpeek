"""The lines drawn with text in ``excel`` mode: underlines and strikethrough (§4.13)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, TypeAlias

TextRuleKind: TypeAlias = Literal["underline", "strike"]
TextDirection: TypeAlias = Literal["ltr", "rtl"]


@dataclass(frozen=True, slots=True)
class TextRule:
    """A horizontal line drawn with a text run: an underline or a strikethrough.

    Positions are in points, in sheet coordinates, as a :class:`~sheetpeek.TextRun`'s, and so is
    ``thickness``, as Excel draws the line at zoom 100 (xlformat's ``rule_metrics``): 1 point up
    to about 15 points of Calibri, more at larger sizes. A double underline's second line lies
    two thicknesses below the first's top. Excel draws a rule at least 1 screen point thick at
    any zoom, so an output draws a 1 point rule one device-pixel thin line (one CSS pixel in
    HTML) and scales a thicker one with the zoom.
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
    thickness: float = 1.0
    """The thickness of each line, in points at zoom 100."""
