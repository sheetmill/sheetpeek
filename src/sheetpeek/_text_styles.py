"""Text styles in ``excel`` mode (``ARCHITECTURE.md`` §4.13).

Superscript and subscript, the underlines, strikethrough, shrink to fit and the text's direction.

Every number here is from the screen suite ``text_style`` (Calibri 11 unless stated), the rules'
positions and thicknesses from xlformat's ``rule_metrics``.
"""

from __future__ import annotations

import functools
import unicodedata
from typing import Literal

import xlformat

from sheetpeek._model import Font
from sheetpeek._rows import script_size
from sheetpeek._text_model import TextRule

ACCOUNTING_INSETS = (3.0, 2.0)
UNDERLINE_OVERRUN = 1.0
"""Points a single or double underline runs past its text's end (Excel 16.115, screen suite
``text_style``, Calibri 11; accounting underlines and strikethrough end with the text)."""
"""An accounting underline under text spans the cell, less these from its left and right
gridlines (from the left gridline, and before the next one)."""

# Where xlformat has no rule positions (Courier New, sizes past 72 points): the offsets of the
# first line's top from the baseline at Calibri 11 (positive is down), scaled with the drawn
# size, 1 point thick (unverified).
_CAPTURED_SIZE = 11.0
_UNDERLINE_OFFSET = {
    "single": 1.0,
    "double": 0.0,
    "singleAccounting": 3.0,
    "doubleAccounting": 3.0,
}
_STRIKE_OFFSET = -3.0
_RULE_FACES = frozenset(("calibri", "aptos narrow", "arial", "times new roman", "georgia"))
"""The faces xlformat's ``rule_metrics`` covers; another face but Courier New is drawn as
Calibri, as it is measured."""


SHRINK_STEPS = 7
"""Shrink to fit draws the largest ``k / 7`` of the size, ``k`` from 1 to 7, that fits (captured
at Calibri 11 at zoom 100; the step at other sizes is unverified)."""


def vert_scale(font: Font) -> float:
    """The drawn size over the font's size for superscript and subscript, else 1.

    The drawn size is ``round(2s/3)`` of the whole size ``s`` (``_rows.script_size``).
    """
    if font.vert_align is None or font.size <= 0:
        return 1.0
    return script_size(font)[1] / font.size


def vert_shift(font: Font) -> float:
    """Points the baseline moves down (positive) or up for superscript and subscript.

    Raised ``floor(s/2)`` and lowered ``floor(s/5)`` from the row's baseline, ``s`` the size
    rounded up (captured, ``row_bottoms`` groups ``scripts_1`` to ``scripts_11``, Calibri 6 to
    40; other faces unverified).
    """
    if font.vert_align == "superscript":
        return -float(script_size(font)[2])
    if font.vert_align == "subscript":
        return float(script_size(font)[3])
    return 0.0


def accounting(font: Font) -> bool:
    """Whether the font has an accounting underline."""
    return font.underline in ("singleAccounting", "doubleAccounting")


def scaled(font: xlformat.Font, scale: float) -> xlformat.Font:
    """An xlformat font at ``scale`` times its size."""
    if scale == 1.0:
        return font
    return xlformat.Font(font.name, font.size * scale, font.bold, font.italic)


def shrink_step(text: str, font: xlformat.Font, budget: float) -> int:
    """The step ``k`` of shrink to fit: the largest from 1 to 7 at which the text fits ``budget``.

    7 when it fits unshrunk; 1 when nothing fits, the text then clipped (decided, §4.13).
    """
    for k in range(SHRINK_STEPS, 0, -1):
        if xlformat.text_width(text, font=scaled(font, k / SHRINK_STEPS)) <= budget:
            return k
    return 1


def rules(
    font: Font,
    size: float,
    baseline: float,
    span: tuple[float, float],
    cell: tuple[float, float] | None,
) -> tuple[TextRule, ...]:
    """The underline and strikethrough of one run.

    Each line's top and thickness come from xlformat's ``rule_metrics`` for the face at the
    drawn size, at zoom 100 (captured in xlformat's suites; here, ``text_style_more`` group
    ``rules`` at 9, 11, 14, 20 and 36 points): a single underline, a double one (both lines
    the double's thickness, the second its gap below), an accounting one (its first
    line at the accounting top) and a strikethrough above the baseline.

    Args:
        font: The cell's font.
        size: The drawn size, in points.
        baseline: The run's baseline.
        span: Where the run's text starts and ends.
        cell: The left and next gridline of the cell, for an accounting underline under text,
            which spans the cell; ``None`` for a number, whose underline spans the number.
    """
    found: list[TextRule] = []
    kind = font.underline
    metrics = _rule_metrics(font.name, size, font.bold, font.italic)
    if kind in _UNDERLINE_OFFSET:
        x0, x1 = span
        if cell is not None and kind in ("singleAccounting", "doubleAccounting"):
            x0, x1 = cell[0] + ACCOUNTING_INSETS[0], cell[1] - ACCOUNTING_INSETS[1]
        elif kind in ("single", "double") and x1 > x0:
            x1 += UNDERLINE_OVERRUN
        lines: Literal[1, 2] = 2 if kind in ("double", "doubleAccounting") else 1
        gap = 2.0 if metrics is None else metrics.double_gap
        if metrics is None:
            top, thickness = float(round(_UNDERLINE_OFFSET[kind] * size / _CAPTURED_SIZE)), 1.0
        elif kind == "single":
            top, thickness = metrics.underline_top, metrics.underline_thickness
        elif kind == "double":
            top, thickness = metrics.double_top, metrics.double_thickness
        elif kind == "singleAccounting":
            top, thickness = metrics.accounting_top, metrics.underline_thickness
        else:
            top, thickness = metrics.accounting_top, metrics.double_thickness
        if x1 > x0:
            found.append(TextRule("underline", x0, x1, baseline + top, lines, thickness, gap))
    if font.strike and span[1] > span[0]:
        if metrics is None:
            top, thickness = float(round(_STRIKE_OFFSET * size / _CAPTURED_SIZE)), 1.0
        else:
            top, thickness = -metrics.strikethrough_top, metrics.strikethrough_thickness
        found.append(TextRule("strike", span[0], span[1], baseline + top, 1, thickness))
    return tuple(found)


@functools.lru_cache(maxsize=256)
def _rule_metrics(name: str, size: float, bold: bool, italic: bool) -> xlformat.RuleMetrics | None:
    """Xlformat's rules for the face at ``size``, zoom 100; ``None`` where it has none."""
    face = name if name.lower() in _RULE_FACES else "Calibri"
    if name.lower() == "courier new":  # its rules raise in xlformat 0.6
        return None
    try:
        return xlformat.rule_metrics(xlformat.Font(face, size, bold, italic))
    except LookupError:
        return None


def right_to_left(text: str) -> bool:
    """Whether the text's first strong character is right-to-left (Hebrew, Arabic and others).

    Text with no strong character is left to right.
    """
    if text.isascii():  # no ASCII character is strongly right-to-left
        return False
    for char in text:
        kind = unicodedata.bidirectional(char)
        if kind == "L":
            return False
        if kind in ("R", "AL"):
            return True
    return False
