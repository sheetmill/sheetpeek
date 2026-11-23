"""Text styles in ``excel`` mode (``ARCHITECTURE.md`` §4.13).

Superscript and subscript, the underlines, strikethrough, shrink to fit and the text's direction.

Every number here is from the screen suite ``text_style`` (Calibri 11 unless stated), the rules'
positions and thicknesses from xlformat's ``rule_metrics``.
"""

from __future__ import annotations

import functools
import unicodedata
from collections.abc import Iterable
from typing import Literal

import xlformat

from sheetpeek._model import Font
from sheetpeek._text_model import TextRule

VERT_SCALE = 2 / 3
"""Superscript and subscript are drawn at about two thirds of the size (cap heights 0.64 to
0.68 of the plain text's at 11 and 24 pt)."""
_RAISE = 0.48  # em of the font's size: 5 pt at 11 pt, 12 pt at 24 pt
_LOWER = 0.18  # 2 pt at 11 pt, 4 pt at 24 pt

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
    """The drawn size over the font's size for superscript and subscript, else 1."""
    return VERT_SCALE if font.vert_align is not None else 1.0


def vert_shift(font: Font) -> float:
    """Points the baseline moves down (positive) or up for superscript and subscript.

    Captured at 11 and 24 pt: raised 5 and 12 pt, lowered 2 and 4 pt, about 0.48 and 0.18 em.
    """
    if font.vert_align == "superscript":
        return -float(round(_RAISE * font.size))
    if font.vert_align == "subscript":
        return float(round(_LOWER * font.size))
    return 0.0


def vert_extra(font: Font) -> float:
    """Points superscript or subscript add to a fitted row.

    Captured: 2 at 11 pt (17 pt rows for 15), 6 at 24 pt (37 for 31); linear between and beyond,
    never below 0 (unverified).
    """
    if font.vert_align is None:
        return 0.0
    return float(max(0, round(2 + (font.size - 11) * 4 / 13)))


def row_extras(fonts: Iterable[Font]) -> list[tuple[Font, float]]:
    """Each font of a row's cells with values, with what its text styles add to its row height.

    A fitted row takes the largest font row height plus addition (an accounting underline's is
    in the row's bottom, ``_excel._fitted``). A row holding both superscript and subscript grows
    1 pt more at 11 pt (18 pt, captured), half the one-kind addition at other sizes: from 12 pt
    Excel's rows differ, a known deviation (row_height_fitted, group styled).
    """
    fonts = list(fonts)
    both = {f.vert_align for f in fonts} >= {"superscript", "subscript"}
    result = []
    for font in fonts:
        extra = vert_extra(font)
        if both and extra:
            extra += round(extra / 2)
        result.append((font, extra))
    return result


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
    the double's thickness, the second two thicknesses below), an accounting one (its first
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
            found.append(TextRule("underline", x0, x1, baseline + top, lines, thickness))
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
