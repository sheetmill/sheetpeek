"""Text styles in ``excel`` mode (``ARCHITECTURE.md`` §4.13).

Superscript and subscript, the underlines, strikethrough, shrink to fit and the text's direction.

Every number here is from the screen suite ``text_style`` (Calibri 11 unless stated), the rules'
positions and thicknesses from xlformat's ``rule_metrics``.
"""

from __future__ import annotations

import functools
import math
import unicodedata
from typing import Literal

import xlformat

from sheetpeek import _platform
from sheetpeek._model import Font
from sheetpeek._rows import script_size
from sheetpeek._text_model import TextRule

ACCOUNTING_INSETS = (3.0, 2.0)
"""An accounting underline under turned text spans the text less these from its left and right
(unverified). Under upright text it spans the cell less its font's insets (§3.6)."""
UNDERLINE_OVERRUN = 1.0
"""Points a single or double underline runs past its text's end (Excel 16.115, screen suite
``text_style``, Calibri 11; accounting underlines and strikethrough end with the text)."""

# Where xlformat has no rule positions (sizes past 72 points, Courier New below 6): the offsets
# of the first line's top from the baseline at Calibri 11 (positive is down), scaled with the
# drawn size, 1 point thick (unverified).
_CAPTURED_SIZE = 11.0
_UNDERLINE_OFFSET = {
    "single": 1.0,
    "double": 0.0,
    "singleAccounting": 3.0,
    "doubleAccounting": 3.0,
}
_STRIKE_OFFSET = -3.0


SHRINK_PADDING = 0.375
"""Advances a shrunk text's fit test charges beyond the text, at the level tried (fitted to the
``text_steps`` capture, §4.13)."""
SHRINK_MARGIN = 1.0
"""Units of the cell a shrunk text's fit test leaves free (fitted with ``SHRINK_PADDING``): points
on the Mac, device pixels on Windows."""


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


def shrink_levels(
    advance: float, platform: str = _platform.DEFAULT_PLATFORM_NAME
) -> tuple[float, ...]:
    """The "H" advances, in points, shrink to fit may draw a text at, largest first.

    Excel for Mac draws every whole point below the unshrunk ``advance`` down to 1; Excel for
    Windows every whole device pixel below it except 6, down to 2 (captured, ``text_steps`` and
    ``windows/text_steps``, §4.13).
    """
    if platform == "windows":
        unit = 0.75
        steps = [n for n in range(math.ceil(advance / unit - 1e-9) - 1, 1, -1) if n != 6]
    else:
        unit = 1.0
        steps = list(range(math.ceil(advance - 1e-9) - 1, 0, -1))
    return tuple(n * unit for n in steps)


def shrink_scale(
    width: float,
    advance: float,
    cell: float,
    budget: float,
    platform: str = _platform.DEFAULT_PLATFORM_NAME,
) -> float:
    """The scale shrink to fit draws a text at: 1 when it fits, else a level's ``a / advance``.

    Excel draws shrunk text at a whole-unit "H" advance ``a`` (:func:`shrink_levels`: points on
    the Mac, device pixels on Windows). Text that fits ``budget``, the cell less its insets, is
    not shrunk. Otherwise the level is the largest ``a`` at which the text's width scaled to it,
    plus ``SHRINK_PADDING`` advances, fits ``cell`` less ``SHRINK_MARGIN`` units: for a run of n
    "H"s, ``(n + 0.375)·a <= cell - 1`` on the Mac. The smallest level when nothing fits, the
    text then clipped (decided).

    Args:
        width: The text's unshrunk width, in points.
        advance: The unshrunk "H" advance of the cell's font, in points.
        cell: The cell's width less any indent, in points.
        budget: What text may take unshrunk: ``cell`` less the insets.
        platform: ``"mac"`` or ``"windows"``.
    """
    levels = shrink_levels(advance, platform)
    if width <= budget or not levels:
        return 1.0
    margin = SHRINK_MARGIN * (0.75 if platform == "windows" else 1.0)
    for level in levels:
        if (width / advance + SHRINK_PADDING) * level <= cell - margin:
            return level / advance
    return levels[-1] / advance


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
        cell: Where an accounting underline under text starts and ends, the cell less its
            insets; ``None`` for a number, whose underline spans the number.
    """
    found: list[TextRule] = []
    kind = font.underline
    metrics = _rule_metrics(font.name, size, font.bold, font.italic)
    if kind in _UNDERLINE_OFFSET:
        x0, x1 = span
        if cell is not None and kind in ("singleAccounting", "doubleAccounting"):
            x0, x1 = cell
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


def _rule_metrics(name: str, size: float, bold: bool, italic: bool) -> xlformat.RuleMetrics | None:
    """Xlformat's rules for the face at ``size``, zoom 100, on the platform measured.

    ``None`` where it has none.
    """
    return _platform_rules(name, size, bold, italic, _platform.current())


@functools.lru_cache(maxsize=256)
def _platform_rules(
    name: str, size: float, bold: bool, italic: bool, platform: xlformat.Platform
) -> xlformat.RuleMetrics | None:
    """``_rule_metrics`` on ``platform``.

    A face xlformat has no rules for takes Calibri's at the size, as the layout measures it
    (reported ``font.no_metrics``; unverified).
    """
    face = _platform.metrics_face(name, "rule_metrics", platform)
    try:
        return xlformat.rule_metrics(xlformat.Font(face, size, bold, italic), platform=platform)
    except LookupError:  # a size xlformat has no rules for
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
