"""Row heights in points, as Excel for Mac lays them out (``ARCHITECTURE.md`` §4.7).

A row's own ``ht`` wins. Otherwise ``<sheetFormatPr>`` decides, together with whether the file
has a ``<row>`` element for the row: Excel ignores ``defaultRowHeight`` unless ``customHeight``
is set, and gives the row the Normal font's height from xlformat's row model.
"""

from __future__ import annotations

import math
from dataclasses import replace
from functools import lru_cache

import xlformat

from sheetpeek._model import Font, Row, Sheet

_LARGEST = 72.0
"""The largest size xlformat's row model covers."""
TALLEST_ROW = 409.5
"""Excel's tallest row, in points: a row fitted to a font or to 90-degree text stops there."""
NO_SHEET_FORMAT = 16.0  # an unlisted row of a file without <sheetFormatPr>, every font
WRAPPED_LEADING = 1.0  # wrapped lines are the font's row height plus 1 apart (Calibri 11: 16)
ACCOUNTING_FALLBACK = 3.0
"""Points an accounting underline adds where xlformat has no accounting rows (Courier New below
5.5 points): Calibri 11's 3 (unverified)."""

_ROW_FACES = frozenset(
    ("calibri", "aptos narrow", "arial", "times new roman", "courier new", "georgia")
)
"""The faces xlformat's row model covers, at every size from 1 to 72 points."""


def _row_font(font: Font, underline: str | None = None) -> xlformat.Font:
    """The font xlformat's row model reads: Calibri for a face it lacks, the size bounded."""
    name = font.name if font.name.lower() in _ROW_FACES else "Calibri"
    size = min(max(font.size, 1.0), _LARGEST)  # past 72: see face_metrics
    return xlformat.Font(name, size, font.bold, font.italic, underline)


@lru_cache(maxsize=1024)
def face_metrics(font: Font) -> tuple[float, float, float, float]:
    """(height ``F``, ascent ``ceil(S) + 1``, bottom offset ``B``, accounting addition ``A``).

    From xlformat's row model (``row_metrics``): a face's height ``F`` at its exact size, its
    ascent, the top-aligned baseline's distance below a row's top, and ``B = F - ceil(S) - 1``,
    the bottom-aligned baseline's distance above a row's bottom in a row of this font alone.
    ``A`` is what the font's accounting underline adds to ``B`` and to each wrapped line, 0
    without one (xlformat's ``RowMetrics.accounting_height``). A face xlformat has no rows for
    takes Calibri's at the size (a known deviation for a face other than xlformat's six).
    """
    plain = xlformat.row_metrics(_row_font(font)).font_height
    ascent = float(math.ceil(min(max(font.size, 1.0), _LARGEST))) + 1.0
    if font.size > _LARGEST:  # 72's scaled with the size (unverified; 409 caps at TALLEST_ROW)
        scale = math.ceil(font.size) / _LARGEST
        plain, ascent = float(math.floor(plain * scale)), float(math.ceil(font.size)) + 1.0
    extra = 0.0
    if font.underline in ("singleAccounting", "doubleAccounting"):
        try:
            extra = xlformat.row_metrics(_row_font(font, font.underline)).accounting_height
        except LookupError:  # Courier New below 5.5 points
            extra = ACCOUNTING_FALLBACK
    return plain, ascent, plain - ascent, extra


def font_row_height(font: Font) -> float:
    """The height of a row in ``font`` alone: xlformat's row model at the font's size (§4.7).

    A family xlformat has no rows for uses Calibri's heights at the same size, as xlformat
    measures it as Calibri. Past 72 points, 72's height scaled with the size (unverified), at
    most Excel's tallest row (captured, ``package_more``: a 409 pt Normal font gives 409.5).
    """
    return min(face_metrics(font)[0], TALLEST_ROW)


def row_height(row: Row, sheet: Sheet) -> float:
    """The height of a row in points; 0 for a row Excel hides by its sheet format."""
    if row.height is not None:
        return row.height
    if sheet.zero_row_height and not row.listed:
        return 0.0
    if sheet.custom_row_height and not sheet.zero_row_height and sheet.default_row_height:
        return sheet.default_row_height
    if not sheet.has_sheet_format and not row.listed:
        return NO_SHEET_FORMAT
    height = font_row_height(sheet.normal_font)
    if not row.listed:
        height += sheet.thick_top + sheet.thick_bottom
    return height


def fits_cells(row: Row, sheet: Sheet) -> bool:
    """Whether Excel fits this row to its cells when it opens the file.

    Captured (suite ``row_height``): a row with a ``<row>`` element and no ``ht`` grows to its
    largest cell font, or to its wrapped lines, and never below the Normal font's height; in a
    horizontal merge the font counts and the wrapped lines do not, and a merge across rows counts
    in no row (suite ``merge_rows``). A stored ``ht`` and ``customHeight`` defaults are kept.
    """
    return row.height is None and row.listed and not sheet.custom_row_height


def wrapped_height(lines: int, font: Font) -> float:
    """The height Excel gives a row for ``lines`` wrapped lines in ``font``.

    Captured: three and six lines of Calibri 11 are 48 and 96 pt (``row_height``); three lines in
    other fonts and sizes in ``row_height_more``. An accounting underline adds its ``A`` to each
    line (captured, ``text_style_more`` group ``rules``: three lines of Calibri 11 take 57 pt).
    """
    height, _, _, extra = face_metrics(font)
    if lines <= 1:
        return height + extra
    return lines * (height + WRAPPED_LEADING + extra)


def baseline_offsets(font: Font) -> tuple[float, float, float]:
    """(line pitch, baseline below a top-aligned line's top, baseline above a bottom one's bottom).

    Excel's screen for a line in ``font`` alone (§3.6): the face's height ``F`` apart, ``A``
    more with an accounting underline; the top-aligned baseline ``ceil(S) + 1`` below the top
    (captured, ``text_style_more`` group ``faces``); the bottom-aligned one ``B``, and ``A``
    more, above the bottom, where a row of several fonts takes the largest (``row_bottom``).
    """
    height, ascent, bottom, extra = face_metrics(font)
    return height + extra, ascent, bottom + extra


def line_inset(font: Font) -> float:
    """A face's line inset ``B - 2``, xlformat's ``VerticalMetrics.line_inset``.

    Wrapped text in a face with a smaller one than the row's unwrapped fonts sits the
    difference higher.
    """
    return face_metrics(font)[2] - 2.0


def print_baseline_offsets(font: Font) -> tuple[float, float, float]:
    """``baseline_offsets`` as Excel prints, kept for a print or PDF output.

    From the PDF export at Calibri 11 (``layout`` rows 1 and 152 to 154): a 15 pt line with its
    baseline 11 pt below the top of a top-aligned block and 3 pt above the bottom of a
    bottom-aligned one, wrapped lines 16 pt apart; other fonts scale the 11 with their line
    height. The print is not the screen (rows 15 pt in the file print 14 pt apart).
    """
    line = font_row_height(font)
    above = round(line * 11 / 15)
    return line, float(above), line - above - 1


# The part of a Calibri superscript's row above the shared baseline, by its whole size: its
# fitted row less its bottom (captured, row_height_fitted group styled, every size from 10 to
# 40; below 10 the Normal font's 12 is the larger). No closed form fits every size: the raise
# plus the drawn size's ascent gives it from 16 to 31 and a point less at 10 to 15 and 32 to 40.
_SUPERSCRIPT_ABOVE = {
    10: 13, 11: 14, 12: 16, 13: 17, 14: 18, 15: 19, 16: 20, 17: 20, 18: 22, 19: 23, 20: 24,
    21: 25, 22: 27, 23: 27, 24: 29, 25: 30, 26: 31, 27: 32, 28: 34, 29: 34, 30: 36, 31: 37,
    32: 39, 33: 40, 34: 42, 35: 42, 36: 44, 37: 45, 38: 46, 39: 47, 40: 49,
}  # fmt: skip


def script_size(font: Font) -> tuple[int, int, int, int]:
    """(whole size ``s``, drawn size ``d``, raise, lowering) of a script font, in points.

    A superscript or subscript cell acts at its size rounded up, ``s = ceil(S)``, and draws at
    ``d = round(2s/3)``, raised ``floor(s/2)`` or lowered ``floor(s/5)`` (captured,
    ``row_bottoms`` groups ``scripts_1`` to ``scripts_11`` and ``row_height_fitted`` group
    ``scripts``: Calibri at every whole and half size from 6 to 40).
    """
    s = math.ceil(min(max(font.size, 1.0), 409.0))
    return s, math.floor(2 * s / 3 + 0.5), s // 2, s // 5


@lru_cache(maxsize=256)
def script_parts(font: Font) -> tuple[float, float, float]:
    """(part above the shared baseline, part below it, line inset) of a script font's row.

    For a superscript or subscript font (§4.13). A superscript's below is its whole size's
    ``B``, its above the Calibri table above; for another face or size, the raise plus the
    drawn size's ascent (unverified). A subscript's line grows by ``g = l + B(d) - B(s)`` on
    both sides: ``ceil(s) + 1 + g`` above and ``l + B(d)`` below, which is also the row's
    shared bottom (captured, ``row_bottoms`` and ``row_height_fitted``; at 6 to 8 pt Excel's
    subscript bottom is a point less).
    """
    s, d, r, lower = script_size(font)
    whole = replace(font, size=float(s), vert_align=None)
    _, ascent, bottom, extra = face_metrics(whole)
    drawn = face_metrics(replace(font, size=float(d), vert_align=None, underline=None))
    if font.vert_align == "superscript":
        table = _SUPERSCRIPT_ABOVE if font.name.lower() == "calibri" else {}
        above = float(table.get(s, r + drawn[1]))
        return above, bottom + extra, bottom - 2.0
    below = lower + drawn[2]
    return ascent + below - bottom, below + extra, bottom - 2.0
