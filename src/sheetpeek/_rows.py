"""Row heights in points, as Excel lays them out on the platform measured (§4.7).

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
from sheetpeek._platform import current, metrics_face

_LARGEST = 72.0
"""The largest size xlformat's row model covers."""
TALLEST_ROW = 409.5
"""Excel's tallest row, in points: a row fitted to a font or to 90-degree text stops there."""
NO_SHEET_FORMAT = 16.0  # an unlisted row of a file without <sheetFormatPr>, every font
ACCOUNTING_FALLBACK = 3.0
"""Points an accounting underline adds where xlformat has no accounting rows (Courier New below
5.5 points): Calibri 11's 3 (unverified)."""


def _row_font(
    font: Font, platform: xlformat.Platform, underline: str | None = None
) -> xlformat.Font:
    """The font xlformat's row model reads on ``platform``.

    Calibri for a face it lacks (reported ``font.no_metrics`` by the layout), the size bounded
    to the model's 1 to 72 points.
    """
    function = "vertical_metrics" if platform.name == "windows" else "row_metrics"
    name = metrics_face(font.name, function, platform)
    size = min(max(font.size, 1.0), _LARGEST)  # past 72: see face_metrics
    return xlformat.Font(name, size, font.bold, font.italic, underline)


def wrapped_leading() -> float:
    """What each wrapped line adds beyond its face's height.

    A point on the Mac (Calibri 11 lines 16 pt apart in a fitted row), nothing on Windows
    (20 px; xlformat's ``line_height``).
    """
    return 0.0 if current().name == "windows" else 1.0


def unit() -> float:
    """The screen unit positions round to, in points: a point on the Mac, a pixel on Windows."""
    return 0.75 if current().name == "windows" else 1.0


def face_metrics(font: Font) -> tuple[float, float, float, float]:
    """(height ``F``, ascent, bottom offset ``B``, accounting addition ``A``), in points.

    On the Mac from xlformat's row model (``row_metrics``): a face's height ``F`` at its exact
    size, its ascent ``ceil(S) + 1``, the top-aligned baseline's distance below a row's top, and
    ``B = F - ceil(S) - 1``, the bottom-aligned baseline's distance above a row's bottom in a row
    of this font alone. On Windows from xlformat's ``vertical_metrics`` in whole pixels: the
    face's height ``H``, the ascent ``H - D + 1`` and ``B = D - 1``, ``D`` its descent. ``A`` is
    what the font's accounting underline adds to ``B`` and to each wrapped line, 0 without one
    (xlformat's ``RowMetrics.accounting_height``). A face xlformat has no rows for takes
    Calibri's at the size (a known deviation for a face other than xlformat's).
    """
    return _face(font, current())[:4]


def line_inset(font: Font) -> float:
    """A face's line inset, xlformat's ``VerticalMetrics.line_inset``: ``B - 2`` on the Mac.

    Wrapped text in a face with a smaller one than the row's unwrapped fonts sits the
    difference higher.
    """
    return _face(font, current())[4]


@lru_cache(maxsize=1024)
def _face(font: Font, platform: xlformat.Platform) -> tuple[float, float, float, float, float]:
    """``face_metrics`` and ``line_inset`` of ``font`` on ``platform``."""
    if platform.name == "windows":
        return _windows_face(font, platform)
    plain = xlformat.row_metrics(_row_font(font, platform), platform=platform).font_height
    ascent = float(math.ceil(min(max(font.size, 1.0), _LARGEST))) + 1.0
    if font.size > _LARGEST:  # 72's scaled with the size (unverified; 409 caps at TALLEST_ROW)
        scale = math.ceil(font.size) / _LARGEST
        plain, ascent = float(math.floor(plain * scale)), float(math.ceil(font.size)) + 1.0
    extra = 0.0
    if font.underline in ("singleAccounting", "doubleAccounting"):
        try:
            spec = _row_font(font, platform, font.underline)
            extra = xlformat.row_metrics(spec, platform=platform).accounting_height
        except LookupError:  # Courier New below 5.5 points
            extra = ACCOUNTING_FALLBACK
    return plain, ascent, plain - ascent, extra, plain - ascent - 2.0


def _windows_face(
    font: Font, platform: xlformat.Platform
) -> tuple[float, float, float, float, float]:
    """``_face`` on Excel for Windows: xlformat's ``vertical_metrics``, in whole pixels.

    A face xlformat has no Windows tables for takes Calibri's at the size, as it is measured
    (reported ``font.no_metrics``; unverified). A size past xlformat's Windows rows, which
    cover 1 to 72 points by ppem, takes the Mac's height and bottom offset rounded to whole
    pixels, its line inset the bottom offset, as at every tabled size (unverified above 72
    points). An accounting underline adds three of its
    thicknesses, the Mac's rule in pixels (captured at Calibri 11, ``windows/text_style``
    group ``fonts``: single and double rows of 23 px; other sizes unverified).
    """
    spec = _row_font(font, platform)
    plain = replace(font, underline=None)
    try:
        found = xlformat.vertical_metrics(spec, platform=platform)
    except LookupError:  # a size outside the Windows tables
        mac = _face(plain, xlformat.Platform.MAC)
        height = 0.75 * round(mac[0] / 0.75)
        bottom = 0.75 * round(mac[2] / 0.75)
        ascent, inset = height - bottom, bottom
    else:
        height, ascent = found.font_height, found.baseline_from_top
        bottom, inset = found.baseline_from_bottom, found.line_inset
    extra = 0.0
    if font.underline in ("singleAccounting", "doubleAccounting"):
        face = metrics_face(font.name, "rule_metrics", platform)
        rules = xlformat.rule_metrics(replace(spec, name=face), platform=platform)
        extra = 3 * rules.underline_thickness
    return height, ascent, bottom, extra, inset


def font_row_height(font: Font) -> float:
    """The height of a row in ``font`` alone: xlformat's row model at the font's size (§4.7).

    A family xlformat has no rows for uses Calibri's heights at the same size, as xlformat
    measures it as Calibri. Past 72 points, 72's height scaled with the size (unverified), at
    most Excel's tallest row (captured, ``package_more``: a 409 pt Normal font gives 409.5).
    """
    return min(face_metrics(font)[0], TALLEST_ROW)


STORED_PIXEL_ROUNDING = 1 / 8
"""Excel for Windows lays out a stored row height of ``s`` points at ``floor(4s/3 + c)`` pixels,
``c`` between 1/15 and 1/5 of a pixel (captured, ``windows/stored_rows``: 15.6 pt, 20.8 px, is
20 px and 15.7 pt, 20.93 px, is 21; 32 pt is 42 px). The exact ``c`` is unverified, bounded by
those rows; this is a value inside the bound."""


def _stored(height: float) -> float:
    """A stored row height as the platform lays it out: whole pixels on Windows (§4.7)."""
    if current().name != "windows" or height <= 0:
        return height
    return 0.75 * math.floor(height * 4 / 3 + STORED_PIXEL_ROUNDING)


def row_height(row: Row, sheet: Sheet) -> float:
    """The height of a row in points; 0 for a row Excel hides by its sheet format."""
    if row.height is not None:
        return _stored(row.height)
    if sheet.zero_row_height and not row.listed:
        return 0.0
    if (
        sheet.custom_row_height
        and not sheet.zero_row_height
        and sheet.default_row_height
        and not row.autofit
    ):
        return _stored(sheet.default_row_height)
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
    in no row (suites ``merge_rows`` and ``windows/merge_rows``, both platforms). A stored ``ht``
    with ``customHeight`` is kept, and so is a ``customHeight`` default, except for an
    ``autofit`` row: an ``ht`` without ``customHeight``, fitted under any default (suites
    ``refit_rows`` and ``windows/refit_rows``).
    """
    if row.height is not None or not row.listed:
        return False
    return row.autofit or not sheet.custom_row_height


def wrapped_height(lines: int, font: Font) -> float:
    """The height Excel gives a row for ``lines`` wrapped lines in ``font``.

    Captured: three and six lines of Calibri 11 are 48 and 96 pt (``row_height``); three lines in
    other fonts and sizes in ``row_height_more``. An accounting underline adds its ``A`` to each
    line (captured, ``text_style_more`` group ``rules``: three lines of Calibri 11 take 57 pt).
    """
    height, _, _, extra = face_metrics(font)
    if lines <= 1:
        return height + extra
    return lines * (height + wrapped_leading() + extra)


def baseline_offsets(font: Font) -> tuple[float, float, float]:
    """(line pitch, baseline below a top-aligned line's top, baseline above a bottom one's bottom).

    Excel's screen for a line in ``font`` alone (§3.6): the face's height ``F`` apart, ``A``
    more with an accounting underline; the top-aligned baseline ``ceil(S) + 1`` below the top
    (captured, ``text_style_more`` group ``faces``); the bottom-aligned one ``B``, and ``A``
    more, above the bottom, where a row of several fonts takes the largest (``row_bottom``).
    """
    height, ascent, bottom, extra = face_metrics(font)
    return height + extra, ascent, bottom + extra


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


def script_parts(font: Font) -> tuple[float, float, float]:
    """(part above the shared baseline, part below it, line inset) of a script font's row.

    For a superscript or subscript font (§4.13). A superscript's below is its whole size's
    ``B``, its above the Calibri table above; for another face or size, the raise plus the
    drawn size's ascent (unverified). A subscript's line grows by ``g = l + B(d) - B(s)`` on
    both sides: ``ceil(s) + 1 + g`` above and ``l + B(d)`` below, which is also the row's
    shared bottom (captured, ``row_bottoms`` and ``row_height_fitted``; at 6 to 8 pt Excel's
    subscript bottom is a point less).
    """
    return _script_parts(font, current())


# Excel for Windows: the part of a Calibri superscript's row above the shared baseline, in pixels,
# by ppem past _SUPERSCRIPT_ABOVE's sizes (captured, windows/row_height_fitted groups styled and
# scripts, 30.5 to 40 points), and a Calibri subscript's part below it at ppem 15 (Calibri 11 and
# 11.5: windows/row_height_fitted, and windows/text_style's "subscript alone"), where the rule
# gives 5 px.
_WINDOWS_SUPERSCRIPT_ABOVE = {
    41: 49, 42: 51, 43: 52, 44: 53, 45: 54, 46: 56, 47: 56, 48: 57, 49: 58, 50: 59, 51: 60,
    52: 62, 53: 63,
}  # fmt: skip
_WINDOWS_SUBSCRIPT_BELOW = {15: 6}


def ppem(size: float) -> int:
    """The pixels an em of a ``size``-point font takes on Excel for Windows, at 96 dpi."""
    return math.floor(size * 4 / 3 + 0.5)


def ppem_size(em: int) -> float:
    """The smallest whole or half point size whose ppem is ``em``."""
    return math.ceil((em - 0.5) * 1.5) / 2


@lru_cache(maxsize=256)
def _script_parts(font: Font, platform: xlformat.Platform) -> tuple[float, float, float]:
    """``script_parts`` on ``platform``."""
    if platform.name == "windows":
        return _windows_script_parts(font, platform)
    s, d, r, lower = script_size(font)
    whole = replace(font, size=float(s), vert_align=None)
    _, ascent, bottom, extra, inset = _face(whole, platform)
    drawn = _face(replace(font, size=float(d), vert_align=None, underline=None), platform)
    if font.vert_align == "superscript":
        mac_calibri = font.name.lower() == "calibri" and platform.name == "mac"
        table = _SUPERSCRIPT_ABOVE if mac_calibri else {}
        above = float(table.get(s, r + drawn[1]))
        return above, bottom + extra, inset
    below = lower + drawn[2]
    return ascent + below - bottom, below + extra, inset


def _windows_script_parts(font: Font, platform: xlformat.Platform) -> tuple[float, float, float]:
    """``script_parts`` on Excel for Windows: the Mac's rules in pixels, by ppem.

    Where the Mac takes a whole size ``s`` in points, Excel for Windows takes the font's ppem
    ``p``: it draws at ``round(2p/3)``, raises ``floor(p/2)`` and lowers ``floor(p/5)`` pixels,
    with Windows' faces at those ppems, and a Calibri superscript's part above the baseline is
    the Mac's table read in pixels by ppem, carried on to ppem 53 (captured,
    ``windows/row_height_fitted`` groups ``styled`` and ``scripts``: Calibri at every whole and
    half size from 6 to 40, 207 of 207 rows by these rules, a Calibri subscript at ppem 15 from
    the capture; sizes sharing a ppem share a row). Other faces and sizes past the tables take
    the raise plus the drawn size's ascent (unverified).
    """
    p = ppem(min(max(font.size, 1.0), 409.0))
    d, r, lower = math.floor(2 * p / 3 + 0.5), p // 2, p // 5
    whole = replace(font, size=ppem_size(p), vert_align=None)
    _, ascent, bottom, extra, inset = _face(whole, platform)
    drawn = _face(replace(font, size=ppem_size(d), vert_align=None, underline=None), platform)
    calibri = font.name.lower() == "calibri"
    if font.vert_align == "superscript":
        table = {**_SUPERSCRIPT_ABOVE, **_WINDOWS_SUPERSCRIPT_ABOVE} if calibri else {}
        above = 0.75 * table[p] if p in table else 0.75 * r + drawn[1]
        return above, bottom + extra, inset
    if calibri and p in _WINDOWS_SUBSCRIPT_BELOW:
        below = 0.75 * _WINDOWS_SUBSCRIPT_BELOW[p]
    else:
        below = 0.75 * lower + drawn[2]
    return ascent + below - bottom, below + extra, inset
