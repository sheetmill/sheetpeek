"""Row heights in points, as Excel for Mac lays them out (``ARCHITECTURE.md`` §4.7).

A row's own ``ht`` wins. Otherwise ``<sheetFormatPr>`` decides, together with whether the file
has a ``<row>`` element for the row: Excel ignores ``defaultRowHeight`` unless ``customHeight``
is set, and gives the row the Normal font's height from the captured table below.
"""

from __future__ import annotations

from functools import lru_cache

from sheetpeek._model import Font, Row, Sheet

# Captured (suite row_height, Excel 16.114 and 16.115): the height of a row without ``ht`` per
# Normal font. Calibri 20 is §4.16's runs capture of a row fitted to a 20 pt cell (Excel 16.114).
FONT_ROW_HEIGHTS: dict[str, dict[int, float]] = {
    "calibri": {9: 12, 10: 14, 11: 15, 12: 16, 14: 19, 16: 21, 20: 26},
    "aptos narrow": {9: 13, 10: 14, 11: 15, 12: 16, 14: 19, 16: 22},
}
NO_SHEET_FORMAT = 16.0  # an unlisted row of a file without <sheetFormatPr>, every font
WRAPPED_LEADING = 1.0  # wrapped lines are the font's row height plus 1 apart (Calibri 11: 16)


@lru_cache(maxsize=256)
def font_row_height(font: Font) -> float:
    """The captured row height for a font, at the nearest measured size.

    A family not in the table uses Calibri's heights, as xlformat measures it as Calibri
    (unverified, as are sizes between the measured ones).
    """
    table = FONT_ROW_HEIGHTS.get(font.name.lower(), FONT_ROW_HEIGHTS["calibri"])
    size = min(table, key=lambda s: (abs(s - font.size), s))
    return float(table[size])


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
    largest cell font, merged cells included, or to its wrapped lines, merged ones excluded, and
    never below the Normal font's height. A stored ``ht`` and ``customHeight`` defaults are kept.
    """
    return row.height is None and row.listed and not sheet.custom_row_height


def wrapped_height(lines: int, font: Font) -> float:
    """The height Excel gives a row for ``lines`` wrapped lines in ``font``.

    Captured: three and six lines of Calibri 11 are 48 and 96 pt; other fonts are unverified.
    """
    if lines <= 1:
        return font_row_height(font)
    return lines * (font_row_height(font) + WRAPPED_LEADING)


def baseline_offsets(font: Font) -> tuple[float, float, float]:
    """(line height, baseline below the line's top, baseline above the line's bottom).

    Provisional, from print at Calibri 11 (``layout`` rows 1 and 152 to 154): a 15 pt line with
    its baseline 11 pt below the top of a top-aligned block and 3 pt above the bottom of a
    bottom-aligned one. Other fonts scale the 11 with their line height.
    """
    line = font_row_height(font)
    above = round(line * 11 / 15)
    return line, float(above), line - above - 1
