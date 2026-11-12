"""Excel's column letters and row numbers, drawn when the backend's options ask for them."""

from __future__ import annotations

from typing import TYPE_CHECKING, TextIO

from sheetpeek._rendered import RenderedRow, RenderedSheet

if TYPE_CHECKING:
    from sheetpeek.backends.html import HTMLBackend
    from sheetpeek.backends.html._table import Page

# Heading bands in sheet points, so they scale with the page like the cells (§4.14). Excel 16.114
# captures at Normal fonts Calibri 11 and 20: the column band is 21 and 25 pt; the row band is
# 16 + 3 pt a digit of the last row at 11 (2, 3 and 4 digits), and 36 pt with 2 digits at 20.
_SIZES = (11.0, 20.0)
_COLUMN_BAND = (21.0, 25.0)
_ROW_BAND_BASE, _ROW_BAND_DIGIT = 16.0, 3.0
_ROW_BAND_GROWTH = 36.0 / 22.0  # at 20 pt, over 11 pt, with 2 digits


def bands(normal_size: float, digits: int) -> tuple[float, float]:
    """The row band's width and the column band's height in points.

    They follow the Normal font's size and the digits of the last row shown. Between and beyond
    the two captured sizes both grow linearly with the size (unverified); they stay positive
    over every size Excel accepts (1 to 409 pt).
    """
    t = (normal_size - _SIZES[0]) / (_SIZES[1] - _SIZES[0])
    column = _COLUMN_BAND[0] + t * (_COLUMN_BAND[1] - _COLUMN_BAND[0])
    row = (_ROW_BAND_BASE + _ROW_BAND_DIGIT * digits) * (1 + t * (_ROW_BAND_GROWTH - 1))
    return row, column


def offsets(sheet: RenderedSheet, backend: HTMLBackend) -> tuple[float, float]:
    """Where the sheet starts, in points: past the row headings and below the column headings.

    The row band follows the digits of the last row on the page, where Excel's follows the last
    row in its window (a choice: a page shows every row).
    """
    digits = max((len(str(r.index)) for r in sheet.rows), default=1)
    row, column = bands(sheet.sheet.normal_font.size, digits)
    return (row if backend.row_headings else 0.0), (column if backend.column_headings else 0.0)


def write_column(page: Page, out: TextIO) -> None:
    """The ``<col>`` of the row headings."""
    if page.backend.row_headings:
        out.write(f'<col style="width:{page.x(0) - page.gx}px">')


def write_header(page: Page, out: TextIO, lead: str = "") -> None:
    """The row of column letters, after ``lead``, the cells of the gutter before them."""
    if not page.backend.column_headings:
        return
    out.write(f'<tr style="height:{page.y(0) - page.gy}px">{lead}')
    if page.backend.row_headings:
        out.write('<th class="sp-corner"></th>')
    for c in page.sheet.columns:
        out.write(f'<th class="sp-colhead" scope="col">{c.letter}</th>')
    out.write("</tr>")


def write_row(page: Page, row: RenderedRow, out: TextIO) -> None:
    """A row's number."""
    if page.backend.row_headings:
        out.write(f'<th class="sp-rowhead" scope="row">{row.index}</th>')
