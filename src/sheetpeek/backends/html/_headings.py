"""Excel's column letters and row numbers, drawn when the backend's options ask for them."""

from __future__ import annotations

from typing import TYPE_CHECKING, TextIO

from sheetpeek._rendered import RenderedRow, RenderedSheet

if TYPE_CHECKING:
    from sheetpeek.backends.html import HTMLBackend
    from sheetpeek.backends.html._table import Page

# Heading bands in sheet points, so they scale with the page like the cells (§4.14; screen suite
# headings_grid). At the Normal font's size, with row numbers of 2 digits: the row band's width
# and the column band's height, at the captured sizes (Calibri 8 to 36; Aptos Narrow 11 as
# Calibri 11); between them both follow the line through the two nearest, beyond them the last
# segment's (unverified).
_SIZES = (8.0, 11.0, 14.0, 20.0, 28.0, 36.0)
_ROW_BAND = (21.0, 22.0, 25.0, 36.0, 50.0, 64.0)
_COLUMN_BAND = (21.0, 21.0, 21.0, 25.0, 36.0, 46.0)
# At Calibri 11 the row band is max(22, 4 + 6 x digits) pt: from 4 digits on it widens 6 pt a
# digit. At other sizes the digit term scales with the size (unverified).
_DIGIT, _BASE_SIZE, _FREE_DIGITS = 6.0, 11.0, 3


def _along(size: float, values: tuple[float, ...]) -> float:
    """``values`` at ``size``: on the line through the two nearest captured sizes."""
    k = min(max(1, sum(1 for s in _SIZES if s <= size)), len(_SIZES) - 1)
    x0, x1, y0, y1 = _SIZES[k - 1], _SIZES[k], values[k - 1], values[k]
    return y0 + (size - x0) * (y1 - y0) / (x1 - x0)


def bands(normal_size: float, digits: int) -> tuple[float, float]:
    """The row band's width and the column band's height in points.

    They follow the Normal font's size and the digits of the last row shown. Below the smallest
    captured size both keep its bands; they stay positive over every size Excel accepts (1 to
    409 pt).
    """
    size = max(normal_size, _SIZES[0])
    column = _along(size, _COLUMN_BAND)
    row = _along(size, _ROW_BAND)
    row += max(0, digits - _FREE_DIGITS) * _DIGIT * size / _BASE_SIZE
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
