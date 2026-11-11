"""Excel's column letters and row numbers, drawn when the backend's options ask for them."""

from __future__ import annotations

from typing import TYPE_CHECKING, TextIO

from sheetpeek._rendered import RenderedRow, RenderedSheet

if TYPE_CHECKING:
    from sheetpeek.backends.html import HTMLBackend
    from sheetpeek.backends.html._table import Page

# Placeholders until the screen capture: heading sizes in points.
_COLUMN_HEADING_HEIGHT = 16.0
_ROW_HEADING_DIGIT = 6.0


def offsets(sheet: RenderedSheet, backend: HTMLBackend) -> tuple[float, float]:
    """Where the sheet starts, in points: past the row headings and below the column headings."""
    digits = max((len(str(r.index)) for r in sheet.rows), default=1)
    ox = _ROW_HEADING_DIGIT * max(3, digits) + 6 if backend.row_headings else 0.0
    oy = _COLUMN_HEADING_HEIGHT if backend.column_headings else 0.0
    return ox, oy


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
