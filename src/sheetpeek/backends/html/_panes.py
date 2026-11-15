"""Lines over the sheet: frozen panes, and page-break marks on request (§4.18).

With ``freeze="panes"`` the sheet is drawn once, and the frozen rows, the frozen columns and
their corner are drawn again, each copy holding only its own rows or columns, stacked over the
sheet in one grid cell and held in place by CSS ``position: sticky`` while the page scrolls. This
is layout, not a script: a browser without sticky positioning shows the copies where they stand,
exactly over the sheet, which is the static drawing. With ``freeze="static"`` only the freeze
lines are drawn.
"""

from __future__ import annotations

import bisect
from typing import TYPE_CHECKING, TextIO

from sheetpeek._model import CellRange
from sheetpeek._rendered import RenderedSheet

if TYPE_CHECKING:
    from sheetpeek.backends.html import HTMLBackend
    from sheetpeek.backends.html._table import Page

FREEZE_LINE = "9b9b9b"
"""Excel's freeze line, gray 155, one screen point wide (captured on Excel 16.114 only)."""

CSS = f"""\
.sheetpeek {{ --sheetpeek-freeze: #{FREEZE_LINE}; }}
.sheetpeek .sp-freeze {{ position: absolute; background: var(--sheetpeek-freeze); }}
.sp-panes {{ display: inline-grid; vertical-align: top; }}
.sp-panes > * {{ grid-area: 1 / 1; align-self: start; justify-self: start; }}
.sp-panes > .sp-pane {{ position: sticky; overflow: hidden; }}
.sp-panes > .sp-pane-left {{ left: 0; z-index: 1; }}
.sp-panes > .sp-pane-top {{ top: 0; z-index: 2; }}
.sp-panes > .sp-pane-corner {{ top: 0; left: 0; z-index: 3; }}
"""
"""The stylesheet of a page with a frozen pane."""


def frozen(sheet: RenderedSheet) -> tuple[int, int]:
    """The drawn rows and columns before the freeze lines, ``(0, 0)`` without a frozen pane.

    A frozen pane splits the sheet before row ``top_left + frozen_rows`` and column ``top_left +
    frozen_columns``. The saved scroll position is not followed (decided): every row and column
    of the range is drawn, those before the split frozen. A split outside the drawn range freezes
    nothing.
    """
    view = sheet.sheet.view
    if view.pane_state not in ("frozen", "frozenSplit"):
        return 0, 0
    try:
        top = CellRange.parse(view.top_left)
    except ValueError:
        top = CellRange(1, 1, 1, 1)
    rows = columns = 0
    if view.frozen_rows:
        split = top.min_row + view.frozen_rows
        rows = sum(1 for r in sheet.rows if r.index < split)
        rows = rows if rows < len(sheet.rows) else 0
    if view.frozen_columns:
        split = top.min_col + view.frozen_columns
        columns = sum(1 for c in sheet.columns if c.index < split)
        columns = columns if columns < len(sheet.columns) else 0
    return rows, columns


BREAK_LINE = "0000d0"
"""A manual page break's mark, as page break preview draws it (unverified)."""

BREAK_CSS = f"""\
.sheetpeek {{ --sheetpeek-page-break: #{BREAK_LINE}; }}
.sheetpeek .sp-pagebreak {{ position: absolute; background: var(--sheetpeek-page-break); }}
"""
"""The stylesheet of a page with page-break marks."""


def css(sheet: RenderedSheet, backend: HTMLBackend) -> str:
    """The stylesheet of a sheet with a frozen pane or, when asked for, page-break marks."""
    found = CSS if frozen(sheet) != (0, 0) else ""
    if backend.page_breaks and _breaks(sheet) != ([], []):
        found += BREAK_CSS
    return found


def _breaks(sheet: RenderedSheet) -> tuple[list[int], list[int]]:
    """The drawn rows and columns that start after a manual break inside the drawn range.

    A break below row ``r`` (``Sheet.row_breaks``) falls before the first drawn row after ``r``,
    when some drawn row is at or before ``r``; columns likewise.
    """
    model = sheet.sheet
    rows = [r.index for r in sheet.rows]
    columns = [c.index for c in sheet.columns]

    def starts(breaks: tuple[int, ...], drawn: list[int]) -> list[int]:
        found = {bisect.bisect_right(drawn, b) for b in breaks}
        return sorted(k for k in found if 0 < k < len(drawn))

    return starts(model.row_breaks, rows), starts(model.column_breaks, columns)


def marks(page: Page) -> str:
    """The page-break marks: 3 CSS px, centered on the gridline after each break."""
    rows, columns = _breaks(page.sheet)
    sheet = page.sheet
    left, top = page.x(0), page.y(0)
    last = sheet.rows[-1]
    bottom = page.y((last.y_pt or 0) + (last.height_pt or 0)) + 1
    found = [
        f'<i class="sp-pagebreak" style="left:{left}px;top:{page.y(sheet.rows[k].y_pt or 0) - 1}px;'
        f'width:{page.width - left}px;height:3px"></i>'
        for k in rows
    ]
    found += [
        f'<i class="sp-pagebreak" style="left:{page.x(sheet.columns[k].x_pt or 0) - 1}px;'
        f'top:{top}px;width:3px;height:{bottom - top}px"></i>'
        for k in columns
    ]
    return "".join(found)


def write(page: Page, out: TextIO) -> None:
    """Write the sheet, with its frozen pane drawn as the backend's ``freeze`` option says."""
    breaks = marks(page) if page.backend.page_breaks else ""
    rows, columns = frozen(page.sheet)
    if (rows, columns) == (0, 0):
        page.write(out, lines=breaks)
        return
    sheet = page.sheet
    y = page.y(sheet.rows[rows].y_pt or 0) if rows else None
    x = page.x(sheet.columns[columns].x_pt or 0) if columns else None
    last_row = sheet.rows[-1]
    bottom = page.y((last_row.y_pt or 0) + (last_row.height_pt or 0)) + 1
    lines = breaks
    if y is not None:
        left = page.x(0)
        lines += f'<i class="sp-freeze" style="left:{left}px;top:{y}px;'
        lines += f'width:{page.width - left}px;height:1px"></i>'
    if x is not None:
        top = page.y(0)
        lines += f'<i class="sp-freeze" style="left:{x}px;top:{top}px;'
        lines += f'width:1px;height:{bottom - top}px"></i>'
    backend = page.backend
    mode = backend.freeze or ("static" if backend.fragment else "panes")
    if mode == "static":
        page.write(out, lines=lines)
        return
    out.write('<div class="sp-panes">')
    page.write(out, lines=lines)
    if x is not None:
        out.write(f'<div class="sp-pane sp-pane-left" style="width:{x + 1}px">')
        page.write(out, columns=columns, lines=lines)
        out.write("</div>")
    if y is not None:
        out.write(f'<div class="sp-pane sp-pane-top" style="height:{y + 1}px">')
        page.write(out, rows=rows, lines=lines)
        out.write("</div>")
    if x is not None and y is not None:
        out.write(f'<div class="sp-pane sp-pane-corner" style="width:{x + 1}px;height:{y + 1}px">')
        page.write(out, rows=rows, columns=columns, lines=lines)
        out.write("</div>")
    out.write("</div>\n")
    page.copy = False
