"""HTML: a static page drawn from ``excel`` mode's geometry (``ARCHITECTURE.md`` §7.3).

A fixed-layout table carries the cells, their fills and their gridlines. Each text is an
absolutely positioned span inside its origin cell, placed in sheet coordinates against the
wrapper, with its clip window around it; the cells themselves are not positioned, so the
browser paints every background first and every text over it, as Excel paints spilled text
over a neighbor's fill. Every position is a whole CSS pixel rounded from its exact point, and
every size the difference of two rounded positions, so nothing drifts on large sheets. Calibri
runs are spaced from Carlito's widths to Excel's whole-point glyphs. The output has no script.

The custom properties ``--sheetpeek-*`` and the class names ``sheetpeek`` and ``sp-*`` are a
public interface for callers' stylesheets.
"""

from __future__ import annotations

import html
from collections.abc import Sequence
from typing import Literal, TextIO

from sheetpeek._model import one_line
from sheetpeek._notices import Notice, Support, merge
from sheetpeek._rendered import Mode, RenderedSheet
from sheetpeek._view_model import MAX_ZOOM, MIN_ZOOM
from sheetpeek.backends import _support
from sheetpeek.backends.html._buttons import css as buttons_css
from sheetpeek.backends.html._css import CSS
from sheetpeek.backends.html._notices import write_notices
from sheetpeek.backends.html._outline import css as outline_css
from sheetpeek.backends.html._panes import css as panes_css
from sheetpeek.backends.html._panes import write as write_sheet
from sheetpeek.backends.html._patterns import letters
from sheetpeek.backends.html._table import Page

ColorScheme = Literal["light", "dark"]
Buttons = Literal["none", "static"]
Freeze = Literal["panes", "static"]


def check_zoom(zoom: int | None) -> int | None:
    """``zoom`` if it is ``None`` or a whole percent Excel allows, else ``ValueError``."""
    if zoom is None:
        return None
    if isinstance(zoom, bool) or not isinstance(zoom, int) or not MIN_ZOOM <= zoom <= MAX_ZOOM:
        raise ValueError(f"zoom must be a whole number from {MIN_ZOOM} to {MAX_ZOOM}, not {zoom!r}")
    return zoom


class HTMLBackend:
    """A static HTML page, or a fragment, in ``excel`` mode.

    Excel draws white cells in its light and dark appearance alike, so the paper is white and
    the automatic ink black; ``color_scheme`` changes only the headings, as in Excel. Colors
    are CSS custom properties on ``.sheetpeek`` (``--sheetpeek-paper``, ``-ink``, ``-grid``,
    ``-heading-paper``, ``-heading-ink``, ``-heading-grid``), which a caller's stylesheet can
    set.

    Args:
        fragment: Write only the sheet (its stylesheet and a ``<div class="sheetpeek">``), for
            embedding, instead of a whole page.
        column_headings: Add Excel's column letters above the sheet.
        row_headings: Add Excel's row numbers before each row.
        color_scheme: ``"light"`` or ``"dark"``, for the headings; never read from the system.
        notices: Add, after the sheet, a collapsed list of what the page does not show
            (``<details class="sp-notices">``).
        new_tab: Open links to web pages and e-mail addresses in a new tab
            (``target="_blank"``). Links within the page are not affected.
        local_links: Also link ``file:`` URLs and relative, UNC and drive-letter paths, which
            Excel follows. Off by default: such links point into the viewer's own files, they
            reveal the paths the workbook names, and a target starting with ``//`` resolves to a
            network host when a page is served, as an ``http`` link would.
        buttons: Whether to draw Excel's buttons, which do nothing on a static page. For now
            this is the outline: ``"static"`` (the default) draws the outline gutter beside the
            headings as Excel shows it, with its bars, dots, minus and plus buttons, and the level
            buttons where the headings are drawn; ``"none"`` leaves the gutter out. A sheet's
            outline is drawn only when the file shows outline symbols.
        zoom: The zoom in percent, 10 to 400, as Excel's zoom box takes it. By default
            (``None``), the zoom saved with each sheet, as Excel opens it. Cells, text, fills
            and borders' positions and the headings scale; lines, pattern tiles and the outline's
            symbols keep their size in screen points, as in Excel. The page scales the layout
            Excel makes at 100 %, so text Excel wraps or places differently at another zoom (it
            lays out each zoom in whole screen pixels) is not reproduced. Only HTML has a zoom.
        freeze: How a sheet's frozen pane is drawn. ``"panes"`` keeps the frozen rows and
            columns in place while the page scrolls, through CSS ``position: sticky`` (no
            script); ``"static"`` draws only the freeze lines. By default (``None``), panes on a
            whole page and static in a fragment, whose host page decides how it scrolls. Every
            row and column of the range is drawn either way; the scroll position saved in the
            file is not followed.
        page_breaks: Mark the sheet's manual page breaks with a solid blue line, as Excel's page
            break preview draws them; off by default, as in Excel's Normal view.
    """

    name = "html"
    mode: Mode = "excel"

    def __init__(
        self,
        *,
        fragment: bool = False,
        column_headings: bool = False,
        row_headings: bool = False,
        color_scheme: ColorScheme = "light",
        notices: bool = False,
        new_tab: bool = False,
        local_links: bool = False,
        buttons: Buttons = "static",
        zoom: int | None = None,
        freeze: Freeze | None = None,
        page_breaks: bool = False,
    ) -> None:
        if color_scheme not in ("light", "dark"):
            raise ValueError(f"color_scheme must be 'light' or 'dark', not {color_scheme!r}")
        self.fragment = fragment
        self.column_headings = column_headings
        self.row_headings = row_headings
        self.color_scheme = color_scheme
        self.notices = notices
        self.new_tab = new_tab
        self.local_links = local_links
        if buttons not in ("none", "static"):
            raise ValueError(f"buttons must be 'none' or 'static', not {buttons!r}")
        self.buttons = buttons
        self.zoom = check_zoom(zoom)
        if freeze not in (None, "panes", "static"):
            raise ValueError(f"freeze must be 'panes', 'static' or None, not {freeze!r}")
        self.freeze = freeze
        self.page_breaks = page_breaks

    @property
    def support(self) -> dict[str, Support]:
        """What this output shows of each feature the census finds (``backends/_support.py``)."""
        return _support.html(
            headings=self.column_headings or self.row_headings,
            buttons=self.buttons,
            level_buttons=self.column_headings and self.row_headings,
            local_links=self.local_links,
        )

    def render(self, sheet: RenderedSheet, out: TextIO, /) -> None:
        """Write the page or the fragment to ``out``."""
        page = Page(sheet, self)
        stylesheet = (
            f"{CSS}{outline_css(sheet, self)}{buttons_css(sheet, self)}"
            f"{panes_css(sheet, self)}{page.css()}"
        )
        self._head(sheet, stylesheet, out)
        write_sheet(page, out)
        if self.notices:
            write_notices(sheet, self, out)
        if not self.fragment:
            out.write("</body>\n</html>\n")

    def _head(self, sheet: RenderedSheet, stylesheet: str, out: TextIO) -> None:
        """The document's start and the stylesheet, or the stylesheet alone in a fragment."""
        if not self.fragment:
            title = html.escape(one_line(sheet.sheet.name))
            background = "#1e1e1e" if self.color_scheme == "dark" else "#ffffff"
            out.write(
                '<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
                '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
                f"<title>{title}</title>\n"
                f"<style>\n{stylesheet}"
                f"body {{ margin: 8px; background: {background}; }}\n</style>\n"
                "</head>\n<body>\n"
            )
        else:
            out.write(f"<style>\n{stylesheet}</style>\n")

    def render_blocks(
        self,
        sheets: Sequence[RenderedSheet],
        out: TextIO,
        notices: tuple[Notice, ...] | None = None,
        /,
    ) -> None:
        """Write several ranges of one sheet, such as its print areas, one below another.

        Each block is a ``<div class="sp-area">`` holding the sheet's markup; the notices, when
        asked for, follow the last block: every block's, merged (``notices``, else worked out
        here). Fill and border classes are made unique to each block.
        """
        pages = [Page(s, self, scope=f"sp-{k}", tag=letters(k) + "-") for k, s in enumerate(sheets)]
        shared = {outline_css(s, self) for s in sheets} | {panes_css(s, self) for s in sheets}
        shared |= {buttons_css(s, self) for s in sheets}
        stylesheet = CSS + "".join(sorted(shared - {""})) + "".join(p.css() for p in pages)
        stylesheet += ".sp-area + .sp-area { margin-top: 16px; }\n"
        self._head(sheets[0], stylesheet, out)
        for page in pages:
            out.write('<div class="sp-area">')
            write_sheet(page, out)
            out.write("</div>\n")
        if self.notices:
            if notices is None:
                from sheetpeek._census import output_notices

                notices = merge(
                    n
                    for s in sheets
                    for n in s.notices + output_notices(s, self.name, self.support)
                )
            write_notices(None, self, out, notices)
        if not self.fragment:
            out.write("</body>\n</html>\n")


__all__ = ["CSS", "Buttons", "ColorScheme", "Freeze", "HTMLBackend"]
