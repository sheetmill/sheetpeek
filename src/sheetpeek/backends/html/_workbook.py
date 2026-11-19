"""The workbook page: every sheet of a workbook on one static HTML page, with a tab strip.

A ``<section>`` per sheet drawn by :class:`HTMLBackend`, and a ``<nav>`` of tabs, one per sheet
in workbook order (hidden ones only on request, as Excel's strip leaves them out), below the
sheets as Excel draws it, or above them on request. Without
script, a tab is a link to its section's id; a browser that has CSS's ``:has()`` shows the
section that is the URL's target, or holds it, and the active sheet when none is. Other browsers
show the sections one after another. A link to a cell of another sheet on the page goes there.
``ARCHITECTURE.md`` §4.15 has the rules.
"""

from __future__ import annotations

import html
import io
import os
import warnings
from collections.abc import Callable
from typing import IO, Any, Literal, TextIO

from sheetpeek._model import one_line
from sheetpeek._notices import Collector, Notice, NotShownWarning, deliver
from sheetpeek._package import LimitError, ReadLimits
from sheetpeek._rendered import RenderedSheet
from sheetpeek._sheets import SheetInfo
from sheetpeek.backends.html import HTMLBackend
from sheetpeek.backends.html._buttons import css as buttons_css
from sheetpeek.backends.html._css import CSS
from sheetpeek.backends.html._links import BookLinks
from sheetpeek.backends.html._notices import write_notices
from sheetpeek.backends.html._outline import css as outline_css
from sheetpeek.backends.html._panes import css as panes_css
from sheetpeek.backends.html._panes import write as write_sheet
from sheetpeek.backends.html._patterns import letters
from sheetpeek.backends.html._table import Page

Source = str | os.PathLike[str] | IO[bytes]

ACTIVE_INK = "278452"
"""The active tab's text: Excel's own green, whatever the tab's color (unverified)."""
_TOP_SHARE, _BOTTOM_SHARE = 0.1, 0.4  # the active tab's gradient, as shares of its color
_DARK_TEXT_FROM = 128
"""An inactive colored tab's name is black from this Rec. 601 luma of its color up, else white
(captured, tab_strip: white on luma 75 to 113, black on 170 to 188; the edge between is
unverified)."""
_NOT_SHOWN = {
    "chartsheet": "is a chartsheet: charts are not shown.",
    "dialogsheet": "is a dialog sheet, a legacy form: it is not shown.",
}
_STATE_TITLES = {"hidden": "hidden sheet", "very_hidden": "very hidden sheet"}

_BOOK_CSS = """\
.sp-book > nav { display: flex; flex-wrap: wrap; gap: 4px 9px; margin: 0;
  font: 12px/27px system-ui, sans-serif; }
.sp-book > nav a { display: block; min-width: 96px; padding: 0 10px; box-sizing: border-box;
  text-align: center; white-space: pre; text-decoration: none; border-radius: 5px; }
.sp-book > nav a.sp-hidden { font-style: italic; }
.sp-book > section { margin: 0 0 16px; }
.sp-book .sp-notshown { font: 13px/1.4 system-ui, sans-serif; }
"""
_SCHEMES = {  # an uncolored tab: (paper, ink), and the active one's
    "light": (("e6e6e6", "444444"), ("ffffff", ACTIVE_INK)),
    "dark": (("3a3a3a", "bbbbbb"), ("ffffff", ACTIVE_INK)),
}
_HAS = "selector(:has(a))"
_NONE_TARGETED = ".sp-book:not(:has(> section:target, > section :target))"


def render_workbook(
    source: Source | Any,
    *,
    hidden_sheets: bool = False,
    show_hidden: bool = False,
    tabs: Literal["bottom", "top"] = "bottom",
    out: TextIO | None = None,
    limits: ReadLimits | None = None,
    on_notice: Callable[[Notice], object] | Literal["ignore"] | None = None,
    **html_options: Any,
) -> str | None:
    """Render every sheet of a workbook as one HTML page, with Excel's tab strip.

    Each worksheet is drawn by the ``html`` backend in a section of its own; a tab shows each
    sheet in workbook order, in its tab color. A chartsheet or dialog sheet gets a tab and a
    section that says it is not shown. A link to a cell of another sheet on the page goes there.
    The page has no script. Notices are delivered as by :func:`~sheetpeek.render`: per sheet,
    then one for the workbook (sheets not shown).

    Args:
        source: A workbook: a path, a binary file object, or an open :class:`~sheetpeek.Book`
            (left open).
        hidden_sheets: Also draw hidden and very hidden sheets, each with a tab. Excel's tab
            strip leaves them out, and so does the page by default.
        show_hidden: Also draw the hidden rows and columns of every sheet, as
            :meth:`Sheet.layout <sheetpeek.Sheet.layout>` does.
        tabs: Where the tab strip goes: ``"bottom"``, below the sheets as Excel draws it, or
            ``"top"``.
        out: A text stream to write to. Without one, the page is returned as a string.
        limits: What reading may cost. ``max_cells`` bounds the whole page's cells, every
            sheet's together. By default, ``ReadLimits()``.
        on_notice: Called with each notice, ``"ignore"``, or ``None`` for one
            :class:`~sheetpeek.NotShownWarning` per sheet with something to report.
        **html_options: Options of :class:`~sheetpeek.backends.HTMLBackend`, such as
            ``column_headings=True``; ``fragment=True`` writes the stylesheet and a
            ``<div class="sp-book">`` instead of a whole page.

    Returns:
        The page, or ``None`` when it was written to ``out``.

    Raises:
        LimitError: The file, a sheet, or the page's cells exceed ``limits``.
        PackageError: The file is not a readable xlsx workbook, or it is refused.
    """
    from sheetpeek._read import Book, book_limits

    if tabs not in ("bottom", "top"):
        raise ValueError(f"tabs must be 'bottom' or 'top', not {tabs!r}")
    backend = HTMLBackend(**html_options)
    if isinstance(source, Book):
        book, owned = source, False
    else:
        book, owned = Book(source, limits=limits), True
    try:
        page = _WorkbookPage(book, backend, hidden_sheets, show_hidden, limits or book_limits(book))
        page.tabs_first = tabs == "top"
    finally:
        if owned:
            book.close()
    if out is not None:
        page.write(out)
        result = None
    else:
        buffer = io.StringIO()
        page.write(buffer)
        result = buffer.getvalue()
    for name, notices in page.sheet_notices:
        deliver(notices, on_notice, name, backend.name, 2)
    _deliver_book(page.book_notices, on_notice, backend.name)
    return result


class _WorkbookPage:
    """The sheets read and laid out, their pages, tabs and notices, ready to write."""

    def __init__(
        self,
        book: Any,
        backend: HTMLBackend,
        hidden_sheets: bool,
        show_hidden: bool,
        limits: ReadLimits,
    ) -> None:
        self.backend = backend
        self.tabs_first = False
        from sheetpeek._read import active_sheet, book_file_name

        self.title = book_file_name(book) or "workbook"
        infos: tuple[SheetInfo, ...] = book.sheets
        self.tabs: list[tuple[int, SheetInfo]] = [
            (index, info)
            for index, info in enumerate(infos, 1)
            if hidden_sheets or info.state == "visible"
        ]
        active = active_sheet(book)
        shown = [info.name for _, info in self.tabs]
        if active is not None and active.state == "visible" and active.name in shown:
            self.first = shown.index(active.name)
        else:
            self.first = 0
        self.pages: dict[int, Page] = {}
        cells = 0
        for k, (index, info) in enumerate(self.tabs):
            if not info.readable:
                continue
            rendered = book.sheet(info.name).layout("excel", show_hidden=show_hidden)
            cells += len(rendered.rows) * len(rendered.columns)
            if cells > limits.max_cells:
                raise LimitError(
                    f"the workbook page is over {limits.max_cells} cells by sheet "
                    f"{info.name!r}: render fewer sheets, or one sheet at a time"
                )
            self.pages[index] = Page(rendered, backend, scope=f"sp-{index}", tag=letters(k) + "-")
        links = BookLinks(list(self.pages.values()))
        for page in self.pages.values():
            page.book = links
        for page in self.pages.values():
            page.links  # noqa: B018 - every link first, since one gives an id to another sheet
        self.sheet_notices: list[tuple[str, tuple[Notice, ...]]] = []
        self.notices: dict[int, tuple[Notice, ...]] = {}
        for index, page in self.pages.items():
            found = _sheet_notices(page.sheet, backend, links.follows(page.sheet.sheet.name))
            self.notices[index] = found
            self.sheet_notices.append((page.sheet.sheet.name, found))
        self.book_notices = _book_notices(book, infos, [info.name for _, info in self.tabs])

    def write(self, out: TextIO) -> None:
        backend = self.backend
        shared = {outline_css(p.sheet, backend) for p in self.pages.values()}
        shared |= {buttons_css(p.sheet, backend) for p in self.pages.values()}
        shared |= {panes_css(p.sheet, backend) for p in self.pages.values()}
        css = CSS + "".join(sorted(shared - {""}))
        css += "".join(p.css() for p in self.pages.values()) + _BOOK_CSS + self.tab_css()
        if backend.fragment:
            out.write(f"<style>\n{css}</style>\n")
        else:
            background = "#1e1e1e" if backend.color_scheme == "dark" else "#ffffff"
            out.write(
                '<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
                '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
                f"<title>{html.escape(one_line(self.title))}</title>\n"
                f"<style>\n{css}body {{ margin: 8px; background: {background}; }}\n</style>\n"
                "</head>\n<body>\n"
            )
        out.write('<div class="sp-book">\n')
        if self.tabs_first:
            self.write_tabs(out)
        for k, (index, info) in enumerate(self.tabs):
            first = ' class="sp-first"' if k == self.first else ""
            out.write(f'<section id="sp-{index}"{first}>\n')
            page = self.pages.get(index)
            if page is None:
                name = html.escape(one_line(repr(info.name)))
                out.write(f'<p class="sp-notshown">{name} {_NOT_SHOWN[info.kind]}</p>\n')
            else:
                write_sheet(page, out)
                if backend.notices:
                    write_notices(None, backend, out, self.notices[index])
            out.write("</section>\n")
        if not self.tabs_first:
            self.write_tabs(out)
        if backend.notices:
            write_notices(None, backend, out, self.book_notices)
        out.write("</div>\n")
        if not backend.fragment:
            out.write("</body>\n</html>\n")

    def write_tabs(self, out: TextIO) -> None:
        """The tab strip: a link to each sheet's section, in its tab color."""
        out.write("<nav>")
        for k, (index, info) in enumerate(self.tabs):
            classes = " ".join(
                c
                for c in (
                    "sp-first" if k == self.first else "",
                    "sp-hidden" if info.state != "visible" else "",
                    self.color_class(k),
                )
                if c
            )
            title = _STATE_TITLES.get(info.state)
            attributes = f' class="{classes}"' if classes else ""
            attributes += f' title="{title}"' if title else ""
            label = html.escape(one_line(info.name))
            out.write(f'<a href="#sp-{index}"{attributes}>{label}</a>')
        out.write("</nav>\n")

    def color_class(self, k: int) -> str:
        """The generated class of the ``k``-th tab's color, or ``""`` for an uncolored tab."""
        return f"sp-t-{letters(k)}" if self.tabs[k][1].tab_color else ""

    def tab_css(self) -> str:
        """Each tab's colors, inactive and active, and which section shows (§4.15)."""
        (paper, ink), (active_paper, active_ink) = _SCHEMES[self.backend.color_scheme]
        rules = [f".sp-book > nav a {{ background: #{paper}; color: #{ink}; }}\n"]
        active: list[str] = []
        fallback = ""
        for k, (index, info) in enumerate(self.tabs):
            color = info.tab_color
            selector = f".sp-book > nav a.sp-t-{letters(k)}"
            if color:
                text = "000000" if _luma(color) >= _DARK_TEXT_FROM else "ffffff"
                rules.append(f"{selector} {{ background: #{color}; color: #{text}; }}\n")
                body = (
                    f"background: linear-gradient(#{_mix(color, _TOP_SHARE)}, "
                    f"#{_mix(color, _BOTTOM_SHARE)}); color: #{active_ink}; font-weight: bold;"
                )
            else:
                body = f"background: #{active_paper}; color: #{active_ink}; font-weight: bold;"
            target = f".sp-book:has(#sp-{index}:target, #sp-{index} :target) > nav "
            target += f'a[href="#sp-{index}"]'
            if k == self.first:
                target += f", {_NONE_TARGETED} > nav a.sp-first"
                fallback = f"@supports not {_HAS} {{ .sp-book > nav a.sp-first {{ {body} }} }}\n"
            active.append(f"  {target} {{ {body} }}\n")
        return (
            "".join(rules)
            + f"@supports {_HAS} {{\n"
            + "  .sp-book > section { display: none; }\n"
            + "  .sp-book > section:target, .sp-book > section:has(:target) { display: block; }\n"
            + f"  {_NONE_TARGETED} > section.sp-first {{ display: block; }}\n"
            + "".join(active)
            + "}\n"
            + fallback
        )


def _luma(color: str) -> float:
    """The Rec. 601 luma of an ``RRGGBB`` color, 0 to 255."""
    r, g, b = (int(color[i : i + 2], 16) for i in (0, 2, 4))
    return 0.299 * r + 0.587 * g + 0.114 * b


def _mix(color: str, share: float) -> str:
    """``share`` of an ``RRGGBB`` color mixed with white."""
    channels = (int(color[i : i + 2], 16) for i in (0, 2, 4))
    return "".join(f"{round(255 + (c - 255) * share):02X}" for c in channels)


def _sheet_notices(
    rendered: RenderedSheet, backend: HTMLBackend, follows: Callable[[Any], bool]
) -> tuple[Notice, ...]:
    """A sheet's notices on the page: its tab color is drawn, its links to the page followed."""
    from sheetpeek._census import output_notices

    own = tuple(n for n in rendered.notices if n.code != "sheet.tab_color")
    return own + output_notices(rendered, backend.name, backend.support, follows)


def _book_notices(book: Any, infos: tuple[SheetInfo, ...], drawn: list[str]) -> tuple[Notice, ...]:
    """The workbook's notices for the sheets on the page, and one for hidden sheets left out."""
    kept = tuple(n for n in book.notices if n.context.get("sheet") in drawn)
    left_out = [info for info in infos if info.name not in drawn]
    if not left_out:
        return kept
    from sheetpeek._detect import hidden_sheets

    collector = Collector(None)
    very = sum(1 for info in left_out if info.state == "very_hidden")
    hidden_sheets(collector, [info.name for info in left_out], very)
    return kept + collector.notices()


def _deliver_book(
    notices: tuple[Notice, ...],
    on_notice: Callable[[Notice], object] | Literal["ignore"] | None,
    output: str,
) -> None:
    """The workbook's notices, as :func:`~sheetpeek._notices.deliver` delivers a sheet's."""
    if on_notice is None and notices:
        codes = ", ".join(n.code for n in notices)
        message = (
            f"workbook to {output}: {len(notices)} kind(s) of content not shown ({codes}); "
            "pass on_notice to see them"
        )
        warnings.warn(NotShownWarning(message), stacklevel=3)
        return
    deliver(notices, on_notice, "", output, 2)
