"""Hyperlinks: an anchor around each linked cell's runs, and ids on the cells links go to.

The anchor wraps the cell's text runs, so text spilled over its neighbors stays clickable. It
adds no look of its own (``.sheetpeek a`` inherits color and decoration): the cell's format draws
a link, as in Excel. ``title`` is the file's tooltip only, never a target. On the workbook page a
link to a cell of another sheet on the page goes there (:class:`BookLinks`).
"""

from __future__ import annotations

import html
from collections.abc import Callable
from typing import TYPE_CHECKING

from sheetpeek._census import shown_cells
from sheetpeek._hyperlink import Hyperlink, classify, location_place

if TYPE_CHECKING:
    from sheetpeek.backends.html._table import Page


class Links:
    """The opening tag of each linked cell's anchor, and the id of each cell a link goes to."""

    def __init__(self, page: Page, holder: dict[tuple[int, int], tuple[int, int]]) -> None:
        """Find the links of the page's cells.

        Args:
            page: The page.
            holder: Sheet position → (row, column) in the page's grid of the cell that draws it.
        """
        sheet = page.sheet
        name = sheet.sheet.name
        shown = shown_cells(sheet)
        new_tab = ' target="_blank"' if page.backend.new_tab else ""
        self.opening: dict[tuple[int, int], str] = {}
        """Sheet position of a linked cell → its anchor's opening tag."""
        for row in sheet.rows:
            for rendered in row.cells:
                cell = rendered.cell
                link = cell.hyperlink
                if link is None or cell.kind == "empty" or cell.value == "":
                    continue
                kind, url, target = classify(link, name, shown, page.backend.local_links)
                if kind == "refused":
                    continue
                attributes = ""
                if kind == "url" and url is not None:
                    attributes = f' href="{html.escape(url)}" rel="noopener noreferrer"{new_tab}'
                elif target is not None and (place := holder.get(target)) is not None:
                    anchor = page.ids.setdefault(place, _id(page, place))
                    attributes = f' href="#{anchor}"'
                elif page.book is not None and (found := page.book.find(link, name)):
                    other, place = found
                    anchor = other.ids.setdefault(place, _id(other, place))
                    attributes = f' href="#{anchor}"'
                if link.tooltip:
                    attributes += f' title="{html.escape(link.tooltip)}"'
                self.opening[(cell.row, cell.column)] = f"<a{attributes}>"


def _id(page: Page, place: tuple[int, int]) -> str:
    """The id of the cell drawn at a grid position: the page's scope (``sp``), its coordinate."""
    return f"{page.scope}-{page.sheet.rows[place[0]].cells[place[1]].cell.coordinate}"


class BookLinks:
    """The sheets of a workbook page, for links from one of them to a cell of another."""

    def __init__(self, pages: list[Page]) -> None:
        self.pages: dict[str, Page] = {}
        """Each page by its sheet's name, compared ignoring case as Excel compares them."""
        for page in pages:
            self.pages.setdefault(page.sheet.sheet.name.casefold(), page)

    def find(self, link: Hyperlink, sheet: str) -> tuple[Page, tuple[int, int]] | None:
        """The page and grid position a link on ``sheet`` goes to on another sheet, if drawn.

        A link to its own sheet is the page's own (:func:`~sheetpeek._hyperlink.classify`); a
        cell not drawn, in a hidden row or column or past the used range, has no grid position.
        """
        if link.target is not None or not link.location:
            return None
        place = location_place(link.location, sheet)
        page = self.pages.get(place[0].casefold()) if place is not None else None
        if place is None or page is None:
            return None
        held = page.grid()[1].get((place[1], place[2]))
        return (page, held) if held is not None else None

    def follows(self, sheet: str) -> Callable[[Hyperlink], bool]:
        """Whether the page follows a link on ``sheet`` to another sheet, for the census."""
        return lambda link: self.find(link, sheet) is not None
