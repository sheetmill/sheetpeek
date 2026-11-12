"""Hyperlinks: an anchor around each linked cell's runs, and ids on the cells links go to.

The anchor wraps the cell's text runs, so text spilled over its neighbors stays clickable. It
adds no look of its own (``.sheetpeek a`` inherits color and decoration): the cell's format draws
a link, as in Excel. ``title`` is the file's tooltip only, never a target.
"""

from __future__ import annotations

import html
from typing import TYPE_CHECKING

from sheetpeek._census import shown_cells
from sheetpeek._hyperlink import classify

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
        self.ids: dict[tuple[int, int], str] = {}
        """Grid position of a cell a link goes to → its id."""
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
                    anchor = self.ids.setdefault(place, _id(page, place))
                    attributes = f' href="#{anchor}"'
                if link.tooltip:
                    attributes += f' title="{html.escape(link.tooltip)}"'
                self.opening[(cell.row, cell.column)] = f"<a{attributes}>"


def _id(page: Page, place: tuple[int, int]) -> str:
    """The id of the cell drawn at a grid position: ``sp-`` and its coordinate."""
    return f"sp-{page.sheet.rows[place[0]].cells[place[1]].cell.coordinate}"
