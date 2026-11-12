"""The grid: a fixed-layout table of cells with their fills, gridlines and merges."""

from __future__ import annotations

from typing import TYPE_CHECKING, TextIO, overload

from sheetpeek._model import CellRange
from sheetpeek._rendered import RenderedCell, RenderedSheet, TextRun
from sheetpeek.backends.html import (
    _borders,
    _headings,
    _links,
    _outline,
    _patterns,
    _text,
)

if TYPE_CHECKING:
    from sheetpeek.backends.html import HTMLBackend


_Spans = dict[CellRange, tuple[int, int, int, int]]


def _px(points: float) -> int:
    """A sheet position in whole CSS pixels."""
    return round(points * 4 / 3)


def _grid_css(color: str | None) -> str:
    """A grid color's class: its gridlines one device pixel wide, 0.5 CSS px (§4.14).

    A generated class, like the fill and border layers', not public interface.
    """
    if color is None:
        return ""
    name = f".sheetpeek.sp-g-{color}"
    return (
        f"{name} {{ --sheetpeek-grid: #{color}; }}\n"
        f"{name} table {{ border-right-width: 0.5px; border-bottom-width: 0.5px; }}\n"
        f"{name} td {{ border-left-width: 0.5px; border-top-width: 0.5px; }}\n"
    )


class Holder:
    """Sheet position → grid position of the cell that draws it, as a read-only mapping.

    A cell outside any merge is at its own row and column, found from two small maps instead of
    an entry per cell; a merged cell, and a merge's anchor, are at the merge's first visible cell.
    """

    def __init__(
        self, grid: list[tuple[RenderedCell, ...]], merged: dict[tuple[int, int], tuple[int, int]]
    ) -> None:
        self.merged = merged
        self.rows = {cells[0].row: i for i, cells in enumerate(grid) if cells}
        self.columns = {cell.column: p for p, cell in enumerate(grid[0])} if grid else {}

    @overload
    def get(self, key: tuple[int, int]) -> tuple[int, int] | None: ...
    @overload
    def get(self, key: tuple[int, int], default: tuple[int, int]) -> tuple[int, int]: ...
    def get(
        self, key: tuple[int, int], default: tuple[int, int] | None = None
    ) -> tuple[int, int] | None:
        found = self.merged.get(key)
        if found is not None:
            return found
        i, p = self.rows.get(key[0]), self.columns.get(key[1])
        if i is None or p is None:
            return default
        return i, p


class Page:
    """One sheet's markup.

    Args:
        sheet: The sheet, laid out in ``excel`` mode.
        backend: The backend and its options.
        scope: What the ids of cells that links go to start with: ``sp``, or on the workbook
            page the sheet's section id, so that ids stay unique on the page.
        tag: Letters and a hyphen after the generated class prefixes (``sp-f-``, ``sp-b-``),
            unique to each sheet on the workbook page; empty for one sheet.
    """

    def __init__(
        self, sheet: RenderedSheet, backend: HTMLBackend, *, scope: str = "sp", tag: str = ""
    ) -> None:
        if sheet.mode != "excel":
            raise ValueError("the HTML backend draws a sheet laid out in excel mode")
        self.sheet = sheet
        self.backend = backend
        self.scope = scope
        self.tag = tag
        self.book: _links.BookLinks | None = None
        """The workbook page's other sheets, for links to them; ``None`` for one sheet."""
        self.ids: dict[tuple[int, int], str] = {}
        """Grid position of a cell a link goes to → its id."""
        self._grid: tuple[_Spans, Holder] | None = None
        self._links: _links.Links | None = None
        self.ox, self.oy = _headings.offsets(sheet, backend)
        self.gx, self.gy = _outline.offsets(sheet, backend)  # px, the outline gutters
        self.outline_rows = _outline.row_marks(sheet, backend)
        self.normal = sheet.sheet.normal_font
        self.normal_face = _text._face(self.normal)
        self.width = self.x(sheet.width_pt) + 1  # px, the wrapper's width: the table's
        self.fills = _patterns.Fills(self)
        self.borders = _borders.Borders(self)
        model = sheet.sheet
        self.grid_color = model.grid_color if model.show_gridlines else None

    def css(self) -> str:
        """The page's own stylesheet: the classes its fill, border and grid colors use."""
        return _patterns.layer_css(self.fills, self.borders.css()) + _grid_css(self.grid_color)

    def x(self, points: float) -> int:
        return self.gx + _px(self.ox + points)

    def y(self, points: float) -> int:
        return self.gy + _px(self.oy + points)

    def write(self, out: TextIO) -> None:
        sheet = self.sheet
        classes = "sheetpeek" if sheet.sheet.show_gridlines else "sheetpeek sp-nogrid"
        if self.grid_color:
            classes += f" sp-g-{self.grid_color}"
        out.write(f'<div class="{classes}" data-color-scheme="{self.backend.color_scheme}">')
        self.fills.write(out)
        self.borders.write(out)
        out.write(f'<table style="width:{self.width}px;font:{_text._font(self.normal)}">\n')
        out.write("<colgroup>")
        _outline.write_column(self, out)
        _headings.write_column(self, out)
        for c in sheet.columns:
            width = self.x((c.x_pt or 0) + (c.width_pt or 0)) - self.x(c.x_pt or 0)
            out.write(f'<col style="width:{width}px">')
        out.write("</colgroup>\n")
        if self.gy or self.backend.column_headings:
            out.write("<thead>")
            _outline.write_header(self, out)
            _headings.write_header(self, out, _outline.lead(self))
            out.write("</thead>\n")
        out.write("<tbody>\n")
        self.write_rows(out)
        out.write("</tbody></table></div>\n")

    def grid(self) -> tuple[_Spans, Holder]:
        """Each merge's grid span, and sheet position → grid position of the cell drawing it."""
        if self._grid is not None:
            return self._grid
        grid = [row.cells for row in self.sheet.rows]
        # Each merge is one cell at its first visible position, spanning its visible area.
        spans: _Spans = {}
        for i, cells in enumerate(grid):
            for p, cell in enumerate(cells):
                if cell.merge is not None:
                    first = spans.get(cell.merge, (i, p, i, p))
                    spans[cell.merge] = (first[0], first[1], i, p)
        merged: dict[tuple[int, int], tuple[int, int]] = {}
        for cells in grid:
            for cell in cells:
                if cell.merge is not None:
                    top, left = spans[cell.merge][:2]
                    merged[(cell.row, cell.column)] = (top, left)
                    anchor = (cell.merge.min_row, cell.merge.min_col)
                    merged.setdefault(anchor, (top, left))
        holder = Holder(grid, merged)
        self._grid = spans, holder
        return self._grid

    @property
    def links(self) -> _links.Links:
        """The page's links, found when first asked.

        On the workbook page every page's are found before any sheet is written, since a link
        gives an id to a cell of another sheet.
        """
        if self._links is None:
            self._links = _links.Links(self, self.grid()[1])
        return self._links

    def write_rows(self, out: TextIO) -> None:
        sheet = self.sheet
        spans, holder = self.grid()
        self.links  # noqa: B018 - found before the cells are written, for their ids
        runs: dict[tuple[int, int], list[TextRun]] = {}
        for i, row in enumerate(sheet.rows):
            for run in row.runs:
                runs.setdefault(holder.get(run.source, (i, 0)), []).append(run)
        for i, row in enumerate(sheet.rows):
            y0, y1 = row.y_pt or 0, (row.y_pt or 0) + (row.height_pt or 0)
            out.write(f'<tr style="height:{self.y(y1) - self.y(y0)}px">')
            _outline.write_row(self, row, out)
            _headings.write_row(self, row, out)
            for p, cell in enumerate(row.cells):
                colspan = rowspan = 1
                if cell.merge is not None:
                    top, left, bottom, right = spans[cell.merge]
                    if (i, p) != (top, left):
                        continue  # covered by its merge
                    colspan, rowspan = right - left + 1, bottom - top + 1
                out.write(self.cell(i, p, cell, colspan, rowspan, runs.get((i, p), ())))
            out.write("</tr>\n")

    def cell(
        self,
        i: int,
        p: int,
        cell: RenderedCell,
        colspan: int,
        rowspan: int,
        runs: list[TextRun] | tuple[()],
    ) -> str:
        rows = self.sheet.rows
        style = []
        solid = _patterns.solid
        # Most cells have no fill: ask the fill layer only about cells that have one.
        left = rows[i].cells[p - 1] if p else None
        left_fill = solid(self, left) if left is not None and left.fill is not None else None
        above = rows[i - 1].cells[p] if i else None
        above_fill = solid(self, above) if above is not None and above.fill is not None else None
        fill = solid(self, cell) if cell.fill is not None else None
        if fill:
            # A fill covers the cell's own gridlines and its far ones (captured).
            style += [f"background:#{fill}", f"border-color:#{fill}"]
        else:
            if left_fill:
                style.append(f"border-left-color:#{left_fill}")
            if above_fill:
                style.append(f"border-top-color:#{above_fill}")
        attrs = ""
        found = self.ids.get((i, p)) if self.ids else None
        if found is not None:
            attrs += f' id="{found}"'
        if p and rows[i].crossed[p - 1]:
            attrs += ' class="sp-crossed"'
        if style:
            attrs += f' style="{";".join(style)}"'
        if colspan > 1:
            attrs += f' colspan="{colspan}"'
        if rowspan > 1:
            attrs += f' rowspan="{rowspan}"'
        return f"<td{attrs}>{self.text(runs) if runs else ''}</td>"

    def text(self, runs: list[TextRun] | tuple[()]) -> str:
        """A cell's runs, each linked cell's inside its anchor."""
        if not runs:
            return ""
        opening = self.links.opening
        if not opening:
            return "".join(_text.run(self, run) for run in runs)
        parts = []
        source = None
        for run in runs:
            if run.source != source:
                if source in opening:
                    parts.append("</a>")
                source = run.source
                if source in opening:
                    parts.append(opening[source])
            parts.append(_text.run(self, run))
        if source in opening:
            parts.append("</a>")
        return "".join(parts)
