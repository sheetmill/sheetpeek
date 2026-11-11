"""The grid: a fixed-layout table of cells with their fills, gridlines and merges."""

from __future__ import annotations

from typing import TYPE_CHECKING, TextIO

from sheetpeek._model import CellRange
from sheetpeek._rendered import RenderedCell, RenderedSheet, TextRun
from sheetpeek.backends.html import _headings, _links, _outline, _text

if TYPE_CHECKING:
    from sheetpeek.backends.html import HTMLBackend


def _px(points: float) -> int:
    """A sheet position in whole CSS pixels."""
    return round(points * 4 / 3)


class Page:
    """One sheet's markup."""

    def __init__(self, sheet: RenderedSheet, backend: HTMLBackend) -> None:
        if sheet.mode != "excel":
            raise ValueError("the HTML backend draws a sheet laid out in excel mode")
        self.sheet = sheet
        self.backend = backend
        self.ox, self.oy = _headings.offsets(sheet, backend)
        self.gx, self.gy = _outline.offsets(sheet, backend)  # px, the outline gutters
        self.outline_rows = _outline.row_marks(sheet, backend)
        self.normal = sheet.sheet.normal_font
        self.width = self.x(sheet.width_pt) + 1  # px, the wrapper's width: the table's

    def x(self, points: float) -> int:
        return self.gx + _px(self.ox + points)

    def y(self, points: float) -> int:
        return self.gy + _px(self.oy + points)

    def write(self, out: TextIO) -> None:
        sheet = self.sheet
        classes = "sheetpeek" if sheet.sheet.show_gridlines else "sheetpeek sp-nogrid"
        out.write(
            f'<div class="{classes}" data-color-scheme="{self.backend.color_scheme}">'
            f'<table style="width:{self.width}px;font:{_text._font(self.normal)}">\n'
        )
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

    def write_rows(self, out: TextIO) -> None:
        sheet = self.sheet
        grid = [row.cells for row in sheet.rows]
        # Each merge is one cell at its first visible position, spanning its visible area.
        spans: dict[CellRange, tuple[int, int, int, int]] = {}
        for i, cells in enumerate(grid):
            for p, cell in enumerate(cells):
                if cell.merge is not None:
                    first = spans.get(cell.merge, (i, p, i, p))
                    spans[cell.merge] = (first[0], first[1], i, p)
        holder: dict[tuple[int, int], tuple[int, int]] = {}  # sheet position -> grid position
        for i, cells in enumerate(grid):
            for p, cell in enumerate(cells):
                if cell.merge is not None:
                    top, left = spans[cell.merge][:2]
                    holder[(cell.row, cell.column)] = (top, left)
                    anchor = (cell.merge.min_row, cell.merge.min_col)
                    holder.setdefault(anchor, (top, left))
                else:
                    holder[(cell.row, cell.column)] = (i, p)
        self.links = _links.Links(self, holder)
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
        left = rows[i].cells[p - 1].fill if p else None
        above = rows[i - 1].cells[p].fill if i else None
        if cell.fill:
            # A fill covers the cell's own gridlines and its far ones (captured).
            style += [f"background:#{cell.fill}", f"border-color:#{cell.fill}"]
        else:
            if left:
                style.append(f"border-left-color:#{left}")
            if above:
                style.append(f"border-top-color:#{above}")
        attrs = ""
        if (i, p) in self.links.ids:
            attrs += f' id="{self.links.ids[(i, p)]}"'
        if p and rows[i].crossed[p - 1]:
            attrs += ' class="sp-crossed"'
        if style:
            attrs += f' style="{";".join(style)}"'
        if colspan > 1:
            attrs += f' colspan="{colspan}"'
        if rowspan > 1:
            attrs += f' rowspan="{rowspan}"'
        return f"<td{attrs}>{self.text(runs)}</td>"

    def text(self, runs: list[TextRun] | tuple[()]) -> str:
        """A cell's runs, each linked cell's inside its anchor."""
        opening = self.links.opening
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
