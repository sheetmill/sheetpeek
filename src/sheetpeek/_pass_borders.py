"""The borders pass: one resolved side per edge of the laid-out grid, and the diagonals.

The rules (``ARCHITECTURE.md`` §4.12, captures on Excel 16.114 only): where two cells' sides
meet, one is drawn whole, by :func:`~sheetpeek._fill_border_model.winner`; each cell keeps its own
sides, also inside a merge, whose inner sides are dropped; an unmerged angled cell's own sides slant
with its text (§4.17) and leave the upright grid; a vertical side that spilled text crosses is not
drawn, over the whole row; hidden rows and columns are not in the grid, so their sides do not
compete. A merge's diagonals are its anchor's, across the whole merge. The pass runs in ``excel``
mode, after the cells, since spill decides which vertical sides are drawn.
"""

from __future__ import annotations

import math
from typing import TYPE_CHECKING

from sheetpeek._fill_border_model import Diagonal, RenderedBorders, Side, SlantedBorders, winner

if TYPE_CHECKING:
    from sheetpeek._layout import Draft
    from sheetpeek._model import CellRange, CellStyle
    from sheetpeek._rendered import RenderedSheet


def borders_pass(draft: Draft) -> None:
    """Resolve the sheet's borders into ``draft.extra["borders"]`` (``RenderedSheet.borders``)."""
    if draft.mode != "excel":
        return
    rows = draft.rows
    if not any(c.cell.style.borders is not None for row in rows for c in row.cells):
        return
    grid = [row.cells for row in rows]
    count, width = len(grid), len(draft.columns)

    def merge_at(i: int, p: int) -> CellRange | None:
        return grid[i][p].merge if 0 <= i < count and 0 <= p < width else None

    slanted = _slanted(draft)
    at = {(s.row, s.column) for s in slanted}

    def own(i: int, p: int, side: int) -> Side | None:
        """Side ``side`` (left, right, top, bottom) of grid cell (i, p), if it is in the grid."""
        if not (0 <= i < count and 0 <= p < width) or (i, p) in at:
            return None
        borders = grid[i][p].cell.style.borders
        return borders.sides[side] if borders is not None else None

    horizontal = []
    for k in range(count + 1):
        line: list[Side | None] = []
        for p in range(width):
            merge = merge_at(k, p)
            if merge is not None and merge == merge_at(k - 1, p):
                line.append(None)  # inside a merge
            else:
                line.append(winner(own(k - 1, p, 3), own(k, p, 2)))
        horizontal.append(tuple(line))
    vertical = []
    for i, row in enumerate(rows):
        line = []
        for q in range(width + 1):
            merge = merge_at(i, q)
            if 0 < q < width and row.crossed and row.crossed[q - 1]:
                line.append(None)  # crossed by spilled text, or inside a merge
            elif merge is not None and merge == merge_at(i, q - 1):
                line.append(None)
            else:
                line.append(winner(own(i, q - 1, 1), own(i, q, 0)))
        vertical.append(tuple(line))
    draft.extra["borders"] = RenderedBorders(
        tuple(horizontal), tuple(vertical), _diagonals(draft), slanted
    )


def slants(style: CellStyle) -> bool:
    """Whether an unmerged cell of ``style`` draws its side borders slanted (§4.17).

    Captured at 15, 45 and -45 degrees: any side border slants the box, and at 90 and -90 the
    sides stay upright.
    """
    borders = style.borders
    return style.rotation not in (0, 90, -90) and borders is not None and any(borders.sides)


def slant_shift(style: CellStyle, height: float) -> float:
    """How far right of the cell's top edge the top of the slanted box lies, in points."""
    return height / math.tan(math.radians(style.rotation))


def _slanted(draft: Draft) -> tuple[SlantedBorders, ...]:
    """The side borders of unmerged angled cells, slanted with the text (:func:`slants`).

    A merge's borders stay upright (unverified; reported as ``border.slanted``).
    """
    found = []
    for i, row in enumerate(draft.rows):
        for p, rendered in enumerate(row.cells):
            style = rendered.cell.style
            borders = style.borders
            if rendered.merge is not None or borders is None or not slants(style):
                continue
            shift = slant_shift(style, row.height_pt or 0.0)
            left, right, top, bottom = borders.sides
            found.append(SlantedBorders(i, p, style.rotation, shift, left, right, top, bottom))
    return tuple(found)


def _diagonals(draft: Draft) -> tuple[Diagonal, ...]:
    """Each unmerged cell's diagonals, and each merge's anchor's across its laid-out area."""
    sheet = draft.sheet
    found = []
    areas: dict[CellRange, list[int]] = {}
    for i, row in enumerate(draft.rows):
        for p, rendered in enumerate(row.cells):
            merge = rendered.merge
            if merge is not None:
                area = areas.setdefault(merge, [i, p, i, p])
                area[2], area[3] = i, p
                continue
            borders = rendered.cell.style.borders
            if borders is not None and borders.diagonal is not None:
                found.append(
                    Diagonal(
                        i, p, i, p, borders.diagonal, borders.diagonal_up, borders.diagonal_down
                    )
                )
    ref = sheet.ref
    for merge, (top, left, bottom, right) in areas.items():
        if ref is None or (merge.min_row, merge.min_col) not in ref:
            continue
        borders = sheet.cell(merge.min_row, merge.min_col).style.borders
        if borders is not None and borders.diagonal is not None:
            found.append(
                Diagonal(
                    top,
                    left,
                    bottom,
                    right,
                    borders.diagonal,
                    borders.diagonal_up,
                    borders.diagonal_down,
                )
            )
    found.sort(key=lambda d: (d.top, d.left))
    return tuple(found)


def rendered_borders(sheet: RenderedSheet) -> RenderedBorders | None:
    """The borders the layout resolved, or ``None`` (no border, or not ``excel`` mode)."""
    return sheet.borders
