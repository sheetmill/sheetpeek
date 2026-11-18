"""The filter-button pass: which header cells draw a button, and its glyph (§4.19).

Excel draws a button on each column of a filter range's first row, merged and empty header
cells included, unless the column's ``hiddenButton`` or ``showButton="0"`` hides it; none on a
hidden column or row. A range read that cuts the header row keeps the buttons of the header
cells inside it; one without the header row draws none, as ``range_edge`` cuts other things.
"""

from __future__ import annotations

import dataclasses
from typing import TYPE_CHECKING

from sheetpeek._filter_model import AutoFilter, FilterButtonState

if TYPE_CHECKING:
    from sheetpeek._layout import Draft
    from sheetpeek._rendered import RenderedSheet


def buttons_pass(draft: Draft) -> None:
    """Set ``RenderedCell.button`` on the header cells that draw one."""
    from sheetpeek._layout import _column_hidden, _row_hidden

    sheet = draft.sheet
    hidden_rows = {r.index for r in draft.kept_rows if _row_hidden(r, sheet)}
    hidden_columns = {c.index for c in draft.kept_columns if _column_hidden(c)}
    wanted: dict[tuple[int, int], FilterButtonState] = {}
    for autofilter in sheet.autofilters:
        row = autofilter.ref.min_row
        if row in hidden_rows:
            continue
        for column in range(autofilter.ref.min_col, autofilter.ref.max_col + 1):
            settings = autofilter.column(column)
            if settings.button and column not in hidden_columns and (row, column) not in wanted:
                wanted[(row, column)] = state(autofilter, column)
    if not wanted:
        return
    rows = list(draft.rows)
    for i, rendered_row in enumerate(rows):
        if not any((rendered_row.index, c.column) in wanted for c in rendered_row.cells):
            continue
        cells = tuple(
            dataclasses.replace(c, button=wanted[(c.row, c.column)])
            if (c.row, c.column) in wanted
            else c
            for c in rendered_row.cells
        )
        rows[i] = dataclasses.replace(rendered_row, cells=cells)
    draft.rows = tuple(rows)


def state(autofilter: AutoFilter, column: int) -> FilterButtonState:
    """The glyph of a column's button: filtered, sorted (each level draws its arrow), or both.

    A sort by cell color draws the ascending arrow (captured); one by font color or icon, and a
    descending sort by color, follow the order (unverified).
    """
    filtered = autofilter.column(column).filtered
    for condition in autofilter.sort_conditions:
        if condition.column == column:
            if condition.descending:
                return "filtered_descending" if filtered else "descending"
            return "filtered_ascending" if filtered else "ascending"
    return "filtered" if filtered else "plain"


def presence(rendered: RenderedSheet) -> int:
    """How many filter buttons the laid-out sheet draws, for the census."""
    return sum(1 for row in rendered.rows for cell in row.cells if cell.button is not None)
