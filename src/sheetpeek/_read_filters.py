"""Reading filters: the sheet's ``autoFilter`` and each table's, with their sorts (§4.19).

xlio reads the elements; this keeps what the buttons need. A column's ``colId`` counts from the
filter range's first column; one that falls outside the range is left out. A sort condition's
column is its ``ref``'s first. A sort's ``columnSort`` (left to right) draws no arrow
(unverified). The sheet's ``sortState`` counts inside its ``autoFilter`` (captured) or beside it
(unverified, as a table's beside its ``autoFilter`` is captured).
"""

from __future__ import annotations

from typing import Any

import xlio

from sheetpeek._filter_model import AutoFilter, FilterColumn, SortCondition
from sheetpeek._model import CellRange

_SORT_BY = frozenset({"value", "cellColor", "fontColor", "icon"})


def read_autofilters(source: xlio.Sheet, ref: CellRange | None) -> tuple[AutoFilter, ...]:
    """The sheet's filter, then each table's in file order, that touch ``ref``."""
    if ref is None:
        return ()
    found: list[AutoFilter] = []
    sheet_filter = _filter(source.autofilter, None, source.sort_state)
    if sheet_filter is not None:
        found.append(sheet_filter)
    for table in source.tables:
        if table.header_row_count:
            table_filter = _filter(table.autofilter, table.name, table.sort_state)
            if table_filter is not None:
                found.append(table_filter)
    return tuple(f for f in found if f.ref.intersects(ref))


def _range(text: str | None) -> CellRange | None:
    if not text:
        return None
    try:
        return CellRange.parse(text)
    except ValueError:
        return None


def _filter(
    autofilter: xlio.AutoFilter | None, table: str | None, beside: xlio.SortState | None
) -> AutoFilter | None:
    if autofilter is None:
        return None
    ref = _range(autofilter.ref)
    if ref is None:
        return None
    columns: dict[int, FilterColumn] = {}
    for column in autofilter.columns:
        index = ref.min_col + column.column_id
        if column.column_id >= 0 and index <= ref.max_col and index not in columns:
            columns[index] = FilterColumn(
                index,
                button=not column.hidden_button and column.show_button,
                filtered=_filtered(column),
            )
    sort = autofilter.sort_state or beside
    return AutoFilter(ref, table, tuple(columns.values()), _conditions(sort, ref))


def _filtered(column: Any) -> bool:
    return bool(
        column.values
        or column.blank
        or column.date_groups
        or column.custom
        or column.top10 is not None
        or column.dynamic is not None
        or column.color is not None
        or column.icon is not None
    )


def _conditions(sort: xlio.SortState | None, ref: CellRange) -> tuple[SortCondition, ...]:
    if sort is None or sort.column_sort:
        return ()
    found: list[SortCondition] = []
    for condition in sort.conditions:
        where = _range(condition.ref)
        if where is None or not ref.min_col <= where.min_col <= ref.max_col:
            continue
        by = condition.sort_by if condition.sort_by in _SORT_BY else "value"
        found.append(SortCondition(where.min_col, condition.descending, by))
    return tuple(found)
