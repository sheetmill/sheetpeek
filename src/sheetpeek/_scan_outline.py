"""Reading outlines: the scan's handlers for levels, collapsed flags and the outline settings.

The handlers keep what they find in ``Scan.extra["outline"]``; :func:`read_outline` puts it into
the model's rows and columns and builds the sheet's :class:`~sheetpeek.Outline`.
"""

from __future__ import annotations

import dataclasses
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

from sheetpeek._model import CellRange, Column, Row
from sheetpeek._notices import Collector
from sheetpeek._outline_model import MAX_LEVEL, Outline
from sheetpeek._structure import _COL, _NS, _ROW, _SHEET_VIEW, Scan, Structure, _flag, handler

KEY = "outline"
_OUTLINE_PR = f"{_NS}outlinePr"


@dataclass(slots=True)
class Found:
    """What the handlers found in one sheet."""

    rows: dict[int, int] = field(default_factory=dict)
    """Row → outline level, for rows with a level."""
    collapsed_rows: set[int] = field(default_factory=set)
    columns: list[tuple[int, int, int, bool]] = field(default_factory=list)
    """(first, last, level, collapsed) of ``<col>`` elements with a level or collapsed."""
    summary_below: bool = True
    summary_right: bool = True
    symbols_view: bool | None = None
    symbols_properties: bool | None = None
    seen_view: bool = False
    clamped: int = 0
    """Rows and column elements whose level is above 7, read as 7."""
    unreadable: int = 0
    """Rows and column elements whose level is not a whole number from 0 up, read as 0."""


def _found(state: Scan) -> Found:
    found = state.extra.get(KEY)
    if not isinstance(found, Found):
        found = state.extra[KEY] = Found()
    return found


def _level(state: Scan, element: ET.Element) -> int:
    """An element's ``outlineLevel``: 7 at most, and 0 when it is not a whole number."""
    text = element.get("outlineLevel")
    if text is None:
        return 0
    try:
        level = int(text)
    except ValueError:  # not a number, or one of thousands of digits
        level = -1
    if level < 0:
        _found(state).unreadable += 1
        return 0
    if level > MAX_LEVEL:
        _found(state).clamped += 1
        return MAX_LEVEL
    return level


@handler(_ROW)
def _row(state: Scan, element: ET.Element) -> None:
    """A row's level and collapsed flag; the core handler has set ``state.row_index``."""
    level = _level(state, element)
    collapsed = _flag(element.get("collapsed"))
    if level or collapsed:
        found = _found(state)
        if level:
            found.rows[state.row_index] = level
        if collapsed:
            found.collapsed_rows.add(state.row_index)


@handler(_COL)
def _col(state: Scan, element: ET.Element) -> None:
    level = _level(state, element)
    collapsed = _flag(element.get("collapsed"))
    if level or collapsed:
        first = int(element.get("min", "1"))
        last = int(element.get("max", element.get("min", "1")))
        _found(state).columns.append((first, last, level, collapsed))


@handler(_OUTLINE_PR)
def _properties(state: Scan, element: ET.Element) -> None:
    found = _found(state)
    found.summary_below = _flag(element.get("summaryBelow"), default=True)
    found.summary_right = _flag(element.get("summaryRight"), default=True)
    if (symbols := element.get("showOutlineSymbols")) is not None:
        found.symbols_properties = _flag(symbols)


@handler(_SHEET_VIEW)
def _view(state: Scan, element: ET.Element) -> None:
    """The first sheet view's ``showOutlineSymbols``, which wins over ``outlinePr``'s."""
    found = _found(state)
    if found.seen_view:
        return
    found.seen_view = True
    if (symbols := element.get("showOutlineSymbols")) is not None:
        found.symbols_view = _flag(symbols)


def read_outline(
    structure: Structure,
    ref: CellRange | None,
    rows: tuple[Row, ...],
    columns: tuple[Column, ...],
    notes: Collector,
) -> tuple[tuple[Row, ...], tuple[Column, ...], Outline | None]:
    """The rows and columns with their levels, and the sheet's outline, or ``None``."""
    found = structure.extra.get(KEY)
    if not isinstance(found, Found):
        return rows, columns, None
    _report(found, notes)
    rows = tuple(
        dataclasses.replace(
            row,
            outline_level=found.rows.get(row.index, 0),
            collapsed=row.index in found.collapsed_rows,
        )
        if row.index in found.rows or row.index in found.collapsed_rows
        else row
        for row in rows
    )
    by_column: dict[int, tuple[int, bool]] = {}
    if ref is not None:
        for first, last, level, collapsed in found.columns:
            for index in range(max(first, ref.min_col - 1), min(last, ref.max_col + 1) + 1):
                by_column[index] = (level, collapsed)  # a later element wins, as for widths
    columns = tuple(
        dataclasses.replace(column, outline_level=level, collapsed=collapsed)
        if (level, collapsed) != (0, False)
        else column
        for column in columns
        for level, collapsed in (by_column.get(column.index, (0, False)),)
    )
    row_depth = max(found.rows.values(), default=0)
    column_depth = max((level for _, _, level, _ in found.columns), default=0)
    if not row_depth and not column_depth:
        return rows, columns, None
    symbols = found.symbols_view
    if symbols is None:
        symbols = found.symbols_properties
    beyond_rows = beyond_columns = (0, 0)
    if ref is not None:
        beyond_rows = (found.rows.get(ref.min_row - 1, 0), found.rows.get(ref.max_row + 1, 0))
        beyond_columns = (
            by_column.get(ref.min_col - 1, (0, False))[0],
            by_column.get(ref.max_col + 1, (0, False))[0],
        )
    return (
        rows,
        columns,
        Outline(
            summary_below=found.summary_below,
            summary_right=found.summary_right,
            show_symbols=True if symbols is None else symbols,
            row_depth=row_depth,
            column_depth=column_depth,
            rows_beyond=beyond_rows,
            columns_beyond=beyond_columns,
        ),
    )


def _report(found: Found, notes: Collector) -> None:
    """Levels outside 0 to 7. What Excel does with them is not captured (unverified)."""
    for count, detail in (
        (found.clamped, "levels above 7 read as 7"),
        (found.unreadable, "levels that are not whole numbers from 0 up read as 0"),
    ):
        if count:
            notes.add("outline.level", "approximated", "file", count=count, detail=detail)
