"""Reading outlines: levels, collapsed flags and the outline settings.

xlio reads the levels and collapsed flags of rows and columns, as Excel takes them (a level is
kept modulo 8; one that is not decimal digits is none), ``summaryBelow`` and ``summaryRight``,
and ``showOutlineSymbols`` with whether the file states it (:func:`with_levels`). The symbols
show as the first ``<sheetView>`` says when it states it, else as ``<outlinePr>`` says. All is
kept in ``Structure.extra["outline"]``; :func:`read_outline` puts it into the model's rows and
columns and builds the sheet's :class:`~sheetpeek.Outline`.
"""

from __future__ import annotations

import dataclasses
from dataclasses import dataclass, field

import xlio

from sheetpeek._model import CellRange, Column, Row
from sheetpeek._notices import Collector
from sheetpeek._outline_model import Outline
from sheetpeek._structure import Structure

KEY = "outline"


@dataclass(slots=True)
class Found:
    """What the sheet's outline holds."""

    rows: dict[int, int] = field(default_factory=dict)
    """Row → outline level, for rows with a level."""
    collapsed_rows: set[int] = field(default_factory=set)
    columns: list[tuple[int, int, int, bool]] = field(default_factory=list)
    """(first, last, level, collapsed) of ``<col>`` elements with a level or collapsed."""
    summary_below: bool = True
    summary_right: bool = True
    symbols_view: bool | None = None
    symbols_properties: bool | None = None


def with_levels(structure: Structure, sheet: xlio.Sheet) -> Structure:
    """``structure`` with the sheet's outline from xlio: levels, collapsed flags, settings."""
    found = Found()
    for index, row in sheet.rows.items():
        if row.outline_level:
            found.rows[index] = row.outline_level
        if row.collapsed:
            found.collapsed_rows.add(index)
    found.columns = [
        (column.min, column.max, column.outline_level, column.collapsed)
        for column in sheet.columns
        if column.outline_level or column.collapsed
    ]
    properties = sheet.properties
    found.summary_below = properties.summary_below
    found.summary_right = properties.summary_right
    if properties.show_outline_symbols_text is not None:
        found.symbols_properties = properties.show_outline_symbols
    views = sheet.views
    if views and views[0].show_outline_symbols_text is not None:
        found.symbols_view = views[0].show_outline_symbols
    return dataclasses.replace(structure, extra={**structure.extra, KEY: found})


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
