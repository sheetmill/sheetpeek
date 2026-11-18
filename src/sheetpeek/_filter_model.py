"""Filter and sort buttons: the filters a sheet holds and the buttons Excel draws for them (§4.19).

The names follow Excel's and xlio's: an :class:`AutoFilter` is the sheet's filter or a table's,
with a :class:`FilterColumn` for each column that says anything and the :class:`SortCondition`
of its sort. Excel does not apply the criteria when it opens a file: the rows the file saved as
hidden are what shows, so the criteria matter only for the button they draw.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, TypeAlias

from sheetpeek._model import CellRange

SortBy: TypeAlias = Literal["value", "cellColor", "fontColor", "icon"]
FilterButtonState: TypeAlias = Literal[
    "plain", "filtered", "ascending", "descending", "filtered_ascending", "filtered_descending"
]


@dataclass(frozen=True, slots=True)
class FilterColumn:
    """A column of a filter range that the file says something about (``<filterColumn>``)."""

    column: int
    """The sheet column, 1-based (the file's ``colId`` counts from the range's first column)."""
    button: bool = True
    """Whether Excel draws the column's button: not with ``hiddenButton`` or ``showButton="0"``,
    which draw alike."""
    filtered: bool = False
    """Whether criteria of any kind are set; every kind draws the same button."""


@dataclass(frozen=True, slots=True)
class SortCondition:
    """One level of a filter's sort (``<sortCondition>``)."""

    column: int
    """The sheet column sorted, 1-based."""
    descending: bool = False
    sort_by: SortBy = "value"
    """What the column is sorted by: values, cell color, font color or icon."""


@dataclass(frozen=True, slots=True)
class AutoFilter:
    """A sheet's filter (``<autoFilter>``) or a table's, with the sort saved with it."""

    ref: CellRange
    """The filter range; its first row holds the buttons."""
    table: str | None = None
    """The table's name for a table's filter, ``None`` for the sheet's own."""
    columns: tuple[FilterColumn, ...] = ()
    """The columns the file says something about, in file order; the others draw a plain
    button."""
    sort_conditions: tuple[SortCondition, ...] = ()

    def column(self, column: int) -> FilterColumn:
        """The settings of a sheet column of the range: the file's, else a plain button's."""
        for found in self.columns:
            if found.column == column:
                return found
        return FilterColumn(column)
