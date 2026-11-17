"""Outlines (grouped rows and columns): what the file states, and the symbols a layout draws.

A row or column belongs to the groups of every level up to its own ``outline_level``. Excel draws
the outline beside the headings: a lane per level, a bar along each group's detail rows ending
in the minus or plus button on its summary row, and a dot on each detail row in the lane after
its level.
``ARCHITECTURE.md`` §4.11 has the rules.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, TypeAlias

MAX_LEVEL = 7
"""The deepest outline level Excel offers; a level in a file is read modulo 8, as Excel does."""


@dataclass(frozen=True, slots=True)
class Outline:
    """A sheet's outline settings, present when any row or column has an outline level.

    The depths are the deepest level found in the whole sheet, not only in the range read, as
    Excel derives them; the file's ``outlineLevelRow`` and ``outlineLevelCol`` are not needed.
    """

    summary_below: bool = True
    """``<outlinePr summaryBelow>``: a group's summary row is below its detail rows, else above."""
    summary_right: bool = True
    """``<outlinePr summaryRight>``: a group's summary column is right of its details, else left."""
    show_symbols: bool = True
    """``showOutlineSymbols`` of the first ``<sheetView>``, else of ``<outlinePr>``: whether
    Excel draws the outline symbols."""
    row_depth: int = 0
    """The deepest row outline level in the sheet, 0 to 7."""
    column_depth: int = 0
    """The deepest column outline level in the sheet, 0 to 7."""
    rows_beyond: tuple[int, int] = (0, 0)
    """The outline levels of the row just above the range read and the row just below it, so
    that a group running past the range's edge is known as such."""
    columns_beyond: tuple[int, int] = (0, 0)
    """The outline levels of the columns just left and just right of the range read."""


OutlineMarkKind: TypeAlias = Literal["start", "bar", "open", "closed", "dot"]
"""What a lane holds on one row (or column):

- ``start``: the end of a group's bar away from its summary, with the tick toward the headings;
- ``bar``: the bar running through;
- ``open``: the summary's minus button, the bar running into it from the detail side;
- ``closed``: the summary's plus button of a collapsed group, with no bar;
- ``dot``: a detail row's dot, in the lane after its own level.
"""


@dataclass(frozen=True, slots=True)
class OutlineMark:
    """One outline symbol: what lane ``lane`` holds on row (or column) ``index``."""

    index: int
    """The sheet row or column."""
    lane: int
    """1 for the outermost groups; a level-``n`` group's bar and button are in lane ``n``, a
    level-``n`` row's dot in lane ``n + 1``."""
    kind: OutlineMarkKind


@dataclass(frozen=True, slots=True)
class RenderedOutline:
    """The outline symbols of a laid-out sheet, for the rows and columns laid out.

    Present when the sheet has an outline and the file shows its symbols. Marks are sorted by
    index, then lane; a row or column without a mark draws an empty gutter.
    """

    rows: tuple[OutlineMark, ...]
    columns: tuple[OutlineMark, ...]
    row_lanes: int
    """Lanes of the row gutter, the row depth plus one; 0 when no row has a level."""
    column_lanes: int
    """Lanes of the column gutter, the column depth plus one; 0 when no column has a level."""
    summary_below: bool = True
    summary_right: bool = True
