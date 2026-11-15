"""Page layout as the file states it: the print area and manual page breaks (§4.18).

Excel's Normal view draws neither (captured); they are read so that a render can choose the
print area as its range and mark the breaks.
"""

from __future__ import annotations

import re
from collections.abc import Iterable

import xlio

from sheetpeek._model import CellRange

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
MAX_ROW = 1_048_576
MAX_COLUMN = 16_384
_COLUMN = re.compile(r"^\$?([A-Za-z]{1,3})$")
_ROW = re.compile(r"^\$?([0-9]{1,7})$")
_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")


def _column_number(letters: str) -> int:
    number = 0
    for char in letters.upper():
        number = number * 26 + ord(char) - 64
    return number


def _range(text: str) -> CellRange | None:
    """``$A$1:$C$5``, ``$A:$C``, ``$1:$5`` or a single cell as a range, else ``None``."""
    first, sep, last = text.partition(":")
    last = last if sep else first
    cells = [_CELL.match(first), _CELL.match(last)]
    if cells[0] and cells[1]:
        rows = [int(c.group(2)) for c in cells if c]
        columns = [_column_number(c.group(1)) for c in cells if c]
    elif sep and (a := _COLUMN.match(first)) and (b := _COLUMN.match(last)):
        rows, columns = [1, MAX_ROW], [_column_number(a.group(1)), _column_number(b.group(1))]
    elif sep and (a := _ROW.match(first)) and (b := _ROW.match(last)):
        rows, columns = [int(a.group(1)), int(b.group(1))], [1, MAX_COLUMN]
    else:
        return None
    if not all(1 <= r <= MAX_ROW for r in rows) or not all(1 <= c <= MAX_COLUMN for c in columns):
        return None
    return CellRange(min(rows), min(columns), max(rows), max(columns))


def print_areas(area: xlio.PrintArea | None) -> tuple[tuple[CellRange, ...], str | None]:
    """The print area as xlio reads it: its ranges in file order, and why a part was left out.

    xlio splits the name's parts, keeps the plain ranges on this sheet and names what it left
    out (``#REF!``, a formula, another sheet's range). A range sheetpeek cannot place is left
    out too and named.
    """
    if area is None:
        return (), None
    ranges: list[CellRange] = []
    problems = [area.problem] if area.problem else []
    for text in area.ranges:
        found = _range(text)
        if found is None:
            problems.append(f"a range sheetpeek cannot place ({text})")
        else:
            ranges.append(found)
    return tuple(ranges), "; ".join(problems) or None


def manual_breaks(breaks: Iterable[xlio.PageBreak], limit: int) -> tuple[int, ...]:
    """The manual page breaks: the rows a break falls below, or the columns it falls right of.

    ``id`` is the boundary (``10`` lies between rows 10 and 11). Only manual breaks are kept:
    automatic ones are Excel's own pagination, which it rewrites as it likes.
    """
    return tuple(sorted({b.id for b in breaks if b.manual and 1 <= b.id < limit}))
