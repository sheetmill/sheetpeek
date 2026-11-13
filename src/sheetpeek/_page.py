"""Page layout as the file states it: the print area and manual page breaks (§4.18).

Excel's Normal view draws neither (captured); they are read so that a render can choose the
print area as its range and mark the breaks.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from collections.abc import Iterable

from sheetpeek._model import CellRange

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
MAX_ROW = 1_048_576
MAX_COLUMN = 16_384
_COLUMN = re.compile(r"^\$?([A-Za-z]{1,3})$")
_ROW = re.compile(r"^\$?([0-9]{1,7})$")
_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")


def _split_top_level(text: str) -> list[str]:
    """Split ``text`` at commas outside quoted sheet names and parentheses.

    ``''`` inside a quoted name is an escaped quote; parentheses hold a function's arguments.
    """
    parts: list[str] = []
    current: list[str] = []
    quoted, depth = False, 0
    for char in text:
        if char == "'":
            quoted = not quoted
        elif not quoted and char == "(":
            depth += 1
        elif not quoted and char == ")":
            depth = max(depth - 1, 0)
        if char == "," and not quoted and depth == 0:
            parts.append("".join(current))
            current = []
        else:
            current.append(char)
    parts.append("".join(current))
    return parts


def _sheet_and_range(part: str) -> tuple[str, str] | None:
    """``Sheet!range`` or ``'My ''Sheet'''!range`` as its sheet name and range, or ``None``."""
    part = part.strip()
    if part.startswith("'"):
        end = 1
        while end < len(part):
            if part[end] == "'":
                if part[end + 1 : end + 2] == "'":
                    end += 2
                    continue
                break
            end += 1
        name, rest = part[1:end].replace("''", "'"), part[end + 1 :]
    else:
        name, bang, rest = part.partition("!")
        if not bang:
            return None
        rest = "!" + rest
    if not rest.startswith("!"):
        return None
    return name, rest[1:]


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


def print_areas(text: str, sheet: str) -> tuple[tuple[CellRange, ...], str | None]:
    """A ``_xlnm.Print_Area`` value as its ranges in file order, and the problem if any.

    Each comma-separated part is ``sheet!range`` (Excel prints each on its own pages). A part
    with ``#REF!`` (its range was deleted), on another sheet, or that is not a plain range (a
    function or a name, a formula Excel evaluates) is left out and named in the problem; the
    usable parts are kept. Sheet names compare as Excel compares them, ignoring case.
    """
    areas: list[CellRange] = []
    problems: list[str] = []
    for part in _split_top_level(text):
        if "#REF!" in part.upper():
            problems.append("a deleted range (#REF!)")
            continue
        found = None if "(" in part.replace("''", "").split("'")[-1] else _sheet_and_range(part)
        if found is None:
            problems.append("a formula")
            continue
        name, ref = found
        if name.casefold() != sheet.casefold():
            problems.append("a range on another sheet")
            continue
        area = _range(ref)
        if area is None:
            problems.append("a formula")
            continue
        areas.append(area)
    problem = "; ".join(dict.fromkeys(problems)) or None
    return tuple(areas), problem


def page_breaks(children: Iterable[ET.Element]) -> tuple[tuple[int, ...], tuple[int, ...]]:
    """The manual page breaks: rows a break falls below and columns it falls right of.

    ``brk id`` counts from zero for the boundary above it, so ``id="10"`` lies between rows 10
    and 11, which is the row number the break falls below. Only ``man="1"`` breaks are kept:
    automatic ones are Excel's own pagination, which it rewrites as it likes.
    """
    found: dict[str, set[int]] = {"rowBreaks": set(), "colBreaks": set()}
    for child in children:
        kind = child.tag.rpartition("}")[2]
        if kind not in found:
            continue
        for brk in child.iter(f"{{{_MAIN}}}brk"):
            value = brk.get("id", "")
            if brk.get("man") in ("1", "true") and value.isascii() and value.isdigit():
                number = int(value)
                limit = MAX_ROW if kind == "rowBreaks" else MAX_COLUMN
                if 1 <= number < limit:
                    found[kind].add(number)
    return tuple(sorted(found["rowBreaks"])), tuple(sorted(found["colBreaks"]))
