"""The outline gutter in characters, for the ``rich`` and plain-text outputs.

One character per lane before the row headings, and one line per lane above the column
headings: ``│`` and ``─`` for bars, ``┌``, ``└`` or ``┐`` where a bar ends in its tick, ``-``
and ``+`` for the buttons, ``·`` for dots. The level buttons are left out.
"""

from __future__ import annotations

from sheetpeek._outline_model import OutlineMark, RenderedOutline
from sheetpeek._rendered import RenderedColumn

_ACROSS = {"bar": "│", "closed": "+", "open": "-", "dot": "·"}


class Gutter:
    """The gutter of a laid-out sheet's outline; empty for ``None``."""

    def __init__(self, outline: RenderedOutline | None) -> None:
        self.row_lanes = outline.row_lanes if outline else 0
        self.column_lanes = outline.column_lanes if outline else 0
        self.below = outline.summary_below if outline else True
        self.right = outline.summary_right if outline else True
        self.rows = _by_index(outline.rows) if outline else {}
        self.columns = _by_index(outline.columns) if outline else {}
        self.width = self.row_lanes + 1 if self.row_lanes else 0
        """Characters before each line: the lanes and a space."""

    def row(self, index: int, line: int, lines: int) -> str:
        """The gutter of one line of a row, ``width`` characters."""
        if not self.width:
            return ""
        chars = [" "] * self.row_lanes
        last = lines - 1
        for lane, kind in self.rows.get(index, {}).items():
            if kind == "start":
                tick = line == 0 if self.below else line == last
                char = ("┌" if self.below else "└") if tick else "│"
            elif kind == "open" and line:
                char = " " if self.below else "│"
            elif kind in ("closed", "dot") and line:
                char = " "
            else:
                char = _ACROSS[kind]
            chars[lane - 1] = char
        return "".join(chars) + " "

    def column_lines(self, columns: tuple[RenderedColumn, ...], separator: str) -> list[str]:
        """The column gutter's lines, one per lane, without the row gutter or headings before."""
        result = []
        for lane in range(1, self.column_lanes + 1):
            kinds = [self.columns.get(c.index, {}).get(lane) for c in columns]
            parts = []
            for p, column in enumerate(columns):
                parts.append(self._segment(kinds[p], column.width or 0))
                if p + 1 < len(columns):
                    joined = self._joins_right(kinds[p]) and self._joins_left(kinds[p + 1])
                    parts.append("─" if joined else separator)
            result.append("".join(parts))
        return result

    def _segment(self, kind: str | None, width: int) -> str:
        if not width:
            return ""
        middle = (width - 1) // 2
        rest = width - middle - 1
        if kind == "bar":
            return "─" * width
        if kind == "start":
            return "┌" + "─" * (width - 1) if self.right else "─" * (width - 1) + "┐"
        if kind == "open":
            before, after = ("─", " ") if self.right else (" ", "─")
            return before * middle + "-" + after * rest
        if kind in ("closed", "dot"):
            return " " * middle + ("+" if kind == "closed" else "·") + " " * rest
        return " " * width

    def _joins_right(self, kind: str | None) -> bool:
        """Whether a column's bar reaches its right edge."""
        return kind in ("bar", "start" if self.right else "open")

    def _joins_left(self, kind: str | None) -> bool:
        return kind in ("bar", "open" if self.right else "start")


def _by_index(marks: tuple[OutlineMark, ...]) -> dict[int, dict[int, str]]:
    found: dict[int, dict[int, str]] = {}
    for mark in marks:
        found.setdefault(mark.index, {})[mark.lane] = mark.kind
    return found
