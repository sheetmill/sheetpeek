"""Fixed-width plain text: every column a whole number of characters, as Excel lays it out."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TextIO

from sheetpeek._rendered import Mode, RenderedSheet


class TextBackend:
    """Fixed-width text in ``characters`` mode.

    Text spills across empty neighbors and over the separators between them, as it does over
    Excel's gridlines; numbers that do not fit show ``#``.

    Args:
        separator: One character drawn between two columns where no text crosses.
        column_headings: Add Excel's column letters above the sheet.
        row_headings: Add Excel's row numbers before each row.
    """

    name = "text"
    mode: Mode = "characters"

    def __init__(
        self, separator: str = " ", *, column_headings: bool = False, row_headings: bool = False
    ) -> None:
        if len(separator) != 1:
            raise ValueError("the separator must be one character")
        self.separator = separator
        self.column_headings = column_headings
        self.row_headings = row_headings

    def render(self, sheet: RenderedSheet, out: TextIO, /) -> None:
        """Write the sheet to ``out``, one line per line of each row."""
        for line in self.lines(sheet):
            out.write(line.rstrip() + "\n")

    def lines(self, sheet: RenderedSheet) -> Iterator[str]:
        """The output lines, trailing spaces included."""
        label = max((len(str(row.index)) for row in sheet.rows), default=1)
        prefix = " " * label + self.separator if self.row_headings else ""
        if self.column_headings and sheet.columns:
            letters = self.separator.join(
                c.letter.center(c.width or 0)[: c.width or 0] for c in sheet.columns
            )
            yield prefix + letters
        for row in sheet.rows:
            for line in range(row.height):
                parts = []
                for p, cell in enumerate(row.cells):
                    parts.append(cell.lines[line])
                    if p < len(row.gaps):
                        crossing = row.gaps[p].lines[line]
                        parts.append(crossing if crossing is not None else self.separator)
                text = "".join(parts)
                if self.row_headings:
                    number = str(row.index) if line == 0 else ""
                    text = number.rjust(label) + self.separator + text
                yield text
