"""Fixed-width plain text: every column a whole number of characters, as Excel lays it out."""

from __future__ import annotations

from collections.abc import Iterator
from typing import Literal, TextIO

from sheetpeek._notices import Support
from sheetpeek._rendered import Mode, RenderedSheet
from sheetpeek.backends import _support
from sheetpeek.backends._buttons_text import with_button
from sheetpeek.backends._outline_text import Gutter

Buttons = Literal["none", "static"]


class TextBackend:
    """Fixed-width text in ``characters`` mode.

    Text spills across empty neighbors and over the separators between them, as it does over
    Excel's gridlines; numbers that do not fit show ``#``.

    Args:
        separator: One character drawn between two columns where no text crosses.
        column_headings: Add Excel's column letters above the sheet.
        row_headings: Add Excel's row numbers before each row.
        buttons: ``"static"`` (the default) draws the outline gutter in characters before
            the row headings and above the column headings, as Excel draws it beside its
            headings, and a ``▾`` in the last character of each filter's header cells;
            ``"none"`` leaves both out. A sheet's outline is drawn only when the file shows
            outline symbols.
    """

    name = "text"
    mode: Mode = "characters"

    def __init__(
        self,
        separator: str = " ",
        *,
        column_headings: bool = False,
        row_headings: bool = False,
        buttons: Buttons = "static",
    ) -> None:
        if len(separator) != 1:
            raise ValueError("the separator must be one character")
        self.separator = separator
        self.column_headings = column_headings
        self.row_headings = row_headings
        if buttons not in ("none", "static"):
            raise ValueError(f"buttons must be 'none' or 'static', not {buttons!r}")
        self.buttons = buttons

    @property
    def support(self) -> dict[str, Support]:
        """What this output shows of each feature the census finds (``backends/_support.py``)."""
        return _support.text(
            headings=self.column_headings or self.row_headings, buttons=self.buttons
        )

    def render(self, sheet: RenderedSheet, out: TextIO, /) -> None:
        """Write the sheet to ``out``, one line per line of each row."""
        for line in self.lines(sheet):
            out.write(line.rstrip() + "\n")

    def lines(self, sheet: RenderedSheet) -> Iterator[str]:
        """The output lines, trailing spaces included."""
        label = max((len(str(row.index)) for row in sheet.rows), default=1)
        gutter = Gutter(sheet.outline if self.buttons != "none" else None)
        prefix = " " * gutter.width + (" " * label + self.separator if self.row_headings else "")
        if sheet.columns:
            for band in gutter.column_lines(sheet.columns, self.separator):
                yield prefix + band
        if self.column_headings and sheet.columns:
            letters = self.separator.join(
                c.letter.center(c.width or 0)[: c.width or 0] for c in sheet.columns
            )
            yield prefix + letters
        buttons = self.buttons != "none"
        for row in sheet.rows:
            for line in range(row.height):
                parts = []
                for p, cell in enumerate(row.cells):
                    chars = cell.lines[line]
                    if buttons and cell.button is not None:
                        chars = with_button(cell, line, row.height, chars)
                    parts.append(chars)
                    if p < len(row.gaps):
                        crossing = row.gaps[p].lines[line]
                        parts.append(crossing if crossing is not None else self.separator)
                text = "".join(parts)
                if self.row_headings:
                    number = str(row.index) if line == 0 else ""
                    text = number.rjust(label) + self.separator + text
                yield gutter.row(row.index, line, row.height) + text
