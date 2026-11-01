"""Markdown tables (GitHub Flavored Markdown pipe tables)."""

from __future__ import annotations

from typing import Literal, TextIO

from sheetpeek._rendered import Mode, RenderedCell, RenderedSheet

_RULE = {"right": "---:", "center": ":---:"}


class MarkdownBackend:
    """A pipe table in ``none`` mode.

    A Markdown table needs a header row. By default the sheet's first visible row is the header;
    ``header="letters"`` uses the column letters instead and keeps every row in the body. Each
    column's alignment in the separator row follows the first body row's cell. Merges and
    widths cannot be shown: a merged range's text is in its anchor cell.

    Args:
        header: ``"first"`` (the first row) or ``"letters"`` (column letters).
        emphasis: Mark bold and italic cells with ``**`` and ``*``.
    """

    name = "markdown"
    mode: Mode = "none"

    def __init__(
        self, header: Literal["first", "letters"] = "first", *, emphasis: bool = False
    ) -> None:
        if header not in ("first", "letters"):
            raise ValueError("header must be 'first' or 'letters'")
        self.header = header
        self.emphasis = emphasis

    def render(self, sheet: RenderedSheet, out: TextIO, /) -> None:
        """Write the table; nothing for a sheet without visible cells."""
        rows = [list(row.cells) for row in sheet.rows]
        if not rows or not sheet.columns:
            return
        if self.header == "letters":
            head = [c.letter for c in sheet.columns]
            body = rows
        else:
            head = [self.cell_text(cell) for cell in rows[0]]
            body = rows[1:]
        out.write(_line(head))
        first = body[0] if body else rows[0]
        out.write(_line([_RULE.get(cell.align, "---") for cell in first]))
        for row in body:
            out.write(_line([self.cell_text(cell) for cell in row]))

    def cell_text(self, cell: RenderedCell) -> str:
        """A cell's text escaped for a table cell, with emphasis if asked for."""
        text = cell.lines[0].replace("\\", "\\\\").replace("|", "\\|")
        text = text.replace("\r\n", "\n").replace("\n", "<br>")
        if self.emphasis and text.strip() and cell.font is not None:
            if cell.font.italic:
                text = f"*{text}*"
            if cell.font.bold:
                text = f"**{text}**"
        return text


def _line(cells: list[str]) -> str:
    return "| " + " | ".join(cells) + " |\n"
