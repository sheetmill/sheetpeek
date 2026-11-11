"""Markdown tables (GitHub Flavored Markdown pipe tables)."""

from __future__ import annotations

import re
from typing import Literal, TextIO
from urllib.parse import quote

from sheetpeek._hyperlink import external_url
from sheetpeek._notices import Support
from sheetpeek._rendered import Mode, RenderedCell, RenderedSheet
from sheetpeek.backends import _support

_RULE = {"right": "---:", "center": ":---:"}
# What makes text act as markup inside a table cell, which holds inline content only, so line-start
# markers such as # and > are inert. $ is GitHub's math, @ its mentions.
_MARKUP = re.compile(r"[\\`*_\[\]<&~|$@]")
# GFM's bare autolinks. An email address stays one in cmark-gfm whatever is escaped: the same text,
# as a mailto: link.
_AUTOLINK = re.compile(r"(?i)\b(?:(?:https?|ftp|mailto|xmpp)(?=:)|www(?=\.))")
# What a link destination keeps as it is; everything else is percent-encoded, so that no
# parenthesis, angle bracket, pipe, space or backslash ends it, and & is written as an entity.
_DESTINATION_SAFE = "/:?#@!$&'*+,;=%-._~"


class MarkdownBackend:
    """A pipe table in ``none`` mode.

    A Markdown table needs a header row. By default the sheet's first visible row is the header;
    ``header="letters"`` uses the column letters instead and keeps every row in the body. Each
    column's alignment in the separator row follows the first body row's cell. Merges and
    widths cannot be shown: a merged range's text is in its anchor cell. A link to a web page or
    an e-mail address is a Markdown link, ``[text](url)``; other links are plain text.

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

    @property
    def support(self) -> dict[str, Support]:
        """What this output shows of each feature the census finds (``backends/_support.py``)."""
        return _support.markdown(emphasis=self.emphasis, letters=self.header == "letters")

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
        # A lone carriage return is no line break in Excel (§4.1), and would end the table's row.
        text = cell.lines[0].replace("\r\n", "\n").replace("\r", "")
        text = _escape(text).replace("\n", "<br>")
        link = cell.cell.hyperlink
        if link is not None and text and cell.cell.kind != "empty":
            url = external_url(link)
            if url is not None:
                text = f"[{text}]({_destination(url)})"
        if self.emphasis and text.strip() and cell.font is not None:
            if cell.font.italic:
                text = f"*{text}*"
            if cell.font.bold:
                text = f"**{text}**"
        return text


def _line(cells: list[str]) -> str:
    return "| " + " | ".join(cells) + " |\n"


def _destination(url: str) -> str:
    """A URL as a link destination that a CommonMark or GFM renderer takes whole."""
    return quote(url, safe=_DESTINATION_SAFE).replace("&", "&amp;")


def _escape(text: str) -> str:
    """Text that a CommonMark or GFM renderer shows literally, whether it allows HTML or not."""
    escaped = _MARKUP.sub(lambda m: "\\" + m[0], text)
    return _AUTOLINK.sub(lambda m: f"{m[0]}\\", escaped)
