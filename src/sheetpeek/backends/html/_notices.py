"""The collapsed list of what the page does not show (``HTMLBackend(notices=True)``)."""

from __future__ import annotations

import html
from typing import TYPE_CHECKING, TextIO

from sheetpeek._model import one_line
from sheetpeek._rendered import RenderedSheet

if TYPE_CHECKING:
    from sheetpeek.backends.html import HTMLBackend


def write_notices(sheet: RenderedSheet, backend: HTMLBackend, out: TextIO) -> None:
    """The collapsed list of what the page does not show; nothing when everything is shown."""
    from sheetpeek._census import output_notices

    notices = sheet.notices + output_notices(sheet, backend.name, backend.support)
    if not notices:
        return
    out.write(
        '<details class="sp-notices" style="margin-top:8px;font:12px/1.4 system-ui,sans-serif">'
        f"<summary>Not shown in this view ({len(notices)})</summary>"
        "<ul>\n"
    )
    for notice in notices:
        out.write(f"<li>{html.escape(one_line(notice.message))}</li>\n")
    out.write("</ul></details>\n")
