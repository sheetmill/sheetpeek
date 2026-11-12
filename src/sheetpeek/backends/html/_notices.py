"""The collapsed list of what the page does not show (``HTMLBackend(notices=True)``)."""

from __future__ import annotations

import html
from typing import TYPE_CHECKING, TextIO

from sheetpeek._model import one_line
from sheetpeek._notices import Notice
from sheetpeek._rendered import RenderedSheet

if TYPE_CHECKING:
    from sheetpeek.backends.html import HTMLBackend


def write_notices(
    sheet: RenderedSheet | None,
    backend: HTMLBackend,
    out: TextIO,
    notices: tuple[Notice, ...] | None = None,
) -> None:
    """The collapsed list of what the page does not show; nothing when everything is shown.

    ``notices`` replaces the sheet's own, which the workbook page works out itself.
    """
    from sheetpeek._census import output_notices

    if notices is None:
        assert sheet is not None
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
