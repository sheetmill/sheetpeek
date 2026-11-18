"""Terminal output through `rich`: the fixed-width layout with fonts, colors and fills.

``rich`` is an optional dependency (``sheetpeek[rich]``), imported only when this backend draws.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any, Literal, TextIO

from sheetpeek._hyperlink import external_url
from sheetpeek._notices import Support
from sheetpeek._rendered import Mode, RenderedSheet
from sheetpeek.backends import _support
from sheetpeek.backends._buttons_text import with_button
from sheetpeek.backends._outline_text import Gutter

if TYPE_CHECKING:
    from sheetpeek._model import Font

Buttons = Literal["none", "static"]


class RichBackend:
    """Styled terminal text in ``characters`` mode.

    Each column is drawn on its own cell's fill, and the text in it takes the font and color of
    the cell it comes from, so text spilled over a filled neighbor shows that neighbor's fill,
    as in Excel. A link to a web page or an e-mail address is a terminal hyperlink, which a
    terminal without them shows as plain text.

    Args:
        separator: One character drawn between two columns where no text crosses.
        column_headings: Add Excel's column letters above the sheet.
        row_headings: Add Excel's row numbers before each row.
        buttons: ``"static"`` (the default) draws the outline gutter in characters before
            the row headings and above the column headings, as Excel draws it beside its
            headings, and a ``▾`` in the last character of each filter's header cells;
            ``"none"`` leaves both out. A sheet's outline is drawn only when the file shows
            outline symbols.
        color_system: Passed to ``rich.console.Console`` (``"auto"``, ``"truecolor"``, ...).
        force_terminal: Passed to ``rich.console.Console``; ``None`` detects.
        local_links: Also link ``file:`` URLs and relative, UNC and drive-letter paths, which
            Excel follows. Off by default: such links point into the viewer's own files, they
            reveal the paths the workbook names, and a target starting with ``//`` names a
            network host.
    """

    name = "rich"
    mode: Mode = "characters"

    def __init__(
        self,
        separator: str = " ",
        *,
        column_headings: bool = False,
        row_headings: bool = False,
        buttons: Buttons = "static",
        color_system: str | None = "auto",
        force_terminal: bool | None = None,
        local_links: bool = False,
    ) -> None:
        if len(separator) != 1:
            raise ValueError("the separator must be one character")
        self.separator = separator
        self.local_links = local_links
        self.column_headings = column_headings
        self.row_headings = row_headings
        if buttons not in ("none", "static"):
            raise ValueError(f"buttons must be 'none' or 'static', not {buttons!r}")
        self.buttons = buttons
        self.color_system = color_system
        self.force_terminal = force_terminal

    @property
    def support(self) -> dict[str, Support]:
        """What this output shows of each feature the census finds (``backends/_support.py``)."""
        return _support.rich(
            headings=self.column_headings or self.row_headings,
            buttons=self.buttons,
            local_links=self.local_links,
        )

    def render(self, sheet: RenderedSheet, out: TextIO, /) -> None:
        """Print the sheet to ``out``, cropped to the terminal's width when ``out`` is one."""
        from rich.console import Console

        console = Console(
            file=out,
            color_system=self.color_system,  # type: ignore[arg-type]
            force_terminal=self.force_terminal,
            highlight=False,
            emoji=False,
            markup=False,
        )
        if not console.is_terminal:
            console.width = max(sheet.width + 12, 1)  # never wrap a file or a pipe
        for line in self.lines(sheet):
            console.print(line, no_wrap=True, overflow="crop", crop=True)

    def lines(self, sheet: RenderedSheet) -> Iterator[Any]:
        """One ``rich.text.Text`` per output line, each ending in a newline."""
        from rich.text import Text

        links = _Links(sheet, self.local_links)
        label = max((len(str(row.index)) for row in sheet.rows), default=1)
        dim = "dim"
        gutter = Gutter(sheet.outline if self.buttons != "none" else None)
        prefix = " " * gutter.width + (" " * label + self.separator if self.row_headings else "")
        # A grid color draws the separators, Excel's gridlines, in that color (approximated).
        grid = sheet.sheet.grid_color if sheet.sheet.show_gridlines else None
        if sheet.columns:
            for band in gutter.column_lines(sheet.columns, self.separator):
                yield Text(prefix + band, style=dim, end="\n")
        if self.column_headings and sheet.columns:
            heading = Text(prefix, style=dim, end="\n")
            for p, column in enumerate(sheet.columns):
                width = column.width or 0
                heading.append(column.letter.center(width)[:width], style=dim)
                if p < len(sheet.columns) - 1:
                    heading.append(self.separator, style=dim)
            yield heading
        buttons = self.buttons != "none"
        for row in sheet.rows:
            for line in range(row.height):
                text = Text(no_wrap=True, end="\n")
                if gutter.width:
                    text.append(gutter.row(row.index, line, row.height), style=dim)
                if self.row_headings:
                    number = str(row.index) if line == 0 else ""
                    text.append(number.rjust(label) + self.separator, style=dim)
                for p, cell in enumerate(row.cells):
                    style = _style(cell.font, cell.color, cell.fill)
                    chars = cell.lines[line]
                    if buttons and cell.button is not None:
                        chars = with_button(cell, line, row.height, chars)
                    links.append(text, chars, style, cell.source)
                    if p < len(row.gaps):
                        gap = row.gaps[p]
                        crossing = gap.lines[line]
                        if crossing is not None:
                            # Text crossing a gap comes from the cell both neighbors show.
                            after = row.cells[p + 1].source if p + 1 < len(row.cells) else None
                            style = _style(gap.font, gap.color, gap.fill)
                            links.append(
                                text, crossing, style, cell.source if after == cell.source else None
                            )
                        else:
                            text.append(self.separator, style=_style(None, grid, gap.fill))
                yield text


class _Links:
    """Each linked cell's ``link`` style, one per URL so that a terminal sees one link."""

    def __init__(self, sheet: RenderedSheet, local: bool = False) -> None:
        from rich.style import Style

        self.styles: dict[tuple[int, int], Any] = {}
        by_url: dict[str, Any] = {}
        for row in sheet.rows:
            for rendered in row.cells:
                cell = rendered.cell
                if cell.hyperlink is None or cell.kind == "empty" or cell.value == "":
                    continue
                url = external_url(cell.hyperlink, local)
                if url is not None:
                    style = by_url.setdefault(url, Style(link=url))
                    self.styles[(cell.row, cell.column)] = style

    def append(self, text: Any, chars: str, style: Any, source: tuple[int, int] | None) -> None:
        """Add ``chars`` to ``text``, the part that is not padding linked if its source is."""
        link = self.styles.get(source) if source is not None else None
        core = chars.strip(" ")
        if link is None or not core:
            text.append(chars, style=style)
            return
        start = chars.index(core)
        end = start + len(core)
        text.append(chars[:start], style=style)
        text.append(core, style=style + link)
        text.append(chars[end:], style=style)


def _style(font: Font | None, color: str | None, fill: str | None) -> Any:
    from rich.style import Style

    if font is None:  # a separator, in the grid color
        return Style(color=f"#{color}" if color else None, bgcolor=f"#{fill}" if fill else None)
    return Style(
        color=f"#{color}" if color else None,
        bgcolor=f"#{fill}" if fill else None,
        bold=font.bold or None,
        italic=font.italic or None,
        underline=font.underline in ("single", "singleAccounting") or None,
        underline2=font.underline in ("double", "doubleAccounting") or None,
        strike=font.strike or None,
    )
