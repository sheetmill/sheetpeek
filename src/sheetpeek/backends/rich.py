"""Terminal output through `rich`: the fixed-width layout with fonts, colors and fills.

``rich`` is an optional dependency (``sheetpeek[rich]``), imported only when this backend draws.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any, TextIO

from sheetpeek._notices import Support
from sheetpeek._rendered import Mode, RenderedSheet
from sheetpeek.backends import _support

if TYPE_CHECKING:
    from sheetpeek._model import Font


class RichBackend:
    """Styled terminal text in ``characters`` mode.

    Each column is drawn on its own cell's fill, and the text in it takes the font and color of
    the cell it comes from, so text spilled over a filled neighbor shows that neighbor's fill,
    as in Excel.

    Args:
        separator: One character drawn between two columns where no text crosses.
        column_headings: Add Excel's column letters above the sheet.
        row_headings: Add Excel's row numbers before each row.
        color_system: Passed to ``rich.console.Console`` (``"auto"``, ``"truecolor"``, ...).
        force_terminal: Passed to ``rich.console.Console``; ``None`` detects.
    """

    name = "rich"
    mode: Mode = "characters"

    def __init__(
        self,
        separator: str = " ",
        *,
        column_headings: bool = False,
        row_headings: bool = False,
        color_system: str | None = "auto",
        force_terminal: bool | None = None,
    ) -> None:
        if len(separator) != 1:
            raise ValueError("the separator must be one character")
        self.separator = separator
        self.column_headings = column_headings
        self.row_headings = row_headings
        self.color_system = color_system
        self.force_terminal = force_terminal

    @property
    def support(self) -> dict[str, Support]:
        """What this output shows of each feature the census finds (``backends/_support.py``)."""
        return _support.rich(headings=self.column_headings or self.row_headings)

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

        label = max((len(str(row.index)) for row in sheet.rows), default=1)
        dim = "dim"
        if self.column_headings and sheet.columns:
            prefix = " " * label + self.separator if self.row_headings else ""
            heading = Text(prefix, style=dim, end="\n")
            for p, column in enumerate(sheet.columns):
                width = column.width or 0
                heading.append(column.letter.center(width)[:width], style=dim)
                if p < len(sheet.columns) - 1:
                    heading.append(self.separator, style=dim)
            yield heading
        for row in sheet.rows:
            for line in range(row.height):
                text = Text(no_wrap=True, end="\n")
                if self.row_headings:
                    number = str(row.index) if line == 0 else ""
                    text.append(number.rjust(label) + self.separator, style=dim)
                for p, cell in enumerate(row.cells):
                    text.append(cell.lines[line], style=_style(cell.font, cell.color, cell.fill))
                    if p < len(row.gaps):
                        gap = row.gaps[p]
                        crossing = gap.lines[line]
                        if crossing is not None:
                            text.append(crossing, style=_style(gap.font, gap.color, gap.fill))
                        else:
                            text.append(self.separator, style=_style(None, None, gap.fill))
                yield text


def _style(font: Font | None, color: str | None, fill: str | None) -> Any:
    from rich.style import Style

    if font is None:
        return Style(bgcolor=f"#{fill}" if fill else None)
    return Style(
        color=f"#{color}" if color else None,
        bgcolor=f"#{fill}" if fill else None,
        bold=font.bold or None,
        italic=font.italic or None,
        underline=font.underline in ("single", "singleAccounting") or None,
        underline2=font.underline in ("double", "doubleAccounting") or None,
        strike=font.strike or None,
    )
