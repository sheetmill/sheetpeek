"""Fonts: name, size, bold, italic, underline, strike, color and superscript or subscript are read.

Outline, shadow, condense and extend are not read: Excel for Mac draws them as plain text
(``ARCHITECTURE.md`` §4.13). A scheme font whose stored name is not the theme's is reported.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING, Any

from sheetpeek._model import Font, VerticalTextPosition
from sheetpeek._package import PackageError

if TYPE_CHECKING:
    from sheetpeek._style import Context, Found

ATTRIBUTE = "font"
FALLBACK = Font("Calibri", 11.0)
"""The Normal font when a workbook names none."""
_SMALLEST = 1.0
_LARGEST = 409.5  # 410 is refused


def read(part: Any, context: Context) -> dict[str, Any]:
    """The cell's font, with the Normal font's name and size where the font gives none."""
    return {"font": to_font(part, context.normal_font, context.color)}


def to_font(font: Any, normal: Font, color_of: Callable[[Any], str | None]) -> Font:
    """An openpyxl ``Font`` as the model's :class:`~sheetpeek.Font`."""
    color = font.color
    # Theme color 1 ("Text 1") is what Excel's own default font uses for ordinary text, so
    # untinted it means the default foreground, like an automatic color.
    if getattr(color, "type", None) == "theme" and color.theme == 1 and not color.tint:
        color = None
    return Font(
        name=font.name or normal.name,
        size=size(font.sz) if font.sz is not None else normal.size,
        bold=bool(font.b),
        italic=bool(font.i),
        underline=font.u or None,
        strike=bool(font.strike),
        color=color_of(color),
        vert_align=_VERT_ALIGN.get(font.vertAlign),
    )


_VERT_ALIGN: dict[Any, VerticalTextPosition] = {
    "superscript": "superscript",
    "subscript": "subscript",
}


def size(value: float) -> float:
    """A font size as Excel takes it: its magnitude, from 1 to 409, or refused (captured)."""
    points = abs(float(value))
    if not _SMALLEST <= points <= _LARGEST:  # NaN compares false
        raise PackageError(f"a font size of {value}, which Excel refuses")
    return points


def unread(part: Any, context: Context) -> Found:
    """A scheme font whose stored name is not the theme's."""
    found: Found = ([], [])
    theme = context.theme_fonts.get(part.scheme or "")
    if theme and part.name and part.name.lower() != theme.lower():
        found[1].append(
            (
                "font.scheme",
                f"drawn in {part.name!r}; Excel draws the theme's {theme!r}",
                (("font", str(part.name)), ("theme", theme)),
            )
        )
    return found
