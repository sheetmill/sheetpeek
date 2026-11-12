"""Fonts: name, size, bold, italic, underline, strike, color and superscript or subscript are read.

Outline, shadow, condense and extend are not read: Excel for Mac draws them as plain text
(``ARCHITECTURE.md`` §4.13). A theme font (``scheme`` major or minor) is drawn in the theme's
typeface whatever name it stores (§4.14), so its name is resolved here, when it is read.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import TYPE_CHECKING, Any, Literal

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
    return {"font": to_font(part, context.normal_font, context.color, context.theme_fonts)}


def to_font(
    font: Any,
    normal: Font,
    color_of: Callable[[Any], str | None],
    theme_fonts: Mapping[str, str] | None = None,
) -> Font:
    """An openpyxl ``Font`` as the model's :class:`~sheetpeek.Font`.

    ``theme_fonts`` holds the theme's Latin typefaces, ``{"major": ..., "minor": ...}``: a font
    of either scheme is drawn in that typeface, whatever name it stores (captured on Excel 16.114).
    """
    color = font.color
    # Theme color 1 ("Text 1") is what Excel's own default font uses for ordinary text, so
    # untinted it means the default foreground, like an automatic color.
    if getattr(color, "type", None) == "theme" and color.theme == 1 and not color.tint:
        color = None
    stored = font.name or normal.name
    scheme = _SCHEMES.get(font.scheme)
    name = (theme_fonts or {}).get(scheme or "") or stored
    return Font(
        name=name,
        size=size(font.sz) if font.sz is not None else normal.size,
        bold=bool(font.b),
        italic=bool(font.i),
        underline=font.u or None,
        strike=bool(font.strike),
        color=color_of(color),
        vert_align=_VERT_ALIGN.get(font.vertAlign),
        scheme=scheme,
        stored_name=stored if name != stored else None,
    )


_VERT_ALIGN: dict[Any, VerticalTextPosition] = {
    "superscript": "superscript",
    "subscript": "subscript",
}
_SCHEMES: dict[Any, Literal["major", "minor"]] = {"major": "major", "minor": "minor"}


def size(value: float) -> float:
    """A font size as Excel takes it: its magnitude, from 1 to 409, or refused (captured)."""
    points = abs(float(value))
    if not _SMALLEST <= points <= _LARGEST:  # NaN compares false
        raise PackageError(f"a font size of {value}, which Excel refuses")
    return points


def unread(part: Any, context: Context) -> Found:
    """Every font attribute Excel for Mac draws is read."""
    return [], []
