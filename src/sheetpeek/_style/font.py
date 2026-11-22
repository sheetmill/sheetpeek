"""Fonts: name, size, bold, italic, underline, strike, color and superscript or subscript are read.

Outline, shadow, condense and extend are not read: Excel for Mac draws them as plain text
(``ARCHITECTURE.md`` §4.13). A theme font (``scheme`` major or minor) is drawn in the theme's
typeface whatever name it stores (§4.14), so its name is resolved here, when it is read.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import TYPE_CHECKING, Any, Literal

from sheetpeek._colors import DEFAULT_PALETTE, apply_tint
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
    if part is None:
        return {"font": context.normal_font}
    return {"font": to_font(part, context.normal_font, context.color, context.theme_fonts)}


_FIXED_INDICES = 8


def _fixed_index(color: Any) -> str | None:
    """A font color in indexed 0 to 7: the default palette's, whatever the workbook's.

    Captured for 3 (``theme_fonts_palette``: a recolored entry 3 draws a fill in the workbook's
    color and a font in the default's, where 10 and 17 draw both in the workbook's); 0 to 7
    taken alike (unverified).
    """
    index = getattr(color, "indexed", None) if color is not None else None
    if not isinstance(index, int) or not 0 <= index < _FIXED_INDICES:
        return None
    tint = getattr(color, "tint", 0.0) or 0.0
    return apply_tint(DEFAULT_PALETTE[index], tint) if tint else DEFAULT_PALETTE[index]


def to_font(
    font: Any,
    normal: Font,
    color_of: Callable[[Any], str | None],
    theme_fonts: Mapping[str, str] | None = None,
) -> Font:
    """An ``xlio.Font`` as the model's :class:`~sheetpeek.Font`.

    ``theme_fonts`` holds the theme's Latin typefaces, ``{"major": ..., "minor": ...}``: a font
    of either scheme is drawn in that typeface, whatever name it stores (captured,
    ``theme_fonts_palette``, §4.14).
    """
    color = font.color
    drawn = _fixed_index(color)
    # Indexed wins over the theme, and the theme over rgb and auto.
    theme = color.theme if color is not None and color.indexed is None else None
    stored = font.name or normal.name
    scheme = _SCHEMES.get(font.scheme)
    name = (theme_fonts or {}).get(scheme or "") or stored
    return Font(
        name=name,
        size=size(font.size) if font.size is not None else normal.size,
        bold=bool(font.bold),
        italic=bool(font.italic),
        underline=font.underline or None,
        strike=bool(font.strike),
        color=drawn if drawn is not None else color_of(color),
        vert_align=_VERT_ALIGN.get(font.vert_align),
        scheme=scheme,
        stored_name=stored if name != stored else None,
        theme_color=theme,
    )


def color_key(font: Any) -> tuple[Any, ...] | None:
    """An ``xlio.Font``'s color as the file states it, for comparing two fonts' colors.

    ``None`` for a font without a color. A cell's font color lets a table style's through where
    it equals the Normal style's font color (§4.20).
    """
    color = getattr(font, "color", None)
    if color is None:
        return None
    return (color.indexed, color.theme, color.auto, color.rgb, color.tint or 0.0)


_VERT_ALIGN: dict[Any, VerticalTextPosition] = {
    "superscript": "superscript",
    "subscript": "subscript",
}
_SCHEMES: dict[Any, Literal["major", "minor"]] = {"major": "major", "minor": "minor"}


def size(value: float) -> float:
    """A font size as Excel takes it: its magnitude, from 1 to 409.5, or refused (captured)."""
    points = abs(float(value))
    if not _SMALLEST <= points <= _LARGEST:  # NaN compares false
        raise PackageError(f"a font size of {value}, which Excel refuses")
    return points


def unread(part: Any, context: Context) -> Found:
    """Every font attribute Excel for Mac draws is read."""
    return [], []
