"""Resolve stored colors (RGB, indexed, theme with tint, automatic) to ``RRGGBB``.

Theme colors follow ECMA-376: the theme's ``clrScheme`` gives the base color and ``tint``
lightens or darkens it in HLS space (§18.8.19). Captured (suite ``theme_colors``, 96 cells):
Excel's index order swaps the first two pairs, and the tint is applied in Windows' integer HLS
on a scale of 240 (``ColorRGBToHLS``), the lightness truncated, which reproduces every cell.
The workbook's indexed palette, the grid color and number-format colors are §4.14.
"""

from __future__ import annotations

import re
from collections.abc import Sequence

from xlformat import DEFAULT_PALETTE as _SECTION_COLORS

_AUTOMATIC_INDEX = (64, 65)  # system foreground and background
SYSTEM_BACKGROUND = 65
"""The indexed system background: white in a cell's font, border and fill (``auto_colors``)."""
_HLS = 240  # Windows' HLSMAX
_RGB = 255


_HEX = re.compile("[0-9A-F]{6}")

PALETTE_SIZE = 64
"""Entries of the indexed palette; Excel caps ``indexedColors`` at 64 ([MS-OI29500] 2.1.709)."""
DEFAULT_PALETTE: tuple[str, ...] = tuple(
    _SECTION_COLORS[f"Color{i}"] for i in (*range(1, 9), *range(1, 57))
)
"""Excel's default palette as ``RRGGBB``, indices 0 to 63: 0 to 7 repeat 8 to 15, and
``[ColorN]`` of a number format is index ``N + 7``."""
_AUTOMATIC_GRID = 64


def workbook_palette(custom: Sequence[str | None] | None) -> tuple[str, ...]:
    """The palette a workbook's indexed colors take: its ``indexedColors``, else the default.

    Each written entry is used as it stands, so indices 0 to 7 no longer repeat 8 to 15 (in a
    fill; a font in 0 to 7 keeps the default's, ``_style.font``). Entries past 64 are dropped, as
    Excel ignores them (a fill in 66 of a list of 70 drew a gray of its own, 646464, not the
    list's entry nor black: a known deviation). A shorter list takes the default for the indices
    it leaves out; an entry that is not a color is black (``theme_fonts_palette``).
    """
    if not custom:
        return DEFAULT_PALETTE

    def entry(index: int) -> str:
        if index >= len(custom):
            return DEFAULT_PALETTE[index]
        found = custom[index]
        color = found[-6:].upper() if isinstance(found, str) else ""
        return color if _HEX.fullmatch(color) else "000000"

    return tuple(entry(i) for i in range(PALETTE_SIZE))


def grid_color(color_id: int | None, palette: Sequence[str]) -> str | None:
    """The gridline color of ``<sheetView defaultGridColor="0" colorId>``, or ``None`` (gray).

    The color is the workbook palette's entry, and 64 is the automatic gray (captured,
    ``headings_grid``). Excel refuses a file whose ``colorId`` is 65 or past the palette (99999);
    sheetpeek takes them as automatic, as it does a malformed one, which Excel draws gray.
    """
    if color_id is None or not 0 <= color_id < min(len(palette), _AUTOMATIC_GRID):
        return None
    return palette[color_id]


def resolve(
    kind: str | None,
    value: str | int | None,
    tint: float,
    theme: Sequence[str],
    palette: Sequence[str],
) -> str | None:
    """A stored color as ``RRGGBB``, or ``None`` for automatic or unresolvable colors.

    Args:
        kind: How the file states the color: ``rgb``, ``indexed``, ``theme``, or anything else
            (``auto``, missing) for automatic.
        value: ``AARRGGBB`` or ``RRGGBB`` text for ``rgb``; the index for ``indexed`` and
            ``theme``.
        tint: Lightening (positive) or darkening (negative), from -1 to 1.
        theme: Theme colors in Excel's index order (:func:`theme_colors`).
        palette: The indexed palette as ``AARRGGBB`` strings (custom or default).
    """
    if kind == "rgb":
        if not isinstance(value, str) or len(value) not in (6, 8):
            return None
        base = value[-6:].upper()
    elif kind == "indexed":
        if not isinstance(value, int) or value in _AUTOMATIC_INDEX or not 0 <= value < len(palette):
            return None
        base = palette[value][-6:].upper()
    elif kind == "theme":
        if not isinstance(value, int) or not 0 <= value < len(theme):
            return None
        base = theme[value]
    else:
        return None
    if not _HEX.fullmatch(base):
        return None
    return apply_tint(base, tint) if tint else base


_TINT_TOLERANCE = 0.01
"""Added to a tinted lightness before truncating it: 16-bit stored tints fall just short."""


def apply_tint(rgb: str, tint: float) -> str:
    """Lighten (``tint > 0``) or darken (``tint < 0``) an ``RRGGBB`` color as Excel does.

    ECMA-376's tint on the lightness, in Windows' integer HLS (scale 240): darkening multiplies
    the lightness by ``1 + tint``, lightening moves it ``tint`` of the way to white, and the
    result is truncated after adding a small tolerance (captured, suites ``theme_colors`` and,
    for stored fractions, ``theme_colors_more``). Files store a tint as a 16-bit fraction (0.6 as
    0.59999389...), which would otherwise land a lightness just below a whole number; any
    tolerance from 0.0055 to 0.2 fits every captured cell.
    """
    if not tint:
        return rgb
    hue, lightness, saturation = _to_hls(*(int(rgb[i : i + 2], 16) for i in (0, 2, 4)))
    tinted = lightness * (1 + tint) if tint < 0 else lightness * (1 - tint) + _HLS * tint
    red, green, blue = _to_rgb(hue, max(0, min(_HLS, int(tinted + _TINT_TOLERANCE))), saturation)
    return f"{red:02X}{green:02X}{blue:02X}"


def _to_hls(red: int, green: int, blue: int) -> tuple[int, int, int]:
    """Windows' ``ColorRGBToHLS``: hue, lightness and saturation from 0 to 240."""
    high, low = max(red, green, blue), min(red, green, blue)
    lightness = ((high + low) * _HLS + _RGB) // (2 * _RGB)
    if high == low:
        return 0, lightness, 0
    spread, total = high - low, high + low
    if lightness <= _HLS // 2:
        saturation = (spread * _HLS + total // 2) // total
    else:
        saturation = (spread * _HLS + (2 * _RGB - total) // 2) // (2 * _RGB - total)
    sixth = _HLS // 6
    r, g, b = (((high - c) * sixth + spread // 2) // spread for c in (red, green, blue))
    if red == high:
        hue = b - g
    elif green == high:
        hue = _HLS // 3 + r - b
    else:
        hue = 2 * _HLS // 3 + g - r
    return hue % _HLS, lightness, saturation


def _to_rgb(hue: int, lightness: int, saturation: int) -> tuple[int, int, int]:
    """Windows' ``ColorHLSToRGB``."""
    if saturation == 0:
        grey = (lightness * _RGB + _HLS // 2) // _HLS
        return grey, grey, grey
    if lightness <= _HLS // 2:
        high = (lightness * (_HLS + saturation) + _HLS // 2) // _HLS
    else:
        high = lightness + saturation - (lightness * saturation + _HLS // 2) // _HLS
    low = 2 * lightness - high
    red, green, blue = (
        (_hue(low, high, hue + shift) * _RGB + _HLS // 2) // _HLS
        for shift in (_HLS // 3, 0, -_HLS // 3)
    )
    return red, green, blue


def _hue(low: int, high: int, hue: int) -> int:
    hue %= _HLS
    sixth = _HLS // 6
    if hue < sixth:
        return low + ((high - low) * hue + _HLS // 12) // sixth
    if hue < _HLS // 2:
        return high
    if hue < 2 * _HLS // 3:
        return low + ((high - low) * (2 * _HLS // 3 - hue) + _HLS // 12) // sixth
    return low
