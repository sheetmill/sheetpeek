"""Resolve stored colors (RGB, indexed, theme with tint, automatic) to ``RRGGBB``.

Theme colors follow ECMA-376: the theme's ``clrScheme`` gives the base color and ``tint``
lightens or darkens it in HLS space (§18.8.19). Captured (suite ``theme_colors``, 96 cells):
Excel's index order swaps the first two pairs, and the tint is applied in Windows' integer HLS
on a scale of 240 (``ColorRGBToHLS``), the lightness truncated, which reproduces every cell.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from collections.abc import Sequence

_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
# clrScheme children in document order.
_SCHEME = (
    "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
    "accent6", "hlink", "folHlink",
)  # fmt: skip
# Excel's theme index → clrScheme entry: the light and dark pairs are swapped.
_THEME_ORDER = (1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11)
_AUTOMATIC_INDEX = (64, 65)  # system foreground and background
_HLS = 240  # Windows' HLSMAX
_RGB = 255


_HEX = re.compile("[0-9A-F]{6}")


def theme_colors(theme_xml: bytes | str | None) -> tuple[str, ...]:
    """The twelve theme colors as ``RRGGBB``, in Excel's theme index order.

    A workbook without a theme part, or with an unreadable one, has no theme colors.
    """
    if not theme_xml:
        return ()
    try:
        root = ET.fromstring(theme_xml)
    except ET.ParseError:
        return ()
    scheme = root.find(f".//{_A}clrScheme")
    if scheme is None:
        return ()
    found: dict[str, str] = {}
    for name in _SCHEME:
        entry = scheme.find(f"{_A}{name}")
        if entry is None:
            return ()
        srgb = entry.find(f"{_A}srgbClr")
        system = entry.find(f"{_A}sysClr")
        if srgb is not None:
            found[name] = srgb.get("val", "000000").upper()
        elif system is not None:
            found[name] = system.get("lastClr", "000000").upper()
        else:
            return ()
        # The theme's text reaches CSS as it stands, so nothing but a color may leave here.
        if not _HEX.fullmatch(found[name]):
            return ()
    return tuple(found[_SCHEME[i]] for i in _THEME_ORDER)


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
    result is truncated after adding a small tolerance (captured, suite ``theme_colors``, and a
    probe on Excel 16.114 of 14 cells). Files store a tint as a 16-bit fraction (0.6 as
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
