"""Number formats: the code is read; a ``[ColorN]`` the workbook palette recolors is reported.

``[ColorN]`` draws the workbook palette's index ``N + 7`` (``ARCHITECTURE.md`` §4.14), and the
layout draws section colors from Excel's default palette, so a format whose color the palette
changes is reported, only for cells with a value.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from sheetpeek._colors import recolored_sections

if TYPE_CHECKING:
    from sheetpeek._style import Context, Found

ATTRIBUTE = "number_format"


def read(part: Any, context: Context) -> dict[str, Any]:
    """The format code, ``General`` when the style names none."""
    return {"number_format": part if isinstance(part, str) and part else "General"}


def unread(part: Any, context: Context) -> Found:
    """``[ColorN]`` sections the workbook palette draws in another color."""
    found: Found = ([], [])
    if not isinstance(part, str):
        return found
    if recolored := recolored_sections(part, context.palette):
        colors = ", ".join(f"[Color{n}]" for n in recolored)
        found[1].append(
            (
                "format.palette",
                f"{colors} drawn in Excel's default palette; this workbook's palette changes "
                f"{'it' if len(recolored) == 1 else 'them'}",
                (("colors", colors),),
            )
        )
    return found
