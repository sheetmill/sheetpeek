"""Number formats: the code is read.

A ``[ColorN]`` section color is drawn from the workbook's palette by the layout (``Sheet.palette``,
``ARCHITECTURE.md`` §4.14), so nothing of a code is left unread.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from sheetpeek._style import Context, Found

ATTRIBUTE = "number_format"


def read(part: Any, context: Context) -> dict[str, Any]:
    """The format code, ``General`` when the style names none."""
    return {"number_format": part if isinstance(part, str) and part else "General"}


def unread(part: Any, context: Context) -> Found:
    """Nothing: every part of a code is drawn."""
    return ([], [])
