"""Filter buttons in characters: a ``▾`` in a header cell's last character (§4.19).

Excel draws the button at the cell's bottom right, over the header's text; in characters it takes
the last character of the cell's bottom line, the same place, whatever its state.
"""

from __future__ import annotations

from sheetpeek._rendered import RenderedCell

MARK = "▾"


def with_button(cell: RenderedCell, line: int, height: int, chars: str) -> str:
    """``chars`` of ``cell``'s line ``line`` of ``height``, its last character the button's."""
    if cell.button is None or line != height - 1 or not chars:
        return chars
    return chars[:-1] + MARK
