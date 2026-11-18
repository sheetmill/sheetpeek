"""Fills and borders in the census, and the style that draws a laid-out cell's fill."""

from __future__ import annotations

from typing import TYPE_CHECKING

from sheetpeek._fill_border_model import WHITE, Gradient, Pattern

if TYPE_CHECKING:
    from sheetpeek._model import CellStyle, Sheet
    from sheetpeek._rendered import RenderedCell


def fill_style(sheet: Sheet, rendered: RenderedCell) -> CellStyle | None:
    """The style whose fill a laid-out cell shows: its displayed style, or its merge anchor's.

    A merge shows its anchor's own style, a table's beneath it not stacked (unverified). ``None``
    for a merge whose anchor lies outside the range read.
    """
    merge = rendered.merge
    if merge is None:
        return rendered.display_style
    ref = sheet.ref
    if ref is None or (merge.min_row, merge.min_col) not in ref:
        return None
    return sheet.cell(merge.min_row, merge.min_col).style


def fill_border_codes(sheet: Sheet, rendered: RenderedCell) -> list[str]:
    """The census codes of what a laid-out cell draws of fills and borders."""
    codes = []
    style = fill_style(sheet, rendered)
    spec = style.fill_spec if style is not None else None
    if isinstance(spec, Pattern):
        if spec.kind != "solid":
            codes.append("fill.pattern")
        elif spec.foreground != WHITE:  # white is the paper an output without fills draws on
            codes.append("fill")
    elif isinstance(spec, Gradient):
        codes.append("fill.gradient")
    borders = rendered.display_style.borders
    if borders is None:
        return codes
    sides = [s for s in borders.sides if s is not None]
    if sides:
        codes.append("border")
    diagonal = borders.diagonal
    if diagonal is not None and (rendered.merge is None or rendered.is_merge_anchor):
        codes.append("border.diagonal")
        sides.append(diagonal)
    if any(s.style == "slantDashDot" for s in sides):
        codes.append("border.slant_dash_dot")
    return codes
