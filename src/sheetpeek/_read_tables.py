"""Reading tables: the sheet's table parts and the styles they name (§4.20).

xlio reads the parts and the styles, the workbook's own and Excel's presets
(``Styles.table_style`` looks the workbook's up first). A differential format sets only what its
``<dxf>`` holds: xlio's font record gives defaults for what is absent, so which font properties
are set is read from the element's own children. A solid fill, or a ``patternFill`` without a
type (as differential formats write one), is drawn in its ``bgColor``.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any, cast

import xlio

from sheetpeek._fill_border_model import TILES, Pattern, PatternKind
from sheetpeek._model import CellRange
from sheetpeek._style import Context
from sheetpeek._style.border import side
from sheetpeek._style.fill import color
from sheetpeek._table_model import (
    ELEMENT_ORDER,
    DifferentialFormat,
    Table,
    TableStyle,
    TableStyleElement,
    TableStyleInfo,
)

_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_UNDERLINES = frozenset({"single", "double", "singleAccounting", "doubleAccounting", "none"})


def read_tables(
    source: xlio.Sheet,
    ref: CellRange | None,
    lookup: Callable[[str], Any],
    context: Context,
) -> tuple[tuple[Table, ...], tuple[TableStyle, ...]]:
    """The tables that touch ``ref``, and the styles they name that ``lookup`` finds."""
    if ref is None:
        return (), ()
    tables: list[Table] = []
    styles: dict[str, TableStyle] = {}
    for part in source.tables:
        table = _table(part)
        if table is None or not table.ref.intersects(ref):
            continue
        tables.append(table)
        name = table.style.name if table.style is not None else None
        if name and name not in styles:
            found = lookup(name)
            if found is not None:
                styles[name] = _style(found, context)
    return tuple(tables), tuple(styles.values())


def _table(part: Any) -> Table | None:
    try:
        ref = CellRange.parse(part.ref)
    except (TypeError, ValueError):
        return None
    info = part.style
    style = (
        None
        if info is None
        else TableStyleInfo(
            info.name or None,
            bool(info.show_first_column),
            bool(info.show_last_column),
            bool(info.show_row_stripes),
            bool(info.show_column_stripes),
        )
    )
    header = 1 if part.header_row_count else 0
    totals = 1 if part.totals_row_count else 0
    return Table(part.display_name or part.name, ref, header, totals, style)


def _style(found: Any, context: Context) -> TableStyle:
    elements = []
    for element in found.elements:
        if element.type not in ELEMENT_ORDER or element.dxf is None:
            continue
        size = element.size if isinstance(element.size, int) and 1 <= element.size <= 9 else 1
        elements.append(TableStyleElement(element.type, differential(element.dxf, context), size))
    return TableStyle(found.name, tuple(elements))


def differential(dxf: Any, context: Context) -> DifferentialFormat:
    """An ``xlio.DifferentialFormat`` as the model's, with only what the ``<dxf>`` sets."""
    font: dict[str, Any] = {}
    if dxf.font is not None:
        given = _children(dxf.font.source)
        if "b" in given:
            font["bold"] = bool(dxf.font.bold)
        if "i" in given:
            font["italic"] = bool(dxf.font.italic)
        if "strike" in given:
            font["strike"] = bool(dxf.font.strike)
        if "u" in given:
            underline = dxf.font.underline or "single"
            font["underline"] = underline if underline in _UNDERLINES else "single"
        if "color" in given and dxf.font.color is not None:
            font["color"] = context.color(dxf.font.color)
    sides: dict[str, Any] = {}
    if dxf.border is not None:
        for name in ("left", "right", "top", "bottom", "horizontal", "vertical"):
            found = side(getattr(dxf.border, name, None), context)
            if found is not None:
                sides[name] = found
    return DifferentialFormat(**font, fill=_fill(dxf.fill, context), **sides)


def _fill(part: Any, context: Context) -> Pattern | None:
    if part is None or not hasattr(part, "pattern"):
        return None  # a gradient in a differential format: not drawn yet
    kind = part.pattern or "solid"
    if kind == "solid":
        found = color(part.background, context, background=True)
        return Pattern("solid", found) if found is not None else None
    if kind not in TILES:
        return None
    tile = cast("PatternKind", kind)
    return Pattern(
        tile, color(part.foreground, context), color(part.background, context, background=True)
    )


def _children(element: Any) -> set[str]:
    if element is None:
        return set()
    return {child.tag.removeprefix(_MAIN) for child in element}
