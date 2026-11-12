"""The census: which features the laid-out cells and sheet use, for the outputs to report on.

An output declares, per feature code, whether it shows it, approximates it or leaves it out
(``Backend.support``); its notices are the census filtered through that declaration, so a
backend writes no reporting code. A feature an output does not declare is reported as left out.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass
from functools import lru_cache
from typing import TYPE_CHECKING, Any

from sheetpeek._fill_border_census import fill_border_codes
from sheetpeek._hyperlink import classify
from sheetpeek._model import CellRange, Font
from sheetpeek._notices import FEATURES, CellSet, Notice, Outcome, Reason, Support, describe
from sheetpeek._pass_outline import presence as outline_presence

if TYPE_CHECKING:
    from sheetpeek._rendered import RenderedSheet

_PLAIN_UNDERLINES = ("single", "double")
_ACCOUNTING_UNDERLINES = ("singleAccounting", "doubleAccounting")


@dataclass(frozen=True, slots=True)
class Presence:
    """Where a feature is used: a count, and for cell-level features the cells."""

    count: int
    cells: tuple[CellRange, ...] = ()
    truncated: bool = False


def census(rendered: RenderedSheet) -> dict[str, Presence]:
    """What the laid-out sheet uses, by feature code, in registry order."""
    sheet = rendered.sheet
    normal = sheet.normal_font
    sets: dict[str, CellSet] = {}
    shown = shown_cells(rendered)

    def add(code: str, row: int, column: int) -> None:
        found = sets.get(code)
        if found is None:
            found = sets[code] = CellSet()
        found.add(row, column)

    for rendered_row in rendered.rows:
        for rendered_cell in rendered_row.cells:
            cell = rendered_cell.cell
            style = cell.style
            row, column = cell.row, cell.column
            for code in fill_border_codes(sheet, rendered_cell):
                add(code, row, column)
            if cell.kind == "empty":
                continue
            if cell.hyperlink is not None and cell.value != "":
                kind = classify(cell.hyperlink, sheet.name, shown)[0]
                if kind != "refused":  # the reader reports it, once for every output
                    add(_LINKS[kind], row, column)
            if cell.runs:
                # A cell with rich-text runs uses its runs' fonts, each feature once (§4.16).
                used = {c for span in cell.runs for c in _font_codes(normal, span.font)}
                for code in used:
                    add(code, row, column)
                add("text.runs", row, column)
            else:
                for code in _font_codes(normal, style.font):
                    add(code, row, column)
            if style.shrink:
                add("text.shrink", row, column)
            if style.reading_order:
                add("text.reading_order", row, column)
            if rendered_cell.format_color is not None:
                add("format.color", row, column)
            if style.h_align != "general":
                add("align.horizontal", row, column)
            if style.v_align != "bottom":
                add("align.vertical", row, column)
            if style.wrap:
                add("align.wrap", row, column)
            if style.indent:
                add("align.indent", row, column)

    found: dict[str, Presence] = {}
    for code, cells in sets.items():
        ranges, truncated = cells.ranges()
        found[code] = Presence(cells.count, ranges, truncated)
    widths = {c.index: c.width for c in sheet.columns}
    sheet_level = {
        "merge": len(sheet.merges),
        "column.width": sum(
            1 for c in rendered.columns if not c.added and widths.get(c.index) is not None
        ),
        "row.height": sum(1 for r in rendered.rows if r.source.height is not None),
        "gridlines.hidden": 0 if sheet.show_gridlines else 1,
        "grid.color": 1 if sheet.show_gridlines and sheet.grid_color else 0,
        "headings.hidden": 0 if sheet.show_headings else 1,
        "outline": outline_presence(rendered),
    }
    for code, count in sheet_level.items():
        if count:
            found[code] = Presence(count)
    order = list(FEATURES)
    return dict(sorted(found.items(), key=lambda item: order.index(item[0])))


@lru_cache(maxsize=1024)
def _font_codes(normal: Font, font: Font) -> tuple[str, ...]:
    """The font features ``font`` uses, in registry order."""
    found = []
    if font.name != normal.name:
        found.append("font.name")
    if font.size != normal.size:
        found.append("font.size")
    if font.bold:
        found.append("font.bold")
    if font.italic:
        found.append("font.italic")
    if font.underline in _PLAIN_UNDERLINES:
        found.append("font.underline")
    elif font.underline in _ACCOUNTING_UNDERLINES:
        found.append("font.underline_accounting")
    if font.strike:
        found.append("font.strike")
    if font.vert_align is not None:
        found.append("font.vert_align")
    if font.color is not None:
        found.append("font.color")
    return tuple(found)


_LINKS = {
    "url": "link",
    "local": "link.local",
    "internal": "link.internal",
    "elsewhere": "link.elsewhere",
}


def shown_cells(rendered: RenderedSheet) -> Callable[[int, int], bool]:
    """Whether a cell is in the rendered rows and columns, worked out when first asked."""
    found: list[tuple[set[int], set[int]]] = []

    def shown(row: int, column: int) -> bool:
        if not found:
            found.append(({r.index for r in rendered.rows}, {c.index for c in rendered.columns}))
        rows, columns = found[0]
        return row in rows and column in columns

    return shown


def output_notices(
    rendered: RenderedSheet, output: str, support: Mapping[str, Support] | None
) -> tuple[Notice, ...]:
    """What ``output`` leaves out or approximates of what the sheet uses."""
    notices = []
    for code, presence in census(rendered).items():
        declared = support.get(code) if support is not None else None
        context: dict[str, str | int | float] = {}
        outcome: Outcome
        reason: Reason
        if declared is None:
            outcome, reason, how = "omitted", "not_implemented", "the output does not declare it"
            context["undeclared"] = 1
        elif declared.outcome is None:
            continue
        else:
            outcome, reason, how = (
                declared.outcome,
                declared.reason or "not_implemented",
                declared.how,
            )
        if presence.truncated:
            context["truncated"] = 1
        message = describe(code, outcome, reason, presence.count, how, output)
        notices.append(
            Notice(
                code,
                outcome,
                reason,
                message,
                rendered.sheet.name,
                presence.cells,
                presence.count,
                output,
                context,
            )
        )
    return tuple(notices)


def declared_support(backend: Any) -> Mapping[str, Support] | None:
    """A backend's declared support, or ``None`` when it declares none."""
    support = getattr(backend, "support", None)
    return support if isinstance(support, Mapping) else None
