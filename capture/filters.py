"""Capture how Excel for Mac draws filter and sort buttons, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.screen run filters
    uv run --group capture python -m capture.screen read filters     # the saved captures again
    uv run --group capture python -m capture.filters build PATH      # the workbook only

The sheet's filter is C4:J6, its header row 15 pt: C plain, D a values filter, E a custom filter,
F sorted ascending, G descending, H filtered and sorted ascending, I with ``hiddenButton``, J with
``showButton="0"``; its ``sortState`` inside its ``autoFilter``. Table ``Sorts`` (L4:O6) sorts,
in a ``sortState`` beside its ``autoFilter``, L by cell color descending, M by font color, N by
icon, and O filtered and descending (a captured glyph, the control). Table ``Tall`` (L9:N11) has
a 30 pt header row: L plain, M filtered, N sorted ascending. One window, at zoom 100 and 200;
the second marker is Q13, which at zoom 200 shows in a window about 1500 points wide (a narrower
window needs only the first marker, ``CAPTURING.md``, "One marker, on a small screen").

Measured per header cell, in screen points (image pixels over the display's scale): the
button's box (the columns, then the rows, at least 60 % unlike the cell's background in the
square the row allows plus 4 to its left), its gaps to the cell's
right and bottom boundaries and its top below the row's top; the glyph inside it (pixels with
luma under 150) as a bitmap with an id, equal ids meaning identical glyphs, its size and mean
color. ``read`` writes ``tests/captures/filters.tsv``, one row per zoom and cell.
"""

from __future__ import annotations

import hashlib
import sys
from collections import Counter
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet import filters
from openpyxl.worksheet.table import Table, TableColumn

CASES = {
    "C4": "plain",
    "D4": "values filter",
    "E4": "custom filter",
    "F4": "sorted ascending",
    "G4": "sorted descending",
    "H4": "filtered and sorted ascending",
    "I4": "hiddenButton",
    "J4": "showButton 0",
    "L4": "sorted by cell color descending",
    "M4": "sorted by font color",
    "N4": "sorted by icon",
    "O4": "filtered and sorted descending",
    "L9": "plain, 30 pt row",
    "M9": "filtered, 30 pt row",
    "N9": "sorted ascending, 30 pt row",
}
EXPECTED = {
    "C4": "plain", "D4": "filtered", "E4": "filtered", "F4": "ascending", "G4": "descending",
    "H4": "filtered_ascending", "I4": None, "J4": None, "L4": "descending", "M4": "ascending",
    "N4": "ascending", "O4": "filtered_descending", "L9": "plain", "M9": "filtered",
    "N9": "ascending",
}  # fmt: skip
"""sheetpeek's glyph for each cell; ``None`` draws no button."""
MARKERS = ("B2", "Q13")
ZOOMS = (100, 200)
LARGEST = 17
SEARCH = 4
ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "filters"
OUTPUT = ROOT / "tests" / "captures" / "filters.tsv"
COLUMNS = (
    "zoom", "cell", "case", "box_w", "box_h", "right_gap", "bottom_gap", "top", "glyph",
    "glyph_w", "glyph_h", "glyph_color",
)  # fmt: skip


def _values(ws: Any, top: int, left: int, names: list[str]) -> None:
    for k, name in enumerate(names):
        ws.cell(top, left + k, name)
        for r in (1, 2):
            ws.cell(top + r, left + k, r * (k + 1))


def _columns(names: list[str]) -> list[TableColumn]:
    """A table's columns, named as their header cells: Excel refuses a table whose names differ."""
    return [TableColumn(id=k + 1, name=name) for k, name in enumerate(names)]


def fill(ws: Any) -> list[tuple[str, str]]:
    """Lay the filters out on a worksheet; each header cell with its case."""
    wb = ws.parent
    for column in "CDEFGHIJLMNO":
        ws.column_dimensions[column].width = 12
    _values(ws, 4, 3, ["plain", "values", "custom", "up", "down", "both", "hidden", "show0"])
    ws.auto_filter.ref = "C4:J6"
    columns = ws.auto_filter.filterColumn
    columns.append(filters.FilterColumn(colId=1, filters=filters.Filters(filter=["2"])))
    custom = filters.CustomFilter(operator="greaterThan", val="1")
    columns.append(
        filters.FilterColumn(colId=2, customFilters=filters.CustomFilters(customFilter=[custom]))
    )
    columns.append(filters.FilterColumn(colId=5, filters=filters.Filters(filter=["2"])))
    columns.append(filters.FilterColumn(colId=6, hiddenButton=True))
    columns.append(filters.FilterColumn(colId=7, showButton=False))
    ws.auto_filter.sortState = filters.SortState(
        ref="C5:J6",
        sortCondition=[
            filters.SortCondition(ref="F5:F6"),
            filters.SortCondition(ref="G5:G6", descending=True),
            filters.SortCondition(ref="H5:H6"),
        ],
    )
    wb._differential_styles.append(DifferentialStyle(fill=PatternFill(bgColor="FFFF0000")))
    _values(ws, 4, 12, ["cell", "font", "icon", "both"])
    sorts = Table(displayName="Sorts", ref="L4:O6")
    sorts.tableColumns = _columns(["cell", "font", "icon", "both"])
    sorts.autoFilter = filters.AutoFilter(
        ref="L4:O6",
        filterColumn=[filters.FilterColumn(colId=3, filters=filters.Filters(filter=["2"]))],
    )
    sorts.sortState = filters.SortState(
        ref="L5:O6",
        sortCondition=[
            filters.SortCondition(ref="L5:L6", descending=True, sortBy="cellColor", dxfId=0),
            filters.SortCondition(ref="M5:M6", sortBy="fontColor", dxfId=0),
            filters.SortCondition(ref="N5:N6", sortBy="icon", iconSet="3Arrows", iconId=0),
            filters.SortCondition(ref="O5:O6", descending=True),
        ],
    )
    ws.add_table(sorts)
    _values(ws, 9, 12, ["plain", "filtered", "up"])
    ws.row_dimensions[9].height = 30
    tall = Table(displayName="Tall", ref="L9:N11")
    tall.tableColumns = _columns(["plain", "filtered", "up"])
    tall.autoFilter = filters.AutoFilter(
        ref="L9:N11",
        filterColumn=[filters.FilterColumn(colId=1, filters=filters.Filters(filter=["2"]))],
    )
    tall.sortState = filters.SortState(
        ref="L10:N11", sortCondition=[filters.SortCondition(ref="N10:N11")]
    )
    ws.add_table(tall)
    return list(CASES.items())


def build(path: Path) -> list[tuple[str, str]]:
    """Write the filters to one workbook at ``path``; each header cell with its case."""
    wb = openpyxl.Workbook()
    wb.active.title = "Filters"
    placed = fill(wb.active)
    wb.save(path)
    return placed


def run() -> None:
    """Capture the window at both zooms, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    capture(fill, cells=list(CASES), out=BUILD_DIR, markers=MARKERS, zooms=ZOOMS, name="filters")
    write()


def _measure(cap: Any, cell: str, case: str | None = None) -> dict[str, str]:
    """One header cell's button and glyph; ``case`` names it (by default its ``CASES`` entry)."""
    from capture.screen.run import luma

    per = cap.scale * 100 / cap.zoom  # image pixels a screen point
    right, bottom, top = cap.edge(cell, "right"), cap.edge(cell, "bottom"), cap.edge(cell, "top")
    side = min(bottom - top, LARGEST * per)
    x0 = round(max(cap.edge(cell, "left"), right - side - SEARCH * per))
    y0, x1, y1 = max(round(top) - 2, round(bottom - side - 2 * per)), round(right), round(bottom)
    bx, by = round(cap.edge(cell, "left")) + 3, round(top) + 3
    background = Counter(
        cap.image.pixel(x, y)
        for y in range(by, max(by + 1, round(bottom) - 2))
        for x in range(bx, max(bx + 1, x0 - 2))
    ).most_common(1)[0][0]

    def off(p: tuple[int, int, int]) -> bool:
        return max(abs(a - b) for a, b in zip(p, background, strict=True)) > 3

    found: dict[str, str] = {
        "zoom": str(cap.zoom),
        "cell": cell,
        "case": CASES[cell] if case is None else case,
    }
    columns = [
        x
        for x in range(x0, x1)
        if sum(off(cap.image.pixel(x, y)) for y in range(y0, y1)) >= 0.6 * side
    ]
    if not columns:
        return found
    bx0, bx1 = columns[0], columns[-1]
    rows = [
        y
        for y in range(y0, y1)
        if sum(off(cap.image.pixel(x, y)) for x in range(bx0, bx1 + 1)) >= 0.6 * (bx1 - bx0)
    ]
    if not rows:
        return found
    by0, by1 = rows[0], rows[-1]

    def spt(pixels: float) -> str:
        return f"{round(pixels / per, 2):g}"

    found |= {
        "box_w": spt(bx1 - bx0 + 1),
        "box_h": spt(by1 - by0 + 1),
        "right_gap": spt(right - bx1 - 1),
        "bottom_gap": spt(bottom - by1 - 1),
        "top": spt(by0 - top),
    }
    dark = [
        (x, y)
        for y in range(by0 + 2, by1 - 1)
        for x in range(bx0 + 2, bx1 - 1)
        if luma(cap.image.pixel(x, y)) < 150
    ]
    if dark:
        gx0, gy0 = min(x for x, _ in dark), min(y for _, y in dark)
        gx1, gy1 = max(x for x, _ in dark), max(y for _, y in dark)
        marked = set(dark)
        bitmap = "/".join(
            "".join("#" if (x, y) in marked else "." for x in range(gx0, gx1 + 1))
            for y in range(gy0, gy1 + 1)
        )
        mean = [round(sum(cap.image.pixel(x, y)[i] for x, y in dark) / len(dark)) for i in range(3)]
        found |= {
            "glyph": hashlib.sha1(bitmap.encode()).hexdigest()[:8],
            "glyph_w": spt(gx1 - gx0 + 1),
            "glyph_h": spt(gy1 - gy0 + 1),
            "glyph_color": "".join(f"{c:02X}" for c in mean),
        }
    return found


def measure(cap: Any) -> list[dict[str, str]]:
    """Every header cell's record at one zoom."""
    return [_measure(cap, cell) for cell in CASES]


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``filters.tsv``."""
    from capture.screen.run import header, load

    first = header(BUILD_DIR)
    records = [record for cap in load(BUILD_DIR) for record in measure(cap)]
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: filters\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cells in {argv[1]}")
        return 0
    if argv == ["run"]:
        run()
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
