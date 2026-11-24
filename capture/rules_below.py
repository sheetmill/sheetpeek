"""Capture whether Excel for Mac draws an underline past its row's bottom (route C, the screen).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.screen run rules_below
    uv run --group capture python -m capture.screen read rules_below     # the saved captures again
    uv run --group capture python -m capture.rules_below build PATH      # the workbook only

In ``text_style_more`` every underline ends inside its row: a fitted row holds a double accounting
underline's second line on its last point, and a fixed row too short for it lifts bottom-aligned
text. This suite sets rows too short for that: "HHH" under a double accounting underline at 11, 20
and 36 pt, a single accounting one at 36 pt and a single one at 36 pt, with plain "HHH" at 36 and
20 pt as controls (columns ``COLUMNS``, C to O, 14 characters, Calibri 11 Normal font), in rows
fixed at 12, 15 and 20 pt (``HEIGHTS``), bottom, center and top aligned, rows 4 to 20, each case
row followed by an empty row fixed at 30 pt (``GAP``) where a line drawn on past the row would show.
One window at zoom 100.

Measured per case, with ``text_style_more``'s code: the ink box and each rule's top in the cell;
and ``below``, the bands of image-pixel rows under the cell's bottom edge (in the empty row below)
that carry ink across at least a quarter of the cell's width less 6 pixels each side, as
``start+height`` in points from the cell's bottom edge, with ``below_ink`` the dark pixels there in
all. The next row's gridline is lighter than the ink threshold and is not counted. ``read`` writes
``tests/captures/rules_below.tsv``: one row per case.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils.cell import coordinate_from_string

from capture import text_style_more
from capture.text_style import Case

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "rules_below"
OUTPUT = ROOT / "tests" / "captures" / "rules_below.tsv"
COLUMNS: tuple[tuple[str, str | None, int], ...] = (
    ("C", "doubleAccounting", 11),
    ("E", "doubleAccounting", 20),
    ("G", "doubleAccounting", 36),
    ("I", "singleAccounting", 36),
    ("K", "single", 36),
    ("M", None, 36),
    ("O", None, 20),
)
"""Each column's underline (``None``: plain, a control) and size."""
HEIGHTS = (12, 15, 20)
VERTICALS = ("bottom", "center", "top")
GAP = 30
"""The empty row under each case row, fixed in points."""
MARKERS = ("B2", "Q3")
ZOOMS = (100,)
FIELDS = (
    "group", "name", "zoom", "cell", "row_height", "ink_top", "ink_bottom", "rule_tops",
    "rule_thickness", "below", "below_ink",
)  # fmt: skip


def fill(ws: Any) -> list[Case]:
    """Lay the cases out on a worksheet; every case."""
    for column, _, _ in COLUMNS:
        ws.column_dimensions[column].width = 14
    found = []
    row = 4
    for height in HEIGHTS:
        for vertical in VERTICALS:
            ws.row_dimensions[row].height = height
            ws.row_dimensions[row + 1].height = GAP
            for column, underline, size in COLUMNS:
                cell = ws[f"{column}{row}"]
                cell.value = "HHH"
                cell.font = Font(name="Calibri", size=size, underline=underline)
                cell.alignment = Alignment(vertical=vertical)
                label = f"{underline or 'plain'} {size} in {height} {vertical}"
                found.append(Case("rules_below", label, cell.coordinate, ("rules", "ink")))
            row += 2
    return found


def cases() -> list[Case]:
    return fill(openpyxl.Workbook().active)


def references() -> list[str]:
    """Row 1 of each column and column A of each case row (``text_style_more.complete``)."""
    columns = {f"{column}1": None for column, _, _ in COLUMNS}
    rows = {f"A{coordinate_from_string(c.cell)[1]}": None for c in cases()}
    return [*columns, *rows]


def build(path: Path) -> list[Case]:
    wb = openpyxl.Workbook()
    found = fill(wb.active)
    wb.save(path)
    return found


def run() -> None:
    """Capture the window, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    capture(
        fill,
        cells=references(),
        out=BUILD_DIR,
        markers=MARKERS,
        zooms=ZOOMS,
        name="rules_below",
    )
    write()


def below(cap: Any, cell: str) -> dict[str, str]:
    """The ink bands under ``cell``'s bottom edge, in the empty row below it."""
    from capture.screen.ink import pt
    from capture.screen.run import luma

    left, right = round(cap.edge(cell, "left")) + 6, round(cap.edge(cell, "right")) - 6
    bottom = round(cap.edge(cell, "bottom"))
    depth = round(GAP * cap.scale) - 2
    rows, total = [], 0
    for dy in range(depth):
        dark = sum(1 for x in range(left, right) if luma(cap.image.pixel(x, bottom + dy)) < 150)
        total += dark
        rows.append(dark >= (right - left) / 4)
    bands, start = [], None
    for dy, on in enumerate([*rows, False]):
        if on and start is None:
            start = dy
        elif not on and start is not None:
            bands.append(f"{pt(cap, start)}+{pt(cap, dy - start)}")
            start = None
    return {"below": " ".join(bands), "below_ink": str(total)}


def measure(cap: Any, case: Case) -> dict[str, str] | None:
    found = text_style_more.measure(cap, case)
    if found is None:
        return None
    return {**found, **below(cap, case.cell)}


def write() -> None:
    """Measure the saved capture again (no Excel) and write ``rules_below.tsv``."""
    from capture.screen.run import header, load

    records = []
    for cap in load(BUILD_DIR):
        for case in cases():
            found = measure(cap, case)
            if found is None:
                continue
            row = coordinate_from_string(case.cell)[1]
            records.append(
                {
                    "group": case.group,
                    "name": case.name,
                    "zoom": str(cap.zoom),
                    "cell": case.cell,
                    "row_height": cap.value(f"A{row}", "geometry.height") or "",
                    **found,
                }
            )
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in header(BUILD_DIR).items():
            if key not in ("markers", "suite", "workbooks_open_after", "zooms"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: rules_below\n" + "\t".join(FIELDS) + "\n")
        for record in records:
            out.write("\t".join(record.get(c, "") for c in FIELDS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        print(f"{len(build(Path(argv[1])))} cases in {argv[1]}")
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
