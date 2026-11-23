"""Capture which cells set a row's shared bottom baseline, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.row_bottoms run
    uv run --group capture python -m capture.row_bottoms run cells   # some groups only
    uv run --group capture python -m capture.row_bottoms read     # the saved captures again
    uv run --group capture python -m capture.row_bottoms build DIR    # the workbooks only

Bottom-aligned text in a row sits on one baseline, the largest bottom offset ``B`` over the
Normal font and the row's fonts (ARCHITECTURE.md §3.6; captured for cells with values in
``text_style_more``). This suite asks which other cells count, and where text that spans rows
or spreads down sits. Every probe is "HHH" (no descender: its ink ends on the baseline) in
Calibri 11 under a Calibri 11 Normal font (``B`` 3), bottom aligned in column C; the neighbor
is Calibri 24 (``B`` 6) in column E unless stated. Rows are fixed at 32 pt, so a counted
neighbor puts the probe's baseline 26 pt below the row's top and one not counted 29. Each group
is its own workbook and window at zoom 100, in the dark appearance as the other suites.

- ``cells``. Row 4: the probe alone (the control). Row 5: a neighbor with a value (the captured
  rule). Row 6: an empty cell styled Calibri 24. Row 7: a value in Calibri 24 in the hidden
  column G. Rows 8, 9 and 10: the neighbor turned 45 degrees, turned 90 degrees, and stacked.
  Rows 11 and 12: the neighbor merged down E11:E12, anchored in row 11, a probe in each row.
  Rows 13 and 14: the neighbor in superscript and in subscript (drawn at about 16 pt, ``B`` 4:
  a baseline of 28 if the drawn size counts). Row 15: the row's own style Calibri 24
  (``customFormat``), no neighbor cell. Row 16: a Calibri 11 neighbor with a 24 pt rich-text
  run. Rows 17 and 18: the probe merged down C17:C18 (64 pt), the neighbor in row 17 (the
  merge's first row); rows 19 and 20 the same with the neighbor in row 20 (its last row): which
  row's bottom a merge across rows takes (61 from the merge's top when no neighbor counts, 58
  when one does). Rows 21 and 22, fixed at 70 pt: three lines "HHH" vertically distributed and
  justified beside the neighbor, rows 23 and 24 the same alone (the last line 64 or 67 from the
  top).
- ``column_style``. Column E styled Calibri 24 (``<col style>``), no cell in it. Row 4, fixed at
  32 pt, and row 5, fitted (no height): the probe. Whether a column's style counts in a row's
  bottom, and in a fitted row's height (31 pt if it counts, 15 if not).

Geometry is read for one cell a column (row 1) and one a row (column A, which gives the row's
height). ``read`` writes ``tests/captures/row_bottoms.tsv``: one row per case, the ink of each
line (``line_tops``, from the cell's top, a merge's from the merge's top) and the row's height,
measured as ``text_style`` measures (its code, by import), lengths in points at zoom 100.
``PREDICTIONS`` holds, for each case, the last baseline each candidate rule gives, written
before the run.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Font
from openpyxl.utils.cell import coordinate_from_string

from capture import text_style
from capture.text_style import Case

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "row_bottoms"
OUTPUT = ROOT / "tests" / "captures" / "row_bottoms.tsv"
GROUPS = ("cells", "column_style")
MARKERS = {"cells": ("B2", "I3"), "column_style": ("B2", "G3")}
"""The second marker sits just right of every case, in row 3."""
PROBE = "HHH"
ROW = 32
"""Every fixed row's height, but the spread rows'."""
SPREAD_ROW = 70
NEIGHBOR = Font(name="Calibri", size=24)
COLUMNS = (
    "group", "name", "zoom", "cell", "row_height", "lines", "line_tops", "baselines",
)  # fmt: skip

# The last baseline, from the cell's (or merge's) top, under each candidate rule. A fixed row of
# 32 pt: 29 with the probe's own bottom (3), 26 with the neighbor's (6); 28 with a 16 pt drawn
# size's (4). A 64 pt merge: 61 or 58. A 70 pt row: 67 or 64, or 61 when the wrapped lines also
# sit the difference of the line insets higher (Calibri 24's 4 against the probe's 1), as
# bottom-aligned wrapped text does (§3.6).
PREDICTIONS: dict[str, dict[str, float]] = {
    "probe alone": {"own": 29},
    "neighbor with a value": {"counted": 26, "not counted": 29},
    "empty styled neighbor": {"counted": 26, "not counted": 29},
    "neighbor in a hidden column": {"counted": 26, "not counted": 29},
    "neighbor at 45 degrees": {"counted": 26, "not counted": 29},
    "neighbor at 90 degrees": {"counted": 26, "not counted": 29},
    "stacked neighbor": {"counted": 26, "not counted": 29},
    "merge anchored above, first row": {"counted": 26, "not counted": 29},
    "merge anchored above, second row": {"counted": 26, "not counted": 29},
    "superscript neighbor": {"counted": 26, "drawn size": 28, "not counted": 29},
    "subscript neighbor": {"counted": 26, "drawn size": 28, "not counted": 29},
    "row style": {"counted": 26, "not counted": 29},
    "neighbor with a large run": {"counted": 26, "not counted": 29},
    "merged probe, neighbor in its first row": {
        "last row's": 61, "first row's": 58, "largest": 58,
    },
    "merged probe, neighbor in its last row": {
        "last row's": 58, "first row's": 61, "largest": 58,
    },
    "distributed beside the neighbor": {"counted": 64, "counted, inset": 61, "not counted": 67},
    "justify beside the neighbor": {"counted": 64, "counted, inset": 61, "not counted": 67},
    "distributed alone": {"own": 67},
    "justify alone": {"own": 67},
    "column style, fixed row": {"counted": 26, "not counted": 29},
    "column style, fitted row": {"counted": 25, "not counted": 12},
}  # fmt: skip
"""Each case's last baseline under each candidate. For the fitted row of ``column_style``, a
counted column style makes the row Calibri 24's 31 pt and puts the probe on 31 - 6."""
FITTED_HEIGHTS = {"column style, fitted row": {"counted": 31, "not counted": 15}}
"""The fitted row's height under each candidate."""
SHEETPEEK: dict[str, str] = {
    "probe alone": "own",
    "neighbor with a value": "counted",
    "empty styled neighbor": "counted",
    "neighbor in a hidden column": "not counted",
    "neighbor at 45 degrees": "counted",
    "neighbor at 90 degrees": "not counted",
    "stacked neighbor": "counted",
    "merge anchored above, first row": "counted",
    "merge anchored above, second row": "not counted",
    "superscript neighbor": "counted",
    "subscript neighbor": "counted",
    "row style": "not counted",
    "neighbor with a large run": "counted",
    "merged probe, neighbor in its first row": "last row's",
    "merged probe, neighbor in its last row": "last row's",
    "distributed beside the neighbor": "counted, inset",
    "justify beside the neighbor": "counted, inset",
    "distributed alone": "own",
    "justify alone": "own",
    "column style, fixed row": "not counted",
    "column style, fitted row": "not counted",
}
"""The candidate sheetpeek's layout takes for each case before the capture."""


def _probe(ws: Any, cell: str, *, vertical: str = "bottom", text: str = PROBE) -> None:
    ws[cell] = text
    ws[cell].font = Font(name="Calibri", size=11)
    ws[cell].alignment = Alignment(vertical=vertical, wrap_text="\n" in text)


def _neighbor(ws: Any, cell: str, **alignment: Any) -> None:
    ws[cell] = PROBE
    ws[cell].font = NEIGHBOR
    if alignment:
        ws[cell].alignment = Alignment(**alignment)


def _cells(ws: Any) -> list[Case]:
    found: list[Case] = []

    def case(name: str, cell: str, last: str | None = None) -> None:
        found.append(Case("cells", name, cell, ("lines",), last))

    for row in range(4, 21):
        ws.row_dimensions[row].height = ROW
    for row in range(21, 25):
        ws.row_dimensions[row].height = SPREAD_ROW
    ws.column_dimensions["G"].hidden = True
    _probe(ws, "C4")
    case("probe alone", "C4")
    _probe(ws, "C5")
    _neighbor(ws, "E5")
    case("neighbor with a value", "C5")
    _probe(ws, "C6")
    ws["E6"].font = NEIGHBOR
    case("empty styled neighbor", "C6")
    _probe(ws, "C7")
    _neighbor(ws, "G7")
    case("neighbor in a hidden column", "C7")
    for row, name, alignment in (
        (8, "neighbor at 45 degrees", {"text_rotation": 45}),
        (9, "neighbor at 90 degrees", {"text_rotation": 90}),
        (10, "stacked neighbor", {"text_rotation": 255}),
    ):
        _probe(ws, f"C{row}")
        _neighbor(ws, f"E{row}", **alignment)
        case(name, f"C{row}")
    _neighbor(ws, "E11")
    ws.merge_cells("E11:E12")
    _probe(ws, "C11")
    _probe(ws, "C12")
    case("merge anchored above, first row", "C11")
    case("merge anchored above, second row", "C12")
    for row, kind in ((13, "superscript"), (14, "subscript")):
        _probe(ws, f"C{row}")
        ws[f"E{row}"] = PROBE
        ws[f"E{row}"].font = Font(name="Calibri", size=24, vertAlign=kind)
        case(f"{kind} neighbor", f"C{row}")
    ws.row_dimensions[15].font = NEIGHBOR
    _probe(ws, "C15")
    case("row style", "C15")
    _probe(ws, "C16")
    ws["E16"] = CellRichText(
        TextBlock(InlineFont(rFont="Calibri", sz=11), "HH"),
        TextBlock(InlineFont(rFont="Calibri", sz=24), "HH"),
    )
    case("neighbor with a large run", "C16")
    for first, last, at, name in (
        (17, 18, 17, "merged probe, neighbor in its first row"),
        (19, 20, 20, "merged probe, neighbor in its last row"),
    ):
        _probe(ws, f"C{first}")
        ws.merge_cells(f"C{first}:C{last}")
        _neighbor(ws, f"E{at}")
        case(name, f"C{first}", f"C{last}")
    lines = "\n".join([PROBE] * 3)
    for row, vertical, neighbor in (
        (21, "distributed", True),
        (22, "justify", True),
        (23, "distributed", False),
        (24, "justify", False),
    ):
        _probe(ws, f"C{row}", vertical=vertical, text=lines)
        if neighbor:
            _neighbor(ws, f"E{row}")
        where = "beside the neighbor" if neighbor else "alone"
        case(f"{vertical} {where}", f"C{row}")
    return found


def _column_style(ws: Any) -> list[Case]:
    ws.column_dimensions["E"].font = NEIGHBOR
    ws.row_dimensions[4].height = ROW
    _probe(ws, "C4")
    _probe(ws, "C5")
    return [
        Case("column_style", "column style, fixed row", "C4", ("lines",)),
        Case("column_style", "column style, fitted row", "C5", ("lines",)),
    ]


def fill(ws: Any, group: str) -> list[Case]:
    """Lay one group out on the first worksheet of its own workbook; its cases."""
    return _cells(ws) if group == "cells" else _column_style(ws)


def cases(group: str | None = None) -> list[Case]:
    """Every case of ``group``, or of every group, in order."""
    chosen = GROUPS if group is None else (group,)
    return [case for g in chosen for case in fill(openpyxl.Workbook().active, g)]


def _boxes(found: list[Case]) -> list[str]:
    cells = [c for case in found for c in (case.cell, case.last) if c]
    return list(dict.fromkeys(cells))


def references(group: str) -> list[str]:
    """The cells whose geometry is read: row 1 for each column, column A for each row."""
    columns, rows = {}, {}
    for cell in _boxes(cases(group)):
        letters, row = coordinate_from_string(cell)
        columns[f"{letters}1"] = None
        rows[f"A{row}"] = None
    return [*columns, *rows]


def complete(cap: Any, cells: list[str]) -> None:
    """Give each of ``cells`` a box from its column's row-1 cell and its row's column-A cell."""
    for cell in cells:
        if cell in cap.geometry:
            continue
        letters, row = coordinate_from_string(cell)
        column, line = cap.geometry[f"{letters}1"], cap.geometry[f"A{row}"]
        cap.geometry[cell] = {
            "left": column["left"],
            "width": column["width"],
            "top": line["top"],
            "height": line["height"],
        }


def build(directory: Path) -> dict[str, Path]:
    """Write each group to its own workbook in ``directory``; the paths by group."""
    directory.mkdir(parents=True, exist_ok=True)
    paths = {}
    for group in GROUPS:
        wb = openpyxl.Workbook()
        wb.active.title = group
        fill(wb.active, group)
        paths[group] = directory / f"{group}.xlsx"
        wb.save(paths[group])
    return paths


def run(groups: tuple[str, ...] = GROUPS) -> None:
    """Capture each group's window, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    for group in groups:
        capture(
            lambda ws, group=group: fill(ws, group),
            cells=references(group),
            out=BUILD_DIR / group,
            markers=MARKERS[group],
            zooms=(100,),
            name=f"row_bottoms_{group}",
        )
    write()


def baselines(line_tops: str) -> list[float]:
    """Each line's baseline from ``line_tops`` ("top+height" spans): the bottom of its ink."""
    found = []
    for span in line_tops.split():
        top, height = span.split("+")
        found.append(float(top) + float(height))
    return found


def measure(cap: Any, case: Case) -> dict[str, str] | None:
    """One case's measurements, or ``None`` when it is outside the window."""
    complete(cap, _boxes([case]))
    found = text_style.measure(cap, case)
    if found is not None:
        found["baselines"] = " ".join(f"{b:g}" for b in baselines(found.get("line_tops", "")))
    return found


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``row_bottoms.tsv``."""
    from capture.screen.run import header, load

    records = []
    first: dict[str, str] = {}
    for group in GROUPS:
        out = BUILD_DIR / group
        if not (out / "formats.tsv").exists():
            print(f"{group}: not captured yet, left out", file=sys.stderr)
            continue
        first = first or header(out)
        found = fill(openpyxl.Workbook().active, group)
        for cap in load(out):
            for case in found:
                measured = measure(cap, case)
                if measured is None:
                    continue
                row = coordinate_from_string(case.cell)[1]
                records.append(
                    {
                        "group": case.group,
                        "name": case.name,
                        "zoom": str(cap.zoom),
                        "cell": case.cell,
                        "row_height": cap.value(f"A{row}", "geometry.height") or "",
                        **measured,
                    }
                )
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after", "zooms"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: row_bottoms\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        paths = build(Path(argv[1]))
        print(f"{len(paths)} workbooks in {argv[1]}")
        return 0
    if argv[:1] == ["run"] and all(group in GROUPS for group in argv[1:]):
        run(tuple(argv[1:]) or GROUPS)
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
