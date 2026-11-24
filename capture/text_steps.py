"""Capture indents under other Normal fonts and shrink to fit by single characters (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.text_steps run
    uv run --group capture python -m capture.text_steps run shrink   # some groups only
    uv run --group capture python -m capture.text_steps read     # the saved captures again
    uv run --group capture python -m capture.text_steps build DIR    # the workbooks only

``text_style`` and ``text_style_more`` measured indents at a Calibri 11 Normal font and shrink to
fit at a few lengths; this suite takes both a step at a time. ARCHITECTURE.md §4.6 (indent) and
§4.13 (shrink to fit). Each group is its own workbook and window at zoom 100, in the dark
appearance as the other suites; its Normal font is the group's.

- ``indent_<face>``, one group per Normal font: Calibri 9, 11, 12, 16 and 24, Aptos Narrow 11
  and 12. Rows 4 to 9: "HHHH" at indent 0 to 5, left aligned in C and right aligned in E
  (columns of 20 characters): the ink's left edge (left) and right edge (right), so how far a
  level moves the text under each Normal font. ``indent_calibri11`` also has the same cells in
  a Calibri 16 cell font (G, I) and an Aptos Narrow 12 one (K, M): whether a level follows the
  Normal font or the cell's.
- ``shrink``, Calibri 11. Rows 4 to 30: "H" repeated as many times as the row's number, left
  aligned with shrink to fit in the default column C, and unshrunk in E (40 characters): the
  ink's width over the reference's, at each length. Rows 4 to 24: "HHHHHHHHHH" unshrunk in G at
  1 to 11 pt in half points: the ink width of each drawn size, to tell which size a shrunk text
  is drawn at.
- ``shrink_calibri16`` and ``shrink_aptos_narrow11``, those Normal fonts. "H" repeated 4 to 51
  times with shrink to fit, left aligned, in the default columns C (4 to 19), E (20 to 35) and G
  (36 to 51), rows 4 to 19; unshrunk "H" repeated 10 (I4) and 20 times (I5) in a column of 60
  characters: each length's ink width against the unshrunk advance, so the size it is drawn at.

Geometry is read for one cell a column (row 1) and one a row (column A). ``read`` writes
``tests/captures/text_steps.tsv``: one row per case, measured as ``text_style`` measures (its
code, by import), lengths in points at zoom 100.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.utils.indexed_list import IndexedList

from capture import text_style
from capture.text_style import Case

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "text_steps"
OUTPUT = ROOT / "tests" / "captures" / "text_steps.tsv"

FACES = {
    "indent_calibri9": ("Calibri", 9),
    "indent_calibri11": ("Calibri", 11),
    "indent_calibri12": ("Calibri", 12),
    "indent_calibri16": ("Calibri", 16),
    "indent_calibri24": ("Calibri", 24),
    "indent_aptos_narrow11": ("Aptos Narrow", 11),
    "indent_aptos_narrow12": ("Aptos Narrow", 12),
    "shrink": ("Calibri", 11),
    "shrink_calibri16": ("Calibri", 16),
    "shrink_aptos_narrow11": ("Aptos Narrow", 11),
}
"""Each group's Normal font."""
GROUPS = tuple(FACES)
MARKERS = {
    group: (
        "B2",
        "O3"
        if group == "indent_calibri11"
        else "I3"
        if group == "shrink"
        else "K3"
        if group.startswith("shrink_")
        else "G3",
    )
    for group in GROUPS
}
"""The second marker sits just right of every case, in row 3, so that it stays in the window
beside the wide columns of a large Normal font."""
INDENTS = range(6)
CELL_FONTS = (("G", "I", "Calibri", 16), ("K", "M", "Aptos Narrow", 12))
"""``indent_calibri11``'s cell-font columns: left, right, face, size."""
LENGTHS = range(4, 31)
"""Shrunk texts' lengths in characters; each in the row of its number."""
SHRUNK_COLUMNS = (("C", range(4, 20)), ("E", range(20, 36)), ("G", range(36, 52)))
"""``shrink_<face>``: the default columns holding "H" repeated 4 to 51 times, 16 lengths each."""
REFERENCES = {"I4": 10, "I5": 20}
"""``shrink_<face>``: unshrunk "H" repeated 10 and 20 times in a column of 60 characters: their
ink widths give the face's "H" advance and side bearings, so each shrunk length's drawn size."""
LADDER = tuple(1 + k / 2 for k in range(21))
"""The reference ladder's sizes, 1 to 11 pt, in rows 4 to 24."""
COLUMNS = (
    "group", "name", "zoom", "cell", "normal_font", "cell_font", "align", "indent",
    "row_height", "cell_width", "ink_left", "ink_right", "ink_top", "ink_bottom", "width_ratio",
    "height_ratio", "ink_width", "ref_width",
)  # fmt: skip


def _normal(ws: Any, group: str) -> None:
    """Make the group's face the workbook's Normal font (before any cell is styled)."""
    name, size = FACES[group]
    normal = Font(name=name, sz=size, family=2, scheme="minor" if name == "Calibri" else None)
    wb = ws.parent
    wb._fonts = IndexedList([normal])
    for style in wb._named_styles:
        if style.name == "Normal":
            style.font = normal


def _indent(ws: Any, group: str) -> list[Case]:
    name, size = FACES[group]
    columns = [("C", "E", name, size, False)]
    if group == "indent_calibri11":
        columns += [(left, right, face, sz, True) for left, right, face, sz in CELL_FONTS]
    found = []
    for left, right, face, sz, own in columns:
        ws.column_dimensions[left].width = ws.column_dimensions[right].width = 20
        for row, level in zip(range(4, 10), INDENTS, strict=True):
            for column, align in ((left, "left"), (right, "right")):
                cell = ws[f"{column}{row}"]
                cell.value = "HHHH"
                cell.alignment = Alignment(horizontal=align, indent=level)
                if own:
                    cell.font = Font(name=face, size=sz)
                label = f"{face} {sz} {align} indent {level}"
                found.append(Case(group, label, f"{column}{row}", ("ink",)))
    return found


def _shrink(ws: Any) -> list[Case]:
    ws.column_dimensions["E"].width = 40
    ws.column_dimensions["G"].width = 20
    found = []
    for length in LENGTHS:
        shrunk, reference = f"C{length}", f"E{length}"
        ws[shrunk] = ws[reference] = "H" * length
        ws[shrunk].alignment = Alignment(horizontal="left", shrink_to_fit=True)
        ws[reference].alignment = Alignment(horizontal="left")
        found.append(Case("shrink", f"shrunk {length}", shrunk, ("scale", "ink"), None, reference))
    for row, size in enumerate(LADDER, 4):
        ws[f"G{row}"] = "H" * 10
        ws[f"G{row}"].font = Font(name="Calibri", size=size)
        ws[f"G{row}"].alignment = Alignment(horizontal="left")
        found.append(Case("shrink", f"ladder {size:g}", f"G{row}", ("ink",)))
    return found


def _shrink_face(ws: Any, group: str) -> list[Case]:
    found = []
    for column, lengths in SHRUNK_COLUMNS:
        for row, length in enumerate(lengths, 4):
            cell = f"{column}{row}"
            ws[cell] = "H" * length
            ws[cell].alignment = Alignment(horizontal="left", shrink_to_fit=True)
            found.append(Case(group, f"shrunk {length}", cell, ("ink",)))
    ws.column_dimensions["I"].width = 60
    for cell, length in REFERENCES.items():
        ws[cell] = "H" * length
        ws[cell].alignment = Alignment(horizontal="left")
        found.append(Case(group, f"reference {length}", cell, ("ink",)))
    return found


def fill(ws: Any, group: str) -> list[Case]:
    """Lay one group out on the first worksheet of its own workbook; its cases."""
    _normal(ws, group)
    if group == "shrink":
        return _shrink(ws)
    if group.startswith("shrink_"):
        return _shrink_face(ws, group)
    return _indent(ws, group)


def cases(group: str | None = None) -> list[Case]:
    """Every case of ``group``, or of every group, in order."""
    chosen = GROUPS if group is None else (group,)
    return [case for g in chosen for case in fill(openpyxl.Workbook().active, g)]


def _boxes(found: list[Case]) -> list[str]:
    cells = [c for case in found for c in (case.cell, case.last, case.reference) if c]
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
            name=f"text_steps_{group}",
        )
    write()


def _described(case: Case, ws: Any) -> dict[str, str]:
    cell = ws[case.cell]
    return {
        "normal_font": "{} {:g}".format(*FACES[case.group]),
        "cell_font": f"{cell.font.name} {cell.font.size:g}",
        "align": cell.alignment.horizontal or "",
        "indent": f"{cell.alignment.indent:g}",
    }


def measure(cap: Any, case: Case) -> dict[str, str] | None:
    """One case's measurements, or ``None`` when it is outside the window."""
    complete(cap, _boxes([case]))
    return text_style.measure(cap, case)


def group_records(cap: Any, group: str) -> list[dict[str, str]]:
    """Every case of ``group`` measured in one capture of its window, in ``COLUMNS``' keys."""
    ws = openpyxl.Workbook().active
    records = []
    for case in fill(ws, group):
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
                "cell_width": f"{cap.geometry[case.cell]['width']:g}",
                **_described(case, ws),
                **measured,
            }
        )
    return records


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``text_steps.tsv``."""
    from capture.screen.run import header, load

    records = []
    first: dict[str, str] = {}
    for group in GROUPS:
        out = BUILD_DIR / group
        if not (out / "formats.tsv").exists():
            print(f"{group}: not captured yet, left out", file=sys.stderr)
            continue
        first = first or header(out)
        for cap in load(out):
            records += group_records(cap, group)
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after", "zooms"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: text_steps\n" + "\t".join(COLUMNS) + "\n")
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
