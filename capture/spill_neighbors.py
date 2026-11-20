"""Capture which neighbors stop Excel's text spill, and how numbers in other fonts overflow.

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.spill_neighbors run
    uv run --group capture python -m capture.spill_neighbors read        # the saved captures again
    uv run --group capture python -m capture.spill_neighbors build PATH  # every group's workbook

ARCHITECTURE.md §4.1 ("a formula returning ``""``, a single space or a formula returning ``0``
stops it too"), §4.2 (``####`` for cells whose font differs from the Normal font) and §4.18 (a
hidden zero's cell still blocks spill). Each group is its own workbook and window, at zoom 100,
in the dark appearance as the other suites, Normal font Calibri 11. Every workbook carries the
calculation properties Excel writes (``calcId`` 191029, ``<fileVersion appName="xl">``), so that
Excel and sheetpeek both show the cached values; each neighbor cell is written as Excel saves it
(``patch_part`` replaces the cell openpyxl wrote):

- ``blockers``: C4 to C12 hold the same long text, its neighbor D empty (the control: the text
  spills to F), the number 1 (the captured blocker), ``=""`` as Excel saves it (``t="str"``,
  ``<f>""</f><v></v>``), a text of one space, and ``=0`` with its cached 0.
- ``zeros``: the sheet's view has ``showZeros="0"``. C4 to C8 hold the long text, D empty (the
  control), the number 0 and ``=0``; both zeros show nothing.
- ``numbers``: numbers of 7 to 12 digits (1234567 to 123456789012), General, in Calibri 11 bold,
  Calibri 9, 14 and 20 and Arial 11, one font a row (rows 4 to 12), in columns C to H at the
  default width and J to O at 22 points (width 3.7109375 in the file, as ``layout`` column N).
  Rows 16 and 18 hold the long text beside ``=""`` written without a cached value (Excel
  computes it on opening) and beside an empty D (the control). It lives here, the group without
  any other formula, so that a recalculation it causes changes nothing else.

Measured, in points from the cell's edges, for the spilling texts: per column from C to H
whether it holds ink (``ink_columns``, ``1`` where it does, luma under 150), per boundary from
C|D to G|H whether its gridline is hidden (``hidden_gridlines``, ``1`` where no gridline gray
shows between 1 and 3 points below the row's top), and where the ink in C ends against C's right
edge (``ink_right``, 0 or a little less where the text is cut at C's edge); Excel's text of C
and of D (``text``, ``neighbor_text``). For the numbers: Excel's text (``#`` runs, reductions),
the row's height, and the ink's left and right against the cell's left and right edges.
Columns E to H are read once per group, in the group's first row. ``read`` writes
``tests/captures/spill_neighbors.tsv``, one row per group and case.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "spill_neighbors"
OUTPUT = ROOT / "tests" / "captures" / "spill_neighbors.tsv"
GROUPS = ("blockers", "zeros", "numbers")
MARKERS = {"blockers": ("B2", "J14"), "zeros": ("B2", "J10"), "numbers": ("B2", "Q20")}
ZOOM = 100
LONG = "a long text that spills over three columns"
"""Calibri 11, about 190 points: from C it ends inside F when nothing stops it."""
SPILL_COLUMNS = "CDEFGH"
NARROW = 3.7109375
"""The file width of a 22 point column (three digits in Calibri 11, ``layout`` column N)."""
DIGITS = (7, 8, 9, 10, 11, 12)
FONTS = {
    "Calibri 11 bold": {"name": "Calibri", "size": 11, "bold": True},
    "Calibri 9": {"name": "Calibri", "size": 9},
    "Calibri 14": {"name": "Calibri", "size": 14},
    "Calibri 20": {"name": "Calibri", "size": 20},
    "Arial 11": {"name": "Arial", "size": 11},
}
NEIGHBORS = {
    "empty": None,
    "number 1": '<c r="{ref}"><v>1</v></c>',
    'saved =""': '<c r="{ref}" t="str"><f>""</f><v></v></c>',
    "one space": '<c r="{ref}" t="inlineStr"><is><t xml:space="preserve"> </t></is></c>',
    "=0": '<c r="{ref}"><f>0</f><v>0</v></c>',
    "number 0": '<c r="{ref}"><v>0</v></c>',
    'uncached =""': '<c r="{ref}"><f>""</f></c>',
}
"""Each neighbor of the long text as its cell's XML (``None``: no cell)."""
COLUMNS = (
    "group", "name", "zoom", "cell", "font", "width", "text", "neighbor_text", "row_height",
    "ink_left", "ink_right", "ink_columns", "hidden_gridlines",
)  # fmt: skip
READS = {"text": "geometry.text", "row_height": "geometry.height"}


@dataclass(frozen=True)
class Case:
    """One measured cell: a long text beside a neighbor, or a number in a font."""

    group: str
    name: str
    cell: str
    neighbor: str | None = None
    """The key of :data:`NEIGHBORS` in the next column (spill cases only)."""
    font: str | None = None
    """The key of :data:`FONTS` (number cases only)."""
    digits: int = 0
    narrow: bool = False
    """In a 22 point column (number cases only)."""

    @property
    def spill(self) -> bool:
        return self.neighbor is not None

    @property
    def row(self) -> int:
        return int(re.sub(r"[A-Z]", "", self.cell))

    @property
    def neighbor_cell(self) -> str:
        return f"D{self.row}"

    @property
    def value(self) -> int:
        return int("123456789012"[: self.digits])


def _spill(group: str, row: int, neighbor: str) -> Case:
    return Case(group, f"beside {neighbor}", f"C{row}", neighbor=neighbor)


def _cases() -> list[Case]:
    found = [
        _spill("blockers", 4 + 2 * k, n)
        for k, n in enumerate(("empty", "number 1", 'saved =""', "one space", "=0"))
    ]
    found += [_spill("zeros", 4 + 2 * k, n) for k, n in enumerate(("empty", "number 0", "=0"))]
    for k, font in enumerate(FONTS):
        for narrow, letters in ((False, "CDEFGH"), (True, "JKLMNO")):
            for digits, letter in zip(DIGITS, letters, strict=True):
                width = "22 pt" if narrow else "default"
                found.append(
                    Case(
                        "numbers",
                        f"{digits} digits, {font}, {width}",
                        f"{letter}{4 + 2 * k}",
                        font=font,
                        digits=digits,
                        narrow=narrow,
                    )
                )
    found += [_spill("numbers", 16, 'uncached =""'), _spill("numbers", 18, "empty")]
    return found


CASES = _cases()


def cases(group: str | None = None) -> list[Case]:
    """Every case of ``group``, or of every group, in order."""
    return [c for c in CASES if group is None or c.group == group]


def _spilled(neighbor_text: str, ink_columns: str, hidden_gridlines: str) -> dict[str, Any]:
    return {
        "text": LONG,
        "neighbor_text": neighbor_text,
        "ink_columns": ink_columns,
        "hidden_gridlines": hidden_gridlines,
    }


EXPECTED: dict[str, dict[str, dict[str, Any]]] = {
    "blockers": {
        "C4": _spilled("", "111100", "11100"),
        "C6": _spilled("1", "110000", "00000"),
        "C8": _spilled("", "100000", "00000"),
        "C10": _spilled(" ", "100000", "00000"),
        "C12": _spilled("0", "110000", "00000"),
    },
    "zeros": {
        "C4": _spilled("", "111100", "11100"),
        "C6": _spilled("", "100000", "00000"),
        "C8": _spilled("", "100000", "00000"),
    },
    "numbers": {
        "C4": {"text": "1234567", "row_height": 15, "left": 9, "right": -2},
        "D4": {"text": "12345678", "row_height": 15, "left": 3, "right": -2},
        "E4": {"text": "1.23E+08", "row_height": 15, "left": 8, "right": -2},
        "F4": {"text": "1.23E+09", "row_height": 15, "left": 8, "right": -2},
        "G4": {"text": "1.23E+10", "row_height": 15, "left": 8, "right": -2},
        "H4": {"text": "1.23E+11", "row_height": 15, "left": 8, "right": -2},
        "J4": {"text": "###", "row_height": 15, "left": 4, "right": -3},
        "K4": {"text": "###", "row_height": 15, "left": 4, "right": -3},
        "L4": {"text": "###", "row_height": 15, "left": 4, "right": -3},
        "M4": {"text": "###", "row_height": 15, "left": 4, "right": -3},
        "N4": {"text": "###", "row_height": 15, "left": 4, "right": -3},
        "O4": {"text": "###", "row_height": 15, "left": 4, "right": -3},
        "C6": {"text": "1234567", "row_height": 15, "left": 16, "right": -2},
        "D6": {"text": "12345678", "row_height": 15, "left": 11, "right": -2},
        "E6": {"text": "123456789", "row_height": 15, "left": 6, "right": -2},
        "F6": {"text": "1.235E+09", "row_height": 15, "left": 11, "right": -2},
        "G6": {"text": "1.235E+10", "row_height": 15, "left": 11, "right": -2},
        "H6": {"text": "1.235E+11", "row_height": 15, "left": 11, "right": -2},
        "J6": {"text": "####", "row_height": 15, "left": 4, "right": -2},
        "K6": {"text": "####", "row_height": 15, "left": 4, "right": -2},
        "L6": {"text": "####", "row_height": 15, "left": 4, "right": -2},
        "M6": {"text": "####", "row_height": 15, "left": 4, "right": -2},
        "N6": {"text": "####", "row_height": 15, "left": 4, "right": -2},
        "O6": {"text": "####", "row_height": 15, "left": 4, "right": -2},
        "C8": {"text": "1E+06", "row_height": 19, "left": 16, "right": -2},
        "D8": {"text": "1E+07", "row_height": 19, "left": 16, "right": -2},
        "E8": {"text": "1E+08", "row_height": 19, "left": 16, "right": -2},
        "F8": {"text": "1E+09", "row_height": 19, "left": 16, "right": -2},
        "G8": {"text": "1E+10", "row_height": 19, "left": 16, "right": -2},
        "H8": {"text": "1E+11", "row_height": 19, "left": 16, "right": -2},
        "J8": {"text": "##", "row_height": 19, "left": 5, "right": -3},
        "K8": {"text": "##", "row_height": 19, "left": 5, "right": -3},
        "L8": {"text": "##", "row_height": 19, "left": 5, "right": -3},
        "M8": {"text": "##", "row_height": 19, "left": 5, "right": -3},
        "N8": {"text": "##", "row_height": 19, "left": 5, "right": -3},
        "O8": {"text": "##", "row_height": 19, "left": 5, "right": -3},
        "C10": {"text": "####", "row_height": 26, "left": 7, "right": -6},
        "D10": {"text": "####", "row_height": 26, "left": 7, "right": -6},
        "E10": {"text": "####", "row_height": 26, "left": 7, "right": -6},
        "F10": {"text": "####", "row_height": 26, "left": 7, "right": -6},
        "G10": {"text": "####", "row_height": 26, "left": 7, "right": -6},
        "H10": {"text": "####", "row_height": 26, "left": 7, "right": -6},
        "J10": {"text": "#", "row_height": 26, "left": 7, "right": -5},
        "K10": {"text": "#", "row_height": 26, "left": 7, "right": -5},
        "L10": {"text": "#", "row_height": 26, "left": 7, "right": -5},
        "M10": {"text": "#", "row_height": 26, "left": 7, "right": -5},
        "N10": {"text": "#", "row_height": 26, "left": 7, "right": -5},
        "O10": {"text": "#", "row_height": 26, "left": 7, "right": -5},
        "C12": {"text": "1234567", "row_height": 15, "left": 9, "right": -2},
        "D12": {"text": "12345678", "row_height": 15, "left": 3, "right": -2},
        "E12": {"text": "1.23E+08", "row_height": 15, "left": 5, "right": -2},
        "F12": {"text": "1.23E+09", "row_height": 15, "left": 5, "right": -2},
        "G12": {"text": "1.23E+10", "row_height": 15, "left": 5, "right": -2},
        "H12": {"text": "1.23E+11", "row_height": 15, "left": 5, "right": -2},
        "J12": {"text": "##", "row_height": 15, "left": 6, "right": -4},
        "K12": {"text": "##", "row_height": 15, "left": 6, "right": -4},
        "L12": {"text": "##", "row_height": 15, "left": 6, "right": -4},
        "M12": {"text": "##", "row_height": 15, "left": 6, "right": -4},
        "N12": {"text": "##", "row_height": 15, "left": 6, "right": -4},
        "O12": {"text": "##", "row_height": 15, "left": 6, "right": -4},
        "C16": _spilled("", "100000", "00000"),
        "C18": _spilled("", "111100", "11100"),
    },
}  # fmt: skip
"""What sheetpeek draws for each case, by group and cell, from its ``excel`` layout.

A spilling text: C's and D's text, ``ink_columns`` (C to H: the run's visible part, or D's own
text) and ``hidden_gridlines`` (C|D to G|H: ``RenderedRow.crossed``). A number: its text, the
row's height, and the left and right of its advance against the cell's left and right edges.
"""


def fill(ws: Any, group: str) -> list[Case]:
    """Lay one group out on a worksheet, each neighbor a placeholder; its cases.

    :func:`postprocess` then writes each neighbor cell as Excel saves it.
    """
    found = cases(group)
    if group == "zeros":
        ws.sheet_view.showZeros = False
    if group == "numbers":
        for letter in "JKLMNO":
            ws.column_dimensions[letter].width = NARROW
    for case in found:
        cell = ws[case.cell]
        if case.spill:
            cell.value = LONG
            if NEIGHBORS[case.neighbor or ""] is not None:
                ws[case.neighbor_cell] = "placeholder"
        else:
            cell.value = case.value
            cell.font = Font(**FONTS[case.font or ""])
    return found


FILE_VERSION = '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="11011"/>'
"""``<fileVersion>`` as Excel writes it (``tests/files/excel_formulas.xlsx``'s). With ``appName``
alone, Excel 16.115 refused this suite's workbooks (-50 with alerts off)."""


def excel_saved(text: str) -> str:
    """The workbook part with the calculation properties and file version Excel writes."""
    text = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="191029"/>', text)
    if "<fileVersion" not in text:
        text = text.replace("<workbookPr", FILE_VERSION + "<workbookPr", 1)
    return text


def neighbors(text: str, group: str) -> str:
    """A worksheet's XML with each case's neighbor cell written as :data:`NEIGHBORS` holds it."""
    for case in cases(group):
        xml = NEIGHBORS[case.neighbor] if case.neighbor is not None else None
        if xml is None:
            continue
        ref = case.neighbor_cell
        text, count = re.subn(rf'<c r="{ref}"[^>]*?(/>|>.*?</c>)', xml.format(ref=ref), text)
        if count != 1:
            raise ValueError(f"{ref}: no placeholder cell")
    return text


def postprocess(group: str) -> Any:
    """Write a group's neighbor cells as Excel saves them, with Excel's calculation properties."""

    def edit(path: Path) -> None:
        from capture.screen.run import patch_part

        patch_part(
            path,
            {
                "xl/worksheets/sheet1.xml": lambda t: neighbors(t, group),
                "xl/workbook.xml": excel_saved,
            },
        )

    return edit


def _first_row(group: str) -> int:
    return min(c.row for c in cases(group) if c.spill)


def _cells(group: str) -> list[str]:
    """The cells read: every case's, each spill case's neighbor, and E to H once."""
    found = []
    for case in cases(group):
        found += [case.cell, case.neighbor_cell] if case.spill else [case.cell]
    found += [f"{letter}{_first_row(group)}" for letter in SPILL_COLUMNS[2:]]
    return list(dict.fromkeys(found))


def build(path: Path) -> list[Case]:
    """Write every group to a workbook of its own in the folder ``path``; every case."""
    path.mkdir(parents=True, exist_ok=True)
    found = []
    for group in GROUPS:
        wb = openpyxl.Workbook()
        found.extend(fill(wb.active, group))
        target = path / f"spill_neighbors_{group}.xlsx"
        wb.save(target)
        postprocess(group)(target)
    return found


def run(groups: tuple[str, ...] | None = None) -> None:
    """Capture each group's window, then write the TSV (needs Excel)."""
    from capture.screen.run import capture

    for group in groups or GROUPS:
        capture(
            lambda ws, group=group: fill(ws, group),
            cells=_cells(group),
            out=BUILD_DIR / group,
            markers=MARKERS[group],
            zooms=(ZOOM,),
            name=f"spill_neighbors_{group}",
            postprocess=postprocess(group),
        )
    write()


# -- measuring -----------------------------------------------------------------------------------


def column_box(cap: Any, case: Case, letter: str) -> tuple[int, int, int, int]:
    """Column ``letter``'s box in the case's row, image pixels, end exclusive, inset by a pixel.

    The column's left and right come from its cell in the group's first spill row (read once),
    the top and bottom from the case's own cell.
    """
    reference = f"{letter}{_first_row(case.group)}"
    if f"{letter}{case.row}" in cap.geometry:
        reference = f"{letter}{case.row}"
    return (
        round(cap.edge(reference, "left")) + 1,
        round(cap.edge(case.cell, "top")) + 1,
        round(cap.edge(reference, "right")) - 1,
        round(cap.edge(case.cell, "bottom")) - 1,
    )


def ink_columns(cap: Any, case: Case) -> str:
    """Per column from C to H, ``1`` where it holds ink."""
    from capture.screen.ink import ink

    return "".join(
        "0" if ink(cap, column_box(cap, case, letter)) is None else "1" for letter in SPILL_COLUMNS
    )


def hidden_gridlines(cap: Any, case: Case) -> str:
    """Per boundary from C|D to G|H, ``1`` where no gridline gray shows near the row's top."""
    from capture.screen.ink import nearest

    gray = cap.gridline
    top = cap.edge(case.cell, "top")
    ys = range(round(top + 1 * cap.scale), round(top + 3 * cap.scale))
    flags = []
    for letter in SPILL_COLUMNS[:-1]:
        x = column_box(cap, case, letter)[2] + 1
        shows = gray is not None and any(
            nearest(cap.image.pixel(px, y), {"grid": gray}, 20) == "grid"
            for y in ys
            for px in range(x - 1, x + 2)
        )
        flags.append("0" if shows else "1")
    return "".join(flags)


def measure(cap: Any, case: Case) -> dict[str, str]:
    """One case's measurements at one zoom."""
    from capture.screen.ink import cells_box, ink, pt

    left, right = cap.edge(case.cell, "left"), cap.edge(case.cell, "right")
    found: dict[str, str] = {}
    glyph = ink(cap, cells_box(cap, case.cell, inset=1))
    if glyph is not None:
        found["ink_left"] = pt(cap, glyph[0] - left)
        found["ink_right"] = pt(cap, glyph[2] - right)
    if case.spill:
        found["ink_columns"] = ink_columns(cap, case)
        found["hidden_gridlines"] = hidden_gridlines(cap, case)
    return found


def record(cap: Any, case: Case) -> dict[str, str]:
    """A case's TSV row at one zoom: its identity, Excel's reads and the measurements."""
    reads = {key: cap.value(case.cell, read) or "" for key, read in READS.items()}
    if case.spill:
        reads["neighbor_text"] = cap.value(case.neighbor_cell, "geometry.text") or ""
    return {
        "group": case.group,
        "name": case.name,
        "zoom": str(cap.zoom),
        "cell": case.cell,
        "font": case.font or "Calibri 11",
        "width": "22 pt" if case.narrow else "default",
        **reads,
        **measure(cap, case),
    }


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``spill_neighbors.tsv``."""
    from capture.screen.run import header, load

    records = []
    first: dict[str, str] = {}
    for group in GROUPS:
        out = BUILD_DIR / group
        if not (out / "formats.tsv").exists():
            print(f"{group}: not captured yet, left out", file=sys.stderr)
            continue
        first = first or header(out)
        records += [record(cap, case) for cap in load(out) for case in cases(group)]
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: spill_neighbors\n" + "\t".join(COLUMNS) + "\n")
        for found in records:
            out.write("\t".join(found.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cases in {argv[1]}")
        return 0
    if argv[:1] == ["run"] and argv[1:] and all(group in GROUPS for group in argv[1:]):
        run(tuple(argv[1:]))
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
