"""Capture how Excel for Mac draws angled and stacked text, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.rotation run
    uv run --group capture python -m capture.rotation read     # the saved captures again
    uv run --group capture python -m capture.rotation build PATH   # the workbook only

Each case is one cell in column E, from row 4, with columns A to D and F to I empty so that text
can spill both ways, and an empty row between cases. Rows are fitted by Excel on opening (no
``ht``) unless a case sets a height. Cases measured on the screen go five to a window, each
window its own workbook and session, at zoom 100; the second marker sits in column J below the
last case. Cases measured by row height alone are read from Excel without a capture. ``read``
writes ``tests/captures/rotation.tsv``: one row a case, lengths in points from the cell's left
and top edges.

Cases (``ARCHITECTURE.md`` §4.17 marks what each one settles):
- ``slant``: which side borders make Excel slant the borders of angled text, and where the text
  goes inside them. All four sides, left and right only, top and bottom only, and a fill alone,
  at 15, 45, -45 and 90 degrees. Measured: the text's ink box, and the x of the dark border
  pixels nearest the cell's left and right edges at 25 % and 75 % of its height.
- ``anchor``: where general text switches from the left to the right of the cell between 0 and
  90 degrees, and between -90 and 0. Measured: the ink box.
- ``align``: fill, justify and distributed, an indent of 2 left and right, and top and center
  vertical alignment in a 90 pt row, at 45 degrees. Measured: the ink box.
- ``height``: fitted row heights for angled text of three lengths at four angles and four font
  sizes, and stacked text of three lengths at three sizes. Measured: the row height.
- ``other``: a number wider than its column at 45 and 90 degrees (``#`` or the number?), and
  wrapped angled text with a line feed. Measured: the ink box and what ``display format``
  reads.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

TEXT = "Rotated text here"
LENGTHS = {"short": "Rot", "medium": TEXT, "long": f"{TEXT}, and more of it"}
COLUMN = "E"
FIRST_ROW = 4
PER_WINDOW = 5
BLUE = "FF0000C0"
"""Text color of bordered cases, so that the text's ink and the borders' can be told apart."""
ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "rotation"
OUTPUT = ROOT / "tests" / "captures" / "rotation.tsv"
COLUMNS = (
    "group", "name", "cell", "rotation", "row_height", "ink_left", "ink_right", "ink_top",
    "ink_bottom", "border_left_25", "border_right_25", "border_left_75", "border_right_75",
    "orientation", "text",
)  # fmt: skip


@dataclass(frozen=True)
class Case:
    """One cell: its text, angle (``textRotation`` as stored) and what else it sets."""

    group: str
    name: str
    text: str
    rotation: int
    borders: str = ""  # "lrtb" letters for the sides drawn
    fill: bool = False
    horizontal: str | None = None
    vertical: str | None = None
    indent: int = 0
    size: float = 11.0
    wrap: bool = False
    height: float | None = None  # a custom row height, or fitted by Excel
    number: bool = False
    measure: tuple[str, ...] = field(default=("ink",))


def _stored(angle: int) -> int:
    """``textRotation`` for an angle of -90 to 90 degrees."""
    return angle if angle >= 0 else 90 - angle


def cases() -> list[Case]:
    """Every case, in row order."""
    found = [
        Case(
            "slant",
            f"{angle} {sides or 'fill'}",
            TEXT,
            _stored(angle),
            borders=sides,
            fill=not sides,
            measure=("ink", "borders"),
        )
        for angle in (15, 45, -45, 90)
        for sides in ("lrtb", "lr", "tb", "")
    ]
    found.extend(
        Case("anchor", f"{angle}", TEXT, _stored(angle))
        for angle in (30, 60, 75, 80, 85, 89, -15, -30, -60, -75, -85, -89)
    )
    found.extend(
        Case("align", horizontal, TEXT, 45, horizontal=horizontal)
        for horizontal in ("fill", "justify", "distributed")
    )
    found.extend(
        Case("align", f"indent {horizontal}", TEXT, 45, horizontal=horizontal, indent=2)
        for horizontal in ("left", "right")
    )
    found.extend(
        Case("align", vertical, TEXT, 45, vertical=vertical, height=90)
        for vertical in ("top", "center")
    )
    found.extend(
        Case(
            "height",
            f"{angle} {length} {size:g}",
            text,
            _stored(angle),
            size=size,
            measure=("height",),
        )
        for size in (9.0, 11.0, 14.0, 20.0)
        for angle in (15, 45, 75, 90)
        for length, text in LENGTHS.items()
    )
    found.extend(
        Case("height", f"stacked {text} {size:g}", text, 255, size=size, measure=("height",))
        for size in (9.0, 14.0, 20.0)
        for text in ("Abc", "Stack", "Stacked!")
    )
    found.extend(
        Case(
            "other", f"number {angle}", "12345678901", angle, number=True, measure=("ink", "format")
        )
        for angle in (45, 90)
    )
    found.append(Case("other", "wrapped", "two\nlines of text", 45, wrap=True))
    return found


def fill(ws: Any, chosen: list[Case]) -> list[tuple[str, Case]]:
    """Place ``chosen`` on a worksheet; each case's cell address with the case."""
    thin = Side(style="thin", color="FF000000")
    sides = {"l": "left", "r": "right", "t": "top", "b": "bottom"}
    placed = []
    for k, case in enumerate(chosen):
        row = FIRST_ROW + 2 * k
        cell = ws[f"{COLUMN}{row}"]
        cell.value = float(case.text) if case.number else case.text
        cell.alignment = Alignment(
            textRotation=case.rotation,
            horizontal=case.horizontal,
            vertical=case.vertical,
            indent=case.indent,
            wrap_text=case.wrap or None,
        )
        color = BLUE if case.borders else None
        if case.size != 11.0 or color:
            cell.font = Font(name="Calibri", size=case.size, color=color)
        if case.borders:
            cell.border = Border(**{sides[letter]: thin for letter in case.borders})
        if case.fill:
            cell.fill = PatternFill("solid", fgColor="FFFFE699")
        if case.height is not None:
            ws.row_dimensions[row].height = case.height
        placed.append((cell.coordinate, case))
    return placed


def build(path: Path) -> list[tuple[str, Case]]:
    """Write every case to one workbook at ``path``; each case's cell address with the case."""
    wb = openpyxl.Workbook()
    wb.active.title = "Rotation"
    placed = fill(wb.active, cases())
    wb.save(path)
    return placed


def _on_screen() -> tuple[list[list[Case]], list[Case]]:
    """The cases a window capture measures, in windows, and those Excel's reads settle."""
    every = cases()
    shown = [c for c in every if "ink" in c.measure or "borders" in c.measure]
    read = [c for c in every if c not in shown]
    windows = [shown[k : k + PER_WINDOW] for k in range(0, len(shown), PER_WINDOW)]
    return windows, read


def run() -> None:
    """Capture every window of cases and read the rest, then write the TSV (needs Excel)."""
    from capture.screen.run import capture, read_only

    windows, read = _on_screen()
    for k, chosen in enumerate(windows):
        last = FIRST_ROW + 2 * (len(chosen) - 1)
        cells = [cell for cell, _ in fill(openpyxl.Workbook().active, chosen)]
        cells += [f"{c}{FIRST_ROW + 2 * i}" for i in range(len(chosen)) for c in "AI"]
        capture(
            lambda ws, chosen=chosen: fill(ws, chosen),
            cells=cells,
            out=BUILD_DIR / f"window_{k}",
            markers=("B2", f"J{last + 2}"),
            name=f"rotation_{k}",
        )
    cells = [cell for cell, _ in fill(openpyxl.Workbook().active, read)]
    read_only(lambda ws: fill(ws, read), cells, BUILD_DIR / "heights", name="rotation_heights")
    write()


def _is_blue(rgb: tuple[int, int, int]) -> bool:
    r, g, b = rgb
    return b > 120 and r < 110 and g < 110


def _is_dark(rgb: tuple[int, int, int]) -> bool:
    from capture.screen.run import luma

    return luma(rgb) < 110 and not _is_blue(rgb)


def _measure(cap: Any, cell: str, case: Case) -> dict[str, str]:
    """The ink box and the borders' x at 25 % and 75 % of the row, in points from the cell."""
    from capture.screen.run import box

    row = cell[1:]
    left, right = cap.edge(cell, "left"), cap.edge(cell, "right")
    top, bottom = cap.edge(cell, "top"), cap.edge(cell, "bottom")
    x0, x1 = cap.edge(f"A{row}", "left") + 1, cap.edge(f"I{row}", "right") - 1
    found: dict[str, str] = {}
    test = _is_blue if case.borders else _is_dark
    ink = box(cap, x0, top + 1, x1, bottom - 1, test)
    if ink is not None:
        found["ink_left"] = f"{(ink[0] - left) / cap.scale:g}"
        found["ink_right"] = f"{(ink[2] + 1 - right) / cap.scale:g}"
        found["ink_top"] = f"{(ink[1] - top) / cap.scale:g}"
        found["ink_bottom"] = f"{(ink[3] + 1 - top) / cap.scale:g}"
    if "borders" in case.measure:
        for share in (25, 75):
            y = top + (bottom - top) * share / 100
            dark = [x for x in range(round(x0), round(x1)) if _is_dark(cap.pixel(x, y))]
            if dark:
                found[f"border_left_{share}"] = f"{(dark[0] - left) / cap.scale:g}"
                found[f"border_right_{share}"] = f"{(dark[-1] + 1 - left) / cap.scale:g}"
    return found


def write() -> None:
    """Measure the saved captures and reads again (no Excel) and write ``rotation.tsv``."""
    from capture.screen.run import header, load

    windows, read = _on_screen()
    records = []
    first: dict[str, str] = {}
    for k, chosen in enumerate(windows):
        out = BUILD_DIR / f"window_{k}"
        first = first or header(out)
        (cap,) = load(out)
        for cell, case in fill(openpyxl.Workbook().active, chosen):
            records.append(_record(cap.rows, cell, case, _measure(cap, cell, case)))
    rows = _reads(BUILD_DIR / "heights")
    for cell, case in fill(openpyxl.Workbook().active, read):
        records.append(_record(rows, cell, case, {}))
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: rotation\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} cases written to {OUTPUT}")


def _reads(out: Path) -> list[tuple[str, str, str]]:
    lines = (out / "formats.tsv").read_text(encoding="utf-8").splitlines()
    return [
        (*[*line.split("\t"), "", ""][:3],)  # type: ignore[misc]
        for line in lines
        if not line.startswith("# ") and not line.startswith("cell\t")
    ]


def _record(
    rows: list[tuple[str, str, str]], cell: str, case: Case, measured: dict[str, str]
) -> dict[str, str]:
    values = {(c, k): v for c, k, v in rows if c == cell}
    return {
        "group": case.group,
        "name": case.name,
        "cell": cell,
        "rotation": str(case.rotation),
        "row_height": values.get((cell, "geometry.height"), ""),
        "orientation": values.get((cell, "display.orientation"), ""),
        "text": values.get((cell, "geometry.text"), ""),
        **measured,
    }


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cases in {argv[1]}")
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
