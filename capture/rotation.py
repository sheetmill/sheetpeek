"""Capture how Excel for Mac draws angled and stacked text, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.rotation run
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.rotation run more
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

A second set (``more_cases``), added after the first was captured, goes in windows of its own
(``more_<k>``; ``run more`` captures only those), so the first set's windows stay as captured:
- ``slant`` again, all four sides unless named: what the text's shift between slanted borders
  depends on. Angles 5 to 85 and -15 to -75; row heights 30, 60 and 120 at 15 and 45 degrees
  (60 at -45); the short and long texts; left, center and right; top and center in a 90 pt row;
  20 pt; indent 2; the left or the bottom side alone; column E 4 and 20 characters wide (cases of
  one width share a window). Measured: the ink box and the borders, as above.
- ``turned``: wrapped angled text in custom rows (45 degrees in 30, 60 and 120 pt; 15, 75, -45
  and 90 in 60 pt), in a fitted row, justified and distributed, and inside four borders.
  Measured: the ink box.
- ``upright``: the three texts upright at 9, 11, 14 and 20 pt (not the long one at 20), and the
  medium text at 90 degrees at 11, 14 and 20 pt: whether Excel measures upright text wider than
  xlformat's tables at the larger sizes too, as it does angled text. Measured: the ink box.
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
    width: float | None = None  # column E's width in characters, or the default
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


TURNED = ((45, 30), (45, 60), (45, 120), (15, 60), (75, 60), (-45, 60), (90, 60))
"""Angle and custom row height of the wrapped angled cases."""


def _slant(name: str, angle: int, **options: Any) -> Case:
    sides = options.pop("borders", "lrtb")
    return Case(
        "slant", name, options.pop("text", TEXT), _stored(angle), borders=sides,
        measure=("ink", "borders"), **options,
    )  # fmt: skip


def more_cases() -> list[Case]:
    """The second set: what the text's shift between slanted borders depends on, turned wrapping,
    and upright text at the sizes where Excel measures angled text wider than xlformat."""
    found = [_slant(f"{angle} lrtb", angle) for angle in (5, 30, 60, 75, 85, -15, -30, -60, -75)]
    found.extend(
        _slant(f"{angle} lrtb height {height}", angle, height=height)
        for angle in (15, 45)
        for height in (30, 60, 120)
    )
    found.append(_slant("-45 lrtb height 60", -45, height=60))
    found.extend(
        _slant(f"{angle} lrtb {length}", angle, text=LENGTHS[length])
        for angle in (15, 45)
        for length in ("short", "long")
    )
    found.extend(
        _slant(f"{angle} lrtb {horizontal}", angle, horizontal=horizontal)
        for angle in (15, 45)
        for horizontal in ("left", "center", "right")
    )
    found.extend(
        _slant(f"45 lrtb vertical {vertical}", 45, vertical=vertical, height=90)
        for vertical in ("top", "center")
    )
    found.extend(_slant(f"{angle} lrtb 20", angle, size=20.0) for angle in (15, 45))
    found.append(_slant("45 lrtb indent 2", 45, horizontal="left", indent=2))
    found.extend(_slant(f"45 {side}", 45, borders=side) for side in ("l", "b"))
    found.extend(
        _slant(f"{angle} lrtb width {width:g}", angle, width=width)
        for width in (4.0, 20.0)
        for angle in (15, 45)
    )
    found.extend(
        Case(
            "turned", f"wrapped {angle} height {height}", TEXT, _stored(angle), wrap=True,
            height=height,
        )
        for angle, height in TURNED
    )  # fmt: skip
    found.append(Case("turned", "wrapped 45 fitted", TEXT, 45, wrap=True))
    found.extend(
        Case("turned", f"{horizontal} 45 height 60", TEXT, 45, horizontal=horizontal, height=60)
        for horizontal in ("justify", "distributed")
    )
    found.append(
        Case("turned", "wrapped 45 lrtb height 60", TEXT, 45, wrap=True, height=60, borders="lrtb",
             measure=("ink", "borders"))
    )  # fmt: skip
    found.extend(
        Case("upright", f"{length} {size:g}", LENGTHS[length], 0, size=size)
        for size in (9.0, 11.0, 14.0, 20.0)
        for length in ("short", "medium", "long")
        if (length, size) != ("long", 20.0)
    )
    found.extend(
        Case("upright", f"90 ink medium {size:g}", TEXT, 90, size=size)
        for size in (11.0, 14.0, 20.0)
    )
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
        if case.width is not None:
            ws.column_dimensions[COLUMN].width = case.width
        placed.append((cell.coordinate, case))
    return placed


def build(path: Path) -> list[tuple[str, str, Case]]:
    """Write every case to one workbook at ``path``: the first set on one sheet, each window of
    the second on its own. Each case's sheet and cell address with the case."""
    wb = openpyxl.Workbook()
    wb.active.title = "Rotation"
    placed = [("Rotation", cell, case) for cell, case in fill(wb.active, cases())]
    for name, chosen in _windows(first=False):
        ws = wb.create_sheet(name)
        placed.extend((name, cell, case) for cell, case in fill(ws, chosen))
    wb.save(path)
    return placed


def _on_screen() -> tuple[list[list[Case]], list[Case]]:
    """The cases a window capture measures, in windows, and those Excel's reads settle."""
    every = cases()
    shown = [c for c in every if "ink" in c.measure or "borders" in c.measure]
    read = [c for c in every if c not in shown]
    windows = [shown[k : k + PER_WINDOW] for k in range(0, len(shown), PER_WINDOW)]
    return windows, read


def _more_windows() -> list[list[Case]]:
    """The second set in windows of its own, so that the first set's windows stay as captured;
    cases that set column E's width share a window only with cases of the same width."""
    windows: list[list[Case]] = []
    every = more_cases()
    for width in dict.fromkeys(case.width for case in every):
        same = [case for case in every if case.width == width]
        windows.extend(same[k : k + PER_WINDOW] for k in range(0, len(same), PER_WINDOW))
    return windows


def _windows(first: bool = True, more: bool = True) -> list[tuple[str, list[Case]]]:
    """Each window's folder name with its cases: the first set's, then the second's."""
    found = [(f"window_{k}", chosen) for k, chosen in enumerate(_on_screen()[0])] if first else []
    if more:
        found += [(f"more_{k}", chosen) for k, chosen in enumerate(_more_windows())]
    return found


def run(first: bool = True, more: bool = True) -> None:
    """Capture every window of cases and read the rest, then write the TSV (needs Excel)."""
    from capture.screen.run import capture, read_only

    for name, chosen in _windows(first, more):
        last = FIRST_ROW + 2 * (len(chosen) - 1)
        cells = [cell for cell, _ in fill(openpyxl.Workbook().active, chosen)]
        cells += [f"{c}{FIRST_ROW + 2 * i}" for i in range(len(chosen)) for c in "AI"]
        capture(
            lambda ws, chosen=chosen: fill(ws, chosen),
            cells=cells,
            out=BUILD_DIR / name,
            markers=("B2", f"J{last + 2}"),
            name=f"rotation_{name}",
        )
    if first:
        read = _on_screen()[1]
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

    read = _on_screen()[1]
    records = []
    first: dict[str, str] = {}
    for name, chosen in _windows():
        out = BUILD_DIR / name
        if not (out / "formats.tsv").exists():
            print(f"{name}: not captured yet, left out", file=sys.stderr)
            continue
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
    if argv == ["run", "more"]:
        run(first=False)
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
