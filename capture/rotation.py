"""Capture how Excel for Mac draws angled and stacked text, from the screen (route C).

    uv run --group capture python -m capture.rotation build PATH   # the workbook only, no Excel

The workbook holds one case a row, in column E, with columns A to D and F to H empty so that
text can spill both ways, and an empty row between cases. Rows are fitted by Excel on opening
(no ``ht``) unless a case sets a height. ``run`` and ``read`` come with the screen tool: each
case is measured from the window capture at zoom 100 as ``Case.measure`` says.

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

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

TEXT = "Rotated text here"
LENGTHS = {"short": "Rot", "medium": TEXT, "long": f"{TEXT}, and more of it"}
COLUMN = "E"


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


def build(path: Path) -> list[tuple[str, Case]]:
    """Write the workbook to ``path``; each case's cell address with the case."""
    thin = Side(style="thin", color="FF000000")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rotation"
    placed = []
    for k, case in enumerate(cases()):
        row = 2 + 2 * k
        cell = ws[f"{COLUMN}{row}"]
        cell.value = float(case.text) if case.number else case.text
        cell.alignment = Alignment(
            textRotation=case.rotation,
            horizontal=case.horizontal,
            vertical=case.vertical,
            indent=case.indent,
            wrap_text=case.wrap or None,
        )
        if case.size != 11.0:
            cell.font = Font(name="Calibri", size=case.size)
        if case.borders:
            sides = {"l": "left", "r": "right", "t": "top", "b": "bottom"}
            cell.border = Border(**{sides[letter]: thin for letter in case.borders})
        if case.fill:
            cell.fill = PatternFill("solid", fgColor="FFFFE699")
        if case.height is not None:
            ws.row_dimensions[row].height = case.height
        placed.append((cell.coordinate, case))
    wb.save(path)
    return placed


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cases in {argv[1]}")
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
