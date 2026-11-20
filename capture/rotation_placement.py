"""Capture where Excel places angled and stacked text at more angles, rows and sizes (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.rotation_placement run
    uv run --group capture python -m capture.rotation_placement read        # the saved captures
    uv run --group capture python -m capture.rotation_placement build PATH  # the workbook only

It extends ``rotation_more`` with its windows, filling and measuring (``capture/rotation_more.py``,
whose ``run`` and ``write`` take this suite's cases): every case "Rotated text here" in Calibri 11
in column E unless named, a workbook of its own per window, at zoom 100 in the dark appearance.
Cases (``ARCHITECTURE.md`` §4.17 and the rule each settles):

- ``spread``: fill, justify, distributed and left (the control) at 30, 60, 85, -15, -75 and -90
  degrees ("fill, justify and distributed are placed as left at positive angles and as right at
  negative ones; distributed at 90 is centered"; "fill does not spill"); justified and
  distributed rows 5 pt over their fitted height, as ``rotation_more``'s are.
- ``top``: top alignment without borders at 15, 30, 60 and 75 degrees in 90 pt rows, and at 10,
  15 and 20 in 30 pt rows; between four sides at 20, 40, 50, 70, -30 and -60 (90 pt rows), and at
  10, 15 and 20 in 30 pt rows, 15 in a 45 pt row and 20 in 50 and 70 pt rows ("the highest corner
  3 pt below the top", with and without sides, and in rows with little room).
- ``stacked``: "Abc" stacked, left, right and center, at 9, 12, 14, 16, 20, 24, 28 and 36 points,
  and "WiW" at 11 and 20, in ``rotation_more``'s 20-character column ("each character centered on
  one axis", and where the axis lies).

Measured as ``rotation_more`` measures: the ink box, the sides of bordered cases and each stacked
line's ink (``lines``).
"""

from __future__ import annotations

import math
import sys
from pathlib import Path

from capture import rotation
from capture import rotation_more as rm

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "rotation_placement"
OUTPUT = ROOT / "tests" / "captures" / "rotation_placement.tsv"
SUITE = "rotation_placement"
WIDTH = 78.0
"""The width of "Rotated text here" in Calibri 11 (xlformat), for the spread rows' heights."""


def _fitted(angle: int) -> int:
    """The row ``rotation`` fits to "Rotated text here" at ``angle``: ``W + 7`` at 90 degrees."""
    if abs(angle) == 90:
        return round(WIDTH + 7)
    a = math.radians(abs(angle))
    return round(WIDTH * math.sin(a) + 11 * math.cos(a) + 7)


def _stacked(name: str, text: str, horizontal: str, size: int) -> rm.Case:
    cell = rotation.Case(
        "stacked", name, text, 255, horizontal=horizontal, size=size, width=rm.STACKED_WIDTH
    )
    return rm.Case("stacked", name, cell, measure=("ink", "lines"))


def cases() -> list[rm.Case]:
    found = []
    for angle in (30, 60, 85, -15, -75, -90):
        for horizontal in ("fill", "justify", "distributed", "left"):
            height = None if horizontal in ("fill", "left") else _fitted(angle) + 5
            found.append(
                rm._plain("spread", f"{angle} {horizontal}", angle, horizontal=horizontal,
                          height=height)
            )  # fmt: skip
    found += [
        rm._plain("top", f"{a} top 90", a, vertical="top", height=90) for a in (15, 30, 60, 75)
    ]
    found += [rm._plain("top", f"{a} top 30", a, vertical="top", height=30) for a in (10, 15, 20)]
    found += [
        rm._slant("top", f"{a} lrtb top 90", a, vertical="top", height=90)
        for a in (20, 40, 50, 70, -30, -60)
    ]
    found += [
        rm._slant("top", f"{a} lrtb top {h}", a, vertical="top", height=h)
        for a, h in ((10, 30), (15, 30), (15, 45), (20, 30), (20, 50), (20, 70))
    ]
    found += [
        _stacked(f"stacked {text} {horizontal} {size}", text, horizontal, size)
        for text, sizes in (("Abc", (9, 12, 14, 16, 20, 24, 28, 36)), ("WiW", (11, 20)))
        for size in sizes
        for horizontal in ("left", "right", "center")
    ]
    return found


CASES = cases()


def run() -> None:
    """Capture every window, then write the TSV (needs Excel)."""
    rm.run(CASES, BUILD_DIR, OUTPUT, SUITE)


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``rotation_placement.tsv``."""
    rm.write(CASES, BUILD_DIR, OUTPUT, SUITE)


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = rm.build(Path(argv[1]), CASES)
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
