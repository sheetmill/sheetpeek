"""Capture the colors Excel for Mac resolves for stored 16-bit tints and tints of black (§3.3).

    export EXCEL_EDITION="Microsoft 365"
    uv run --group capture python -m capture.theme_colors_more run
    uv run --group capture python -m capture.theme_colors_more read         # write the TSV again
    uv run --group capture python -m capture.theme_colors_more build PATH   # the workbook

One workbook, openpyxl's default theme, no window and no screenshot: like the ``theme_colors``
suite, each cell's solid fill and font color are a theme color with a tint, and
``theme_colors.applescript`` (shared with that suite, unchanged) reads the fill and font color
Excel resolves for each cell (``color of interior object``, ``color of font object``). Every
cell holds an ``x``; no case draws magenta.

- ``stored``: rows 1 to 10 hold theme indices 0 to 9, columns A to G the tints as files store
  them, 16-bit fractions just short of a round tint (0.6 stored as 0.59999389629810485, and
  0.4, 0.8, 0.25, 0.5, 0.15 and -0.25 likewise). ARCHITECTURE §3.3, ``_colors.apply_tint``:
  "Stored 16-bit tint fractions draw as the round tint" (the tolerance added before truncating
  the lightness), unverified until this capture.
- ``third``: rows 12 and 13 hold theme 1 (dark 1, black) and theme 3 (dark 2, ``1F497D``),
  columns A to C the tints 0.333333, 0.33333 and 0.3333333333. §3.3: "Tint 0.333333 on black
  is 555555" (lightness 80 of 240, not 79), unverified until this capture.

Tints are written into ``styles.xml`` as the ``repr`` of the case's float, exactly: openpyxl's
``%.16g`` is rewritten (:func:`build`). ``run`` saves Excel's answers with the environment
(``capture/build/theme_colors_more/excel.tsv``); ``read`` (:func:`write`) turns them into
``tests/captures/theme_colors_more.tsv`` without Excel, one record per cell.
"""

from __future__ import annotations

import datetime as dt
import os
import re
import subprocess
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import Color

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "theme_colors_more.tsv"
BUILD_DIR = ROOT / "capture" / "build" / "theme_colors_more"
SAVED = BUILD_DIR / "excel.tsv"
CONTAINER_DIR = (
    Path.home() / "Library" / "Containers" / "com.microsoft.Excel" / "Data" / "Documents"
    / "sheetpeek"
)  # fmt: skip
WORKBOOK = CONTAINER_DIR / "theme_colors_more.xlsx"
SCRIPT = Path(__file__).with_name("theme_colors.applescript")

STORED_TINTS: tuple[tuple[float, float], ...] = (
    (0.59999389629810485, 0.6),
    (0.3999755851924192, 0.4),
    (0.79998168889431442, 0.8),
    (0.249977111117893, 0.25),
    (0.499984740745262, 0.5),
    (0.14999847407452621, 0.15),
    (-0.249977111117893, -0.25),
)
"""(the tint as files store it, the round tint it stands for)."""
STORED_THEMES = range(10)
THIRD_TINTS = (0.333333, 0.33333, 0.3333333333)
THIRD_THEMES = (1, 3)
THIRD_ROW = 12
COLUMNS = ("group", "cell", "theme", "tint", "round", "fill", "font")


def _letter(k: int) -> str:
    return "ABCDEFGHIJ"[k]


def cases() -> list[tuple[str, str, int, float, float]]:
    """(group, address, theme index, tint, round tint) in the order Excel reads them."""
    found = [
        ("stored", f"{_letter(k)}{theme + 1}", theme, tint, round_)
        for theme in STORED_THEMES
        for k, (tint, round_) in enumerate(STORED_TINTS)
    ]
    found += [
        ("third", f"{_letter(k)}{THIRD_ROW + j}", theme, tint, tint)
        for j, theme in enumerate(THIRD_THEMES)
        for k, tint in enumerate(THIRD_TINTS)
    ]
    return found


CASES = cases()

EXPECTED: dict[str, tuple[str, str]] = {
    # what sheetpeek draws today for each cell, fill and font color (taken by running it on the
    # built workbook); the capture shows where Excel differs
    "A1": ("FFFFFF", "FFFFFF"), "B1": ("FFFFFF", "FFFFFF"), "C1": ("FFFFFF", "FFFFFF"),
    "D1": ("FFFFFF", "FFFFFF"), "E1": ("FFFFFF", "FFFFFF"), "F1": ("FFFFFF", "FFFFFF"),
    "G1": ("BFBFBF", "BFBFBF"),
    "A2": ("999999", "999999"), "B2": ("666666", "666666"), "C2": ("CCCCCC", "CCCCCC"),
    "D2": ("404040", "404040"), "E2": ("808080", "808080"), "F2": ("262626", "262626"),
    "G2": ("000000", "000000"),
    "A3": ("F8F7F3", "F8F7F3"), "B3": ("F4F3EC", "F4F3EC"), "C3": ("FBFBF9", "FBFBF9"),
    "D3": ("F2F1E8", "F2F1E8"), "E3": ("F7F5F0", "F7F5F0"), "F3": ("F0EFE6", "F0EFE6"),
    "G3": ("C4BD97", "C4BD97"),
    "A4": ("8DB4E2", "8DB4E2"), "B4": ("538DD5", "538DD5"), "C4": ("C5D9F1", "C5D9F1"),
    "D4": ("3072C2", "3072C2"), "E4": ("70A0DC", "70A0DC"), "F4": ("2962A7", "2962A7"),
    "G4": ("16365C", "16365C"),
    "A5": ("B8CCE4", "B8CCE4"), "B5": ("95B3D7", "95B3D7"), "C5": ("DCE6F1", "DCE6F1"),
    "D5": ("7A9FCD", "7A9FCD"), "E5": ("A7BFDE", "A7BFDE"), "F5": ("6994C7", "6994C7"),
    "G5": ("366092", "366092"),
    "A6": ("E6B8B7", "E6B8B7"), "B6": ("DA9694", "DA9694"), "C6": ("F2DCDB", "F2DCDB"),
    "D6": ("D07B79", "D07B79"), "E6": ("DFA7A6", "DFA7A6"), "F6": ("C96967", "C96967"),
    "G6": ("963634", "963634"),
    "A7": ("D8E4BC", "D8E4BC"), "B7": ("C4D79B", "C4D79B"), "C7": ("EBF1DE", "EBF1DE"),
    "D7": ("B4CC82", "B4CC82"), "E7": ("CDDDAC", "CDDDAC"), "F7": ("AAC671", "AAC671"),
    "G7": ("76933C", "76933C"),
    "A8": ("CCC0DA", "CCC0DA"), "B8": ("B1A0C7", "B1A0C7"), "C8": ("E4DFEC", "E4DFEC"),
    "D8": ("9F8AB9", "9F8AB9"), "E8": ("BFB0D0", "BFB0D0"), "F8": ("937AAF", "937AAF"),
    "G8": ("60497A", "60497A"),
    "A9": ("B7DEE8", "B7DEE8"), "B9": ("92CDDC", "92CDDC"), "C9": ("DAEEF3", "DAEEF3"),
    "D9": ("77C1D5", "77C1D5"), "E9": ("A5D6E2", "A5D6E2"), "F9": ("64B9CE", "64B9CE"),
    "G9": ("31869B", "31869B"),
    "A10": ("FCD5B4", "FCD5B4"), "B10": ("FABF8F", "FABF8F"), "C10": ("FDE9D9", "FDE9D9"),
    "D10": ("F9AF73", "F9AF73"), "E10": ("FBCAA2", "FBCAA2"), "F10": ("F8A561", "F8A561"),
    "G10": ("E26B0A", "E26B0A"),
    "A12": ("555555", "555555"), "B12": ("555555", "555555"), "C12": ("555555", "555555"),
    "A13": ("4081D0", "4081D0"), "B13": ("4081D0", "4081D0"), "C13": ("4081D0", "4081D0"),
}  # fmt: skip


def build(path: Path = WORKBOOK) -> Path:
    """Write the workbook, each tint in ``styles.xml`` as the ``repr`` of its float."""
    from capture.screen.run import patch_part

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "theme"
    for _group, address, theme, tint, _round in CASES:
        cell = ws[address]
        cell.value = "x"
        cell.fill = PatternFill("solid", fgColor=Color(theme=theme, tint=tint))
        cell.font = Font(color=Color(theme=theme, tint=tint))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    patch_part(path, {"xl/styles.xml": exact_tints})
    return path


def exact_tints(styles: str) -> str:
    """``styles.xml`` with every ``tint`` attribute rewritten as the ``repr`` of its value."""
    return re.sub(
        r'tint="([^"]*)"', lambda m: f'tint="{float(m.group(1))!r}"', styles
    )  # fmt: skip


def run() -> None:
    """Build, have Excel read the colors, save its answers, write the TSV."""
    from capture.screen.session import environment

    edition = os.environ.get("EXCEL_EDITION")
    if not edition:
        raise SystemExit("state the Excel edition (Excel > About Excel) in EXCEL_EDITION")
    found = environment(edition)
    build()
    try:
        result = subprocess.run(
            ["osascript", str(SCRIPT), str(WORKBOOK), *(c[1] for c in CASES)],
            capture_output=True, text=True, check=False, timeout=600,
        )  # fmt: skip
    finally:
        WORKBOOK.unlink(missing_ok=True)
    if result.returncode != 0:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    measure(result.stdout)  # stops on an answer that does not fit before anything is saved
    found["date"] = dt.date.today().isoformat()
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    with SAVED.open("w", encoding="utf-8", newline="\n") as out:
        for key, value in found.items():
            out.write(f"# {key}: {value}\n")
        out.write(result.stdout.strip() + "\n")
    write()


def measure(answer: str) -> list[dict[str, str]]:
    """The records for Excel's answer: one line per case, ``cell<TAB>r,g,b<TAB>r,g,b``."""
    lines = [line for line in answer.strip().split("\n") if line and not line.startswith("# ")]
    if len(lines) != len(CASES):
        raise ValueError(f"expected {len(CASES)} cells, got {len(lines)}")
    records = []
    for (group, address, theme, tint, round_), line in zip(CASES, lines, strict=True):
        got, fill, font = line.split("\t")
        if got != address:
            raise ValueError(f"Excel answered for {got}, expected {address}")
        records.append(
            {
                "group": group,
                "cell": address,
                "theme": str(theme),
                "tint": repr(tint),
                "round": repr(round_),
                "fill": _hex(fill),
                "font": _hex(font),
            }
        )
    return records


def _hex(rgb: str) -> str:
    parts = rgb.split(",")
    if len(parts) != 3:
        raise ValueError(f"not an r,g,b color: {rgb!r}")
    return "".join(f"{int(v):02X}" for v in parts)


def write(saved: Path = SAVED, output: Path = OUTPUT) -> list[dict[str, str]]:
    """Measure the saved answers again (no Excel) and write ``theme_colors_more.tsv``."""
    text = saved.read_text(encoding="utf-8")
    first = [line for line in text.splitlines() if line.startswith("# ")]
    records = measure(text)
    with output.open("w", encoding="utf-8", newline="\n") as out:
        for line in first:
            out.write(line + "\n")
        out.write("# suite: theme_colors_more\n")
        out.write("# openpyxl's default theme; each cell's solid fill and font color are the\n")
        out.write("# theme index and tint given (tint as stored, round: the tint it stands for);\n")
        out.write("# fill and font: the RRGGBB Excel resolves (color of interior object, color\n")
        out.write("# of font object); method: AppleScript, per cell, no window\n")
        out.write("\t".join(COLUMNS) + "\n")
        for record in records:
            out.write("\t".join(record[c] for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {output}")
    return records


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        build(Path(argv[1]))
        print(f"{len(CASES)} cells in {argv[1]}")
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
