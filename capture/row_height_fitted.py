"""Capture the rows Excel fits to turned text and to styled text, read from Excel.

    uv run --group capture python -m capture.row_height_fitted build PATH  # the workbooks only
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.row_height_fitted run
    uv run --group capture python -m capture.row_height_fitted read  # the saved answer again

Like ``row_height_more`` it reads what Excel reports, not the screen: the same
``row_height_more.applescript`` opens each workbook, reads ``row height`` and the font Excel
applied in column A of each row named, and closes it without saving. Every workbook has
openpyxl's default styles (Normal Calibri 11, ``<sheetFormatPr baseColWidth="8"
defaultRowHeight="15">``, no ``customHeight``), so Excel fits each listed row to its cells. One
case a row, from row 1, in column A; no row has a stored height. Groups, a workbook each:

- ``turned`` (``ARCHITECTURE.md`` §4.17, fitted rows): text at an angle in Calibri regular.
  Three texts ("Rot", "Rotated text here", "Rotated text here, and more of it") at 90 degrees
  at every size from 9 to 36 points; runs of one glyph (``o``, ``m``, ``i``, 1 to 40 of them) at
  90 degrees at 11, 14 and 20; the three texts at 88, 89 and -90 degrees at 20; and short runs
  (``o`` 1, 2, 4 and ``i`` 3) at 16 to 36. They settle the 90-degree rule's pad at every size,
  -90 degrees and the edge at 88 and 89.
- ``styled`` (§4.7, §4.16): "HHH" in Calibri at every half size from 6 to 40 and every size from
  41 to 72; superscript, subscript, a single accounting underline, and a superscript cell beside
  a subscript one (column B), at 6 to 40; and Arial, Times New Roman, Courier New and Aptos Narrow
  at 6 to 40. They settle the row a font fits at any size and the row superscript and subscript
  text fits.

``run`` saves Excel's answer, with the Excel it came from, to ``capture/build/row_height_fitted/
answer.tsv``; ``read`` writes ``tests/captures/row_height_fitted.tsv`` from it again, one record
per row. No screenshot, no clipboard, no marker colors.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "row_height_fitted.tsv"
BUILD_DIR = ROOT / "capture" / "build" / "row_height_fitted"
ANSWER = BUILD_DIR / "answer.tsv"
CONTAINER_DIR = (
    Path.home()
    / "Library"
    / "Containers"
    / "com.microsoft.Excel"
    / "Data"
    / "Documents"
    / "sheetpeek"
    / "row_height_fitted"
)
SCRIPT = Path(__file__).with_name("row_height_more.applescript")
GROUPS = ("turned", "styled")
TEXTS = {
    "short": "Rot",
    "medium": "Rotated text here",
    "long": "Rotated text here, and more of it",
}
STYLED = "HHH"
FACES = ("Arial", "Times New Roman", "Courier New", "Aptos Narrow")
COLUMNS = ("group", "name", "row", "rotation", "size", "applied", "height", "text")


@dataclass(frozen=True)
class Case:
    """One row read: its group (workbook), name, text and the cell's font and angle."""

    group: str
    name: str
    text: str
    size: float
    face: str = "Calibri"
    rotation: int = 0  # degrees, -90 to 90
    vert_align: str | None = None
    underline: str | None = None
    beside: str | None = None  # column B's vertical alignment, for a superscript beside a subscript


def _turned() -> list[Case]:
    found = []
    for size in range(9, 37):
        for kind, text in TEXTS.items():
            found.append(Case("turned", f"90 {kind} {size}", text, size, rotation=90))
    for size in (11, 14, 20):
        found.extend(
            Case("turned", f"90 {glyph}x{n} {size}", glyph * n, size, rotation=90)
            for glyph in "omi"
            for n in range(1, 41)
        )
    for size in range(16, 37):
        if size == 20:
            continue  # among the runs above
        for glyph, n in (("o", 1), ("o", 2), ("o", 4), ("i", 3)):
            found.append(Case("turned", f"90 {glyph}x{n} {size}", glyph * n, size, rotation=90))
    for angle in (88, 89, -90):
        for kind, text in TEXTS.items():
            found.append(Case("turned", f"{angle} {kind} 20", text, 20, rotation=angle))
    return found


def _styled() -> list[Case]:
    sizes = [s / 2 for s in range(12, 81)] + list(range(41, 73))
    found = [Case("styled", f"plain {size:g}", STYLED, size) for size in sizes]
    for size in range(6, 41):
        found += [
            Case("styled", f"superscript {size}", STYLED, size, vert_align="superscript"),
            Case("styled", f"subscript {size}", STYLED, size, vert_align="subscript"),
            Case(
                "styled",
                f"both {size}",
                STYLED,
                size,
                vert_align="superscript",
                beside="subscript",
            ),
            Case(
                "styled",
                f"singleAccounting {size}",
                STYLED,
                size,
                underline="singleAccounting",
            ),
        ]
        for face in FACES:
            found.append(Case("styled", f"{face} {size}", STYLED, size, face=face))
    return found


CASES = _turned() + _styled()


def stored_rotation(angle: int) -> int:
    """``textRotation`` as stored: 0 to 90 as is, -1 to -90 as 90 minus the angle."""
    return angle if angle >= 0 else 90 - angle


def workbook_name(group: str) -> str:
    return f"row_height_fitted_{group}.xlsx"


def rows(group: str) -> list[tuple[int, Case]]:
    """Each case of ``group`` with its row, from 1."""
    return list(enumerate((c for c in CASES if c.group == group), 1))


def build_one(group: str, folder: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = group
    for row, case in rows(group):
        cell = ws.cell(row, 1, case.text)
        if (case.face, case.size, case.vert_align, case.underline) != ("Calibri", 11, None, None):
            cell.font = Font(
                name=case.face, size=case.size, vertAlign=case.vert_align, underline=case.underline
            )
        if case.rotation:
            cell.alignment = Alignment(textRotation=stored_rotation(case.rotation))
        if case.beside:
            other = ws.cell(row, 2, case.text)
            other.font = Font(name=case.face, size=case.size, vertAlign=case.beside)
    path = folder / workbook_name(group)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def build(folder: Path = CONTAINER_DIR) -> list[Path]:
    """Write each group's workbook into ``folder``; return their paths in group order."""
    return [build_one(group, folder) for group in GROUPS]


def run() -> None:
    """Build, have Excel read the rows (needs Excel), save its answer, write the TSV."""
    from capture.edition import excel_description

    arguments = []
    for group, path in zip(GROUPS, build(), strict=True):
        arguments += [str(path), ",".join(str(row) for row, _ in rows(group))]
    try:
        result = subprocess.run(
            ["osascript", str(SCRIPT), *arguments],
            capture_output=True,
            text=True,
            check=False,
            timeout=1200,  # kills osascript, which ignores SIGTERM while it waits on Excel
        )
    except subprocess.TimeoutExpired:
        sys.exit("osascript did not answer in 20 minutes; check Excel before another run")
    if result.returncode != 0:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    with ANSWER.open("w", encoding="utf-8", newline="\n") as out:
        out.write(f"# excel: {excel_description()}, en_US\n")
        out.write(f"# captured: {dt.date.today().isoformat()}\n")
        out.write(result.stdout.strip() + "\n")
    write()


def records(answer: str) -> tuple[list[str], list[dict[str, str]]]:
    """The header lines and the records of a saved answer, checked against ``CASES``.

    Each answer line is ``file name, row, applied font name, applied size, row height``, tab
    separated, in case order.
    """
    lines = answer.strip().split("\n")
    head = [line for line in lines if line.startswith("# ")]
    body = [line for line in lines if line and not line.startswith("# ")]
    expected = [(group, row, case) for group in GROUPS for row, case in rows(group)]
    if len(body) != len(expected):
        raise ValueError(f"expected {len(expected)} lines from Excel, got {len(body)}")
    found = []
    for (group, row, case), line in zip(expected, body, strict=True):
        name, answered, applied_name, applied_size, height = line.split("\t")
        if (name, int(answered)) != (workbook_name(group), row):
            raise ValueError(f"Excel answered for {name} row {answered}, expected {case}")
        found.append(
            {
                "group": group,
                "name": case.name,
                "row": str(row),
                "rotation": str(case.rotation),
                "size": f"{case.size:g}",
                "applied": f"{applied_name} {_number(applied_size)}",
                "height": _number(height),
                "text": case.text,
            }
        )
    return head, found


def write(answer: Path = ANSWER, output: Path = OUTPUT) -> None:
    """Write ``row_height_fitted.tsv`` from Excel's saved answer (no Excel)."""
    if not answer.exists():
        sys.exit(f"not captured yet: no {answer}")
    head, found = records(answer.read_text(encoding="utf-8"))
    excel = next((h[len("# excel: ") :] for h in head if h.startswith("# excel: ")), "unknown")
    date = next((h[len("# captured: ") :] for h in head if h.startswith("# captured: ")), "")
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8", newline="\n") as out:
        out.write(f"# suite: row_height_fitted (capture.row_height_fitted); {excel}\n")
        out.write(
            "# a workbook per group, openpyxl's default styles (Normal Calibri 11, no "
            "customHeight); one case a row in column A, fitted by Excel on opening\n"
        )
        out.write(f"# captured: {date}; method: row height (AppleScript)\n")
        out.write(
            "# name: the case (CASES in capture/row_height_fitted.py); rotation: degrees; size: "
            "the cell's font size; applied: the font Excel applied to column A; height: the "
            "row's height in points\n"
        )
        out.write("\t".join(COLUMNS) + "\n")
        for record in found:
            out.write("\t".join(record[c] for c in COLUMNS) + "\n")
    print(f"{len(found)} rows written to {output}")


def _number(text: str) -> str:
    """``15.0`` as ``15``; other text unchanged."""
    try:
        value = float(text)
    except ValueError:
        return text
    return repr(int(value)) if value.is_integer() else repr(value)


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        paths = build(Path(argv[1]))
        print(f"{len(paths)} workbooks, {len(CASES)} rows in {argv[1]}")
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
