"""Capture the rows Excel fits around vertical merges: wrapped text and larger fonts in a merge.

    uv run --group capture python -m capture.merge_rows build PATH  # the workbook only
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.merge_rows run
    uv run --group capture python -m capture.merge_rows read  # the saved answer again

The suite ``row_height`` captured a horizontal merge (A1:C1): its font counts in the row Excel
fits and its wrapped lines do not. This suite asks the same of merges across rows, which no
capture covered (``ARCHITECTURE.md`` §4.4, §4.7). Like ``row_height_more``, whose AppleScript
it uses, it reads what Excel itself reports, not the screen: ``row height`` of every row of
each case, and the font Excel applied in column A.

One workbook under openpyxl's Calibri 11 Normal font and ``<sheetFormatPr baseColWidth="8"
defaultRowHeight="15">`` (no ``customHeight``, so Excel fits every row with a ``<row>``
element). Each case (``BLOCKS``) holds its text in column A, merged over its rows (and over
column B for the one merge two columns wide); column C of each of its rows holds an "x" in the
Normal font, so that every row has an element and is fitted; one empty row separates the
cases. Two controls repeat what is captured: three wrapped lines in one cell (48 pt) and in a
merge over A:C (15 pt).

``PREDICTIONS`` gives each row's height under each candidate rule, written before the run:
``sheetpeek``, what sheetpeek lays out (a merge's font counts in its first row, its wrapped
lines in none); ``last_row`` and ``first_row``, the wrapped lines' height beyond the merged rows'
added to its last or first row; ``font_every_row``, a merge's font counting in each of its rows.

``run`` saves Excel's answer to ``capture/build/merge_rows/answer.tsv``; ``read`` writes
``tests/captures/merge_rows.tsv`` from it again, one record per row read. No screenshot, no
clipboard, no marker colors; one workbook, under a minute of Excel.
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
OUTPUT = ROOT / "tests" / "captures" / "merge_rows.tsv"
BUILD_DIR = ROOT / "capture" / "build" / "merge_rows"
ANSWER = BUILD_DIR / "answer.tsv"
CONTAINER_DIR = (
    Path.home()
    / "Library"
    / "Containers"
    / "com.microsoft.Excel"
    / "Data"
    / "Documents"
    / "sheetpeek"
    / "merge_rows"
)
SCRIPT = Path(__file__).with_name("row_height_more.applescript")
WORKBOOK = "merge_rows.xlsx"
COLUMNS = ("label", "offset", "rows", "font", "wrapped_lines", "applied", "height")


@dataclass(frozen=True)
class Block:
    """One case: column A's text over ``rows`` rows (and ``wide`` columns), in ``size`` points."""

    label: str
    rows: int
    lines: int  # wrapped lines; 0 for a plain "H"
    size: int = 11
    wide: int = 1  # columns merged
    neighbor: int = 0  # a cell font in column B of the case's last row, 0 for none


BLOCKS: tuple[Block, ...] = (
    Block("control: 3 wrapped lines, one cell", 1, 3),
    Block("control: 3 wrapped lines, merge A:C", 1, 3, wide=3),
    Block("2 wrapped lines over 2 rows", 2, 2),
    Block("3 wrapped lines over 2 rows", 2, 3),
    Block("6 wrapped lines over 3 rows", 3, 6),
    Block("3 wrapped lines in 14 pt over 2 rows", 2, 3, size=14),
    Block("H in 14 pt over 2 rows", 2, 0, size=14),
    Block("H in 20 pt over 2 rows", 2, 0, size=20),
    Block("H in 20 pt over 3 rows", 3, 0, size=20),
    Block("3 wrapped lines over 2 rows, 14 pt beside in the second", 2, 3, neighbor=14),
    Block("3 wrapped lines over 2 rows and 2 columns", 2, 3, wide=2),
)


@dataclass(frozen=True)
class Case:
    """One row read: its block, its offset in the block (1-based) and its sheet row."""

    block: Block
    offset: int
    row: int


def _cases() -> list[Case]:
    found, row = [], 1
    for block in BLOCKS:
        found += [Case(block, k, row + k - 1) for k in range(1, block.rows + 1)]
        row += block.rows + 1
    return found


CASES = _cases()


def build(folder: Path = CONTAINER_DIR) -> Path:
    """Write the workbook into ``folder``; return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "rows"
    for block in BLOCKS:
        first = next(c.row for c in CASES if c.block is block)
        last = first + block.rows - 1
        text = "\n".join("abcdef"[: block.lines]) if block.lines else "H"
        cell = ws.cell(first, 1, text)
        if block.size != 11:
            cell.font = Font(name="Calibri", sz=block.size, family=2)
        if block.lines:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        for row in range(first, last + 1):
            ws.cell(row, 3, "x")
        if block.neighbor:
            ws.cell(last, 2, "x").font = Font(name="Calibri", sz=block.neighbor, family=2)
        if block.rows > 1 or block.wide > 1:
            end = openpyxl.utils.get_column_letter(block.wide)
            ws.merge_cells(f"A{first}:{end}{last}")
    path = folder / WORKBOOK
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


# Each row's height under each candidate, by (label, offset); written before the run.
PREDICTIONS: dict[str, dict[tuple[str, int], float]] = {
    "sheetpeek": {
        ("control: 3 wrapped lines, one cell", 1): 48.0,
        ("control: 3 wrapped lines, merge A:C", 1): 15.0,
        ("2 wrapped lines over 2 rows", 1): 15.0,
        ("2 wrapped lines over 2 rows", 2): 15.0,
        ("3 wrapped lines over 2 rows", 1): 15.0,
        ("3 wrapped lines over 2 rows", 2): 15.0,
        ("6 wrapped lines over 3 rows", 1): 15.0,
        ("6 wrapped lines over 3 rows", 2): 15.0,
        ("6 wrapped lines over 3 rows", 3): 15.0,
        ("3 wrapped lines in 14 pt over 2 rows", 1): 19.0,
        ("3 wrapped lines in 14 pt over 2 rows", 2): 15.0,
        ("H in 14 pt over 2 rows", 1): 19.0,
        ("H in 14 pt over 2 rows", 2): 15.0,
        ("H in 20 pt over 2 rows", 1): 26.0,
        ("H in 20 pt over 2 rows", 2): 15.0,
        ("H in 20 pt over 3 rows", 1): 26.0,
        ("H in 20 pt over 3 rows", 2): 15.0,
        ("H in 20 pt over 3 rows", 3): 15.0,
        ("3 wrapped lines over 2 rows, 14 pt beside in the second", 1): 15.0,
        ("3 wrapped lines over 2 rows, 14 pt beside in the second", 2): 19.0,
        ("3 wrapped lines over 2 rows and 2 columns", 1): 15.0,
        ("3 wrapped lines over 2 rows and 2 columns", 2): 15.0,
    },
    "last_row": {
        ("control: 3 wrapped lines, one cell", 1): 48.0,
        ("control: 3 wrapped lines, merge A:C", 1): 48.0,
        ("2 wrapped lines over 2 rows", 1): 15.0,
        ("2 wrapped lines over 2 rows", 2): 17.0,
        ("3 wrapped lines over 2 rows", 1): 15.0,
        ("3 wrapped lines over 2 rows", 2): 33.0,
        ("6 wrapped lines over 3 rows", 1): 15.0,
        ("6 wrapped lines over 3 rows", 2): 15.0,
        ("6 wrapped lines over 3 rows", 3): 66.0,
        ("3 wrapped lines in 14 pt over 2 rows", 1): 19.0,
        ("3 wrapped lines in 14 pt over 2 rows", 2): 41.0,
        ("H in 14 pt over 2 rows", 1): 19.0,
        ("H in 14 pt over 2 rows", 2): 15.0,
        ("H in 20 pt over 2 rows", 1): 26.0,
        ("H in 20 pt over 2 rows", 2): 15.0,
        ("H in 20 pt over 3 rows", 1): 26.0,
        ("H in 20 pt over 3 rows", 2): 15.0,
        ("H in 20 pt over 3 rows", 3): 15.0,
        ("3 wrapped lines over 2 rows, 14 pt beside in the second", 1): 15.0,
        ("3 wrapped lines over 2 rows, 14 pt beside in the second", 2): 33.0,
        ("3 wrapped lines over 2 rows and 2 columns", 1): 15.0,
        ("3 wrapped lines over 2 rows and 2 columns", 2): 33.0,
    },
    "first_row": {
        ("control: 3 wrapped lines, one cell", 1): 48.0,
        ("control: 3 wrapped lines, merge A:C", 1): 48.0,
        ("2 wrapped lines over 2 rows", 1): 17.0,
        ("2 wrapped lines over 2 rows", 2): 15.0,
        ("3 wrapped lines over 2 rows", 1): 33.0,
        ("3 wrapped lines over 2 rows", 2): 15.0,
        ("6 wrapped lines over 3 rows", 1): 66.0,
        ("6 wrapped lines over 3 rows", 2): 15.0,
        ("6 wrapped lines over 3 rows", 3): 15.0,
        ("3 wrapped lines in 14 pt over 2 rows", 1): 45.0,
        ("3 wrapped lines in 14 pt over 2 rows", 2): 15.0,
        ("H in 14 pt over 2 rows", 1): 19.0,
        ("H in 14 pt over 2 rows", 2): 15.0,
        ("H in 20 pt over 2 rows", 1): 26.0,
        ("H in 20 pt over 2 rows", 2): 15.0,
        ("H in 20 pt over 3 rows", 1): 26.0,
        ("H in 20 pt over 3 rows", 2): 15.0,
        ("H in 20 pt over 3 rows", 3): 15.0,
        ("3 wrapped lines over 2 rows, 14 pt beside in the second", 1): 29.0,
        ("3 wrapped lines over 2 rows, 14 pt beside in the second", 2): 19.0,
        ("3 wrapped lines over 2 rows and 2 columns", 1): 33.0,
        ("3 wrapped lines over 2 rows and 2 columns", 2): 15.0,
    },
    "font_every_row": {
        ("control: 3 wrapped lines, one cell", 1): 48.0,
        ("control: 3 wrapped lines, merge A:C", 1): 15.0,
        ("2 wrapped lines over 2 rows", 1): 15.0,
        ("2 wrapped lines over 2 rows", 2): 15.0,
        ("3 wrapped lines over 2 rows", 1): 15.0,
        ("3 wrapped lines over 2 rows", 2): 15.0,
        ("6 wrapped lines over 3 rows", 1): 15.0,
        ("6 wrapped lines over 3 rows", 2): 15.0,
        ("6 wrapped lines over 3 rows", 3): 15.0,
        ("3 wrapped lines in 14 pt over 2 rows", 1): 19.0,
        ("3 wrapped lines in 14 pt over 2 rows", 2): 19.0,
        ("H in 14 pt over 2 rows", 1): 19.0,
        ("H in 14 pt over 2 rows", 2): 19.0,
        ("H in 20 pt over 2 rows", 1): 26.0,
        ("H in 20 pt over 2 rows", 2): 26.0,
        ("H in 20 pt over 3 rows", 1): 26.0,
        ("H in 20 pt over 3 rows", 2): 26.0,
        ("H in 20 pt over 3 rows", 3): 26.0,
        ("3 wrapped lines over 2 rows, 14 pt beside in the second", 1): 15.0,
        ("3 wrapped lines over 2 rows, 14 pt beside in the second", 2): 19.0,
        ("3 wrapped lines over 2 rows and 2 columns", 1): 15.0,
        ("3 wrapped lines over 2 rows and 2 columns", 2): 15.0,
    },
}


def run() -> None:
    """Build, have Excel read the rows (needs Excel), save its answer, write the TSV."""
    from capture.edition import excel_description

    path = build()
    rows = ",".join(str(c.row) for c in CASES)
    try:
        result = subprocess.run(
            ["osascript", str(SCRIPT), str(path), rows],
            capture_output=True,
            text=True,
            check=False,
            timeout=300,  # kills osascript, which ignores SIGTERM while it waits on Excel
        )
    except subprocess.TimeoutExpired:
        sys.exit("osascript did not answer in 5 minutes; check Excel before another run")
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
    if len(body) != len(CASES):
        raise ValueError(f"expected {len(CASES)} lines from Excel, got {len(body)}")
    found = []
    for case, line in zip(CASES, body, strict=True):
        name, row, applied_name, applied_size, height = line.split("\t")
        if (name, int(row)) != (WORKBOOK, case.row):
            raise ValueError(f"Excel answered for {name} row {row}, expected {case}")
        found.append(
            {
                "label": case.block.label,
                "offset": str(case.offset),
                "rows": str(case.block.rows),
                "font": f"Calibri {case.block.size}",
                "wrapped_lines": str(case.block.lines),
                "applied": f"{applied_name} {_number(applied_size)}",
                "height": _number(height),
            }
        )
    return head, found


def write(answer: Path = ANSWER, output: Path = OUTPUT) -> None:
    """Write ``merge_rows.tsv`` from Excel's saved answer (no Excel)."""
    if not answer.exists():
        sys.exit(f"not captured yet: no {answer}")
    head, found = records(answer.read_text(encoding="utf-8"))
    excel = next((h[len("# excel: ") :] for h in head if h.startswith("# excel: ")), "unknown")
    date = next((h[len("# captured: ") :] for h in head if h.startswith("# captured: ")), "")
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8", newline="\n") as out:
        out.write(f"# suite: merge_rows (capture.merge_rows); {excel}\n")
        out.write(
            '# one workbook, Calibri 11 Normal, openpyxl\'s <sheetFormatPr baseColWidth="8" '
            'defaultRowHeight="15"> (no customHeight)\n'
        )
        out.write(f"# captured: {date}; method: row height (AppleScript)\n")
        out.write(
            "# label: the case (BLOCKS in capture/merge_rows.py); offset: the row within it; "
            "rows: the rows merged; font: column A's font; wrapped_lines: its wrapped lines (0: "
            'one "H"); applied: the font Excel applied; height: the row\'s height in points\n'
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
        path = build(Path(argv[1]))
        print(f"{len(CASES)} rows in {path}")
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
