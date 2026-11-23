"""Capture the text Excel for Mac shows in merged cells and for error values (ARCHITECTURE.md §4.9).

    uv run --group capture python -m capture.cell_text run

Two suites, read with ``string value`` (``string_values.applescript``), which is the text the
cell shows at its width:

- ``merge_width``: one sheet per column width (file widths 2.71 to 8.71, every column alike);
  each row holds a number in A, alone or merged across two, three or four columns, or across
  three with the middle one hidden. Output ``tests/captures/merge_width.tsv``.
- ``error_width``: every error value a file can hold, in columns B to I of widths 1.71 to 8.71.
  Output ``tests/captures/error_width.tsv``.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from capture.edition import excel_description

ROOT = Path(__file__).resolve().parents[1]
CAPTURES = ROOT / "tests" / "captures"
CONTAINER_DIR = (
    Path.home() / "Library" / "Containers" / "com.microsoft.Excel" / "Data" / "Documents"
    / "sheetpeek"
)  # fmt: skip
WORKBOOK = CONTAINER_DIR / "cell_text.xlsx"
MERGE_WIDTHS = (2.7109375, 3.7109375, 4.7109375, 5.7109375, 6.7109375, 8.7109375)
# (value, format) shown in each merge
MERGE_VALUES = (
    (123456789, "#,##0.00"),
    (1234567.891, "#,##0.00"),
    (123456789012, "General"),
    (0.000123456789, "General"),
    (-98765.4321, "0"),
)
# (label, last column of the merge or None for a single cell, hidden column or None)
MERGES = (
    ("single", None, None),
    ("A:B", "B", None),
    ("A:C", "C", None),
    ("A:D", "D", None),
    ("A:C, B hidden", "C", "B"),
)
ERRORS = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A")
ERROR_WIDTHS = tuple(w + 0.7109375 for w in range(1, 9))


def build(path: Path = WORKBOOK) -> Path:
    """Write the merge sheets (``w<width>``) and the error sheet (``errors``)."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for width in MERGE_WIDTHS:
        ws = wb.create_sheet(f"w{width!r}")
        for letter in "ABCD":
            ws.column_dimensions[letter].width = width
        row = 0
        for _label, last, hidden in MERGES:
            if hidden:
                continue  # hidden columns apply to the whole sheet: on the h sheets below
            for value, fmt in MERGE_VALUES:
                row += 1
                ws.cell(row, 1, value).number_format = fmt
                if last:
                    ws.merge_cells(f"A{row}:{last}{row}")
    for width in MERGE_WIDTHS:
        ws = wb.create_sheet(f"h{width!r}")
        for letter in "ABCD":
            ws.column_dimensions[letter].width = width
        ws.column_dimensions["B"].hidden = True
        for row, (value, fmt) in enumerate(MERGE_VALUES, 1):
            ws.cell(row, 1, value).number_format = fmt
            ws.merge_cells(f"A{row}:C{row}")
    ws = wb.create_sheet("errors")
    for k, width in enumerate(ERROR_WIDTHS):
        ws.column_dimensions[get_column_letter(k + 2)].width = width
    for row, error in enumerate(ERRORS, 1):
        ws.cell(row, 1, error)
        for k in range(len(ERROR_WIDTHS)):
            ws.cell(row, k + 2, error)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def specs() -> list[str]:
    """The ranges whose text is read, in order: each merge sheet's A, then the error cells."""
    plain = [label for label, _, hidden in MERGES if hidden is None]
    rows = len(plain) * len(MERGE_VALUES)
    found = [f"w{w!r}!A1:A{rows}" for w in MERGE_WIDTHS]
    found += [f"h{w!r}!A1:A{len(MERGE_VALUES)}" for w in MERGE_WIDTHS]
    last = get_column_letter(len(ERROR_WIDTHS) + 1)
    found.append(f"errors!B1:{last}{len(ERRORS)}")
    return found


def run() -> None:
    """Build, have Excel read every cell's string value, write the two TSVs."""
    build()
    script = Path(__file__).with_name("string_values.applescript")
    result = subprocess.run(
        ["osascript", str(script), str(WORKBOOK), *specs()], capture_output=True, text=True,
        check=False,
    )  # fmt: skip
    if result.returncode != 0:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    shown: dict[tuple[str, int, int], str] = {}
    for line in result.stdout.rstrip("\n").split("\n"):
        spec, position, text = line.split("\t", 2)
        r, c = (int(v) for v in position.split(","))
        shown[(spec, r, c)] = text
    today = dt.date.today().isoformat()
    write_tsvs(
        shown,
        f"{excel_description()}, en_US",
        f"# captured: {today}; method: string value (AppleScript, bulk)",
        CAPTURES,
    )


def write_tsvs(
    shown: dict[tuple[str, int, int], str], excel: str, method: str, folder: Path
) -> list[Path]:
    """Write ``merge_width.tsv`` and ``error_width.tsv`` into ``folder`` from the text read.

    Args:
        shown: The text of each cell read, by (range of ``specs()``, row, column), 1-based.
        excel: Which Excel answered, for the first header line.
        method: The header line saying when and how the text was read.
        folder: Where the two files go.
    """
    plain = [label for label, _, hidden in MERGES if hidden is None]
    rows = len(plain) * len(MERGE_VALUES)
    last = get_column_letter(len(ERROR_WIDTHS) + 1)
    merge_path, error_path = folder / "merge_width.tsv", folder / "error_width.tsv"
    with merge_path.open("w", encoding="utf-8", newline="\n") as out:
        out.write(f"# suite: merge_width (capture.cell_text); {excel}\n")
        out.write("# A holds the value, alone or merged; every column of a sheet has the file\n")
        out.write("# width given; shown: string value of A\n")
        out.write(method + "\n")
        out.write("file_width\tmerge\tvalue\tformat\tshown\n")
        for width in MERGE_WIDTHS:
            for k, (label, value_index, _) in enumerate(
                [(m, v, "") for m in plain for v in range(len(MERGE_VALUES))], 1
            ):
                value, fmt = MERGE_VALUES[value_index]
                text = shown[(f"w{width!r}!A1:A{rows}", k, 1)]
                out.write(f"{width!r}\t{label}\t{value!r}\t{fmt}\t{text}\n")
            for k, (value, fmt) in enumerate(MERGE_VALUES, 1):
                text = shown[(f"h{width!r}!A1:A{len(MERGE_VALUES)}", k, 1)]
                out.write(f"{width!r}\tA:C, B hidden\t{value!r}\t{fmt}\t{text}\n")
    with error_path.open("w", encoding="utf-8", newline="\n") as out:
        out.write(f"# suite: error_width (capture.cell_text); {excel}\n")
        out.write("# each error value in columns of the file width given; shown: string value\n")
        out.write(method + "\n")
        out.write("error\tfile_width\tshown\n")
        for r, error in enumerate(ERRORS, 1):
            for c, width in enumerate(ERROR_WIDTHS, 1):
                text = shown[(f"errors!B1:{last}{len(ERRORS)}", r, c)]
                out.write(f"{error}\t{width!r}\t{text}\n")
    print(f"wrote {merge_path.name} and {error_path.name} in {folder}")
    return [merge_path, error_path]


if __name__ == "__main__":
    {"build": build, "run": run}[sys.argv[1]]()
