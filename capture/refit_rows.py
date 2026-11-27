"""Capture what Excel does with a row's stored height when the row has no ``customHeight``.

    uv run --group capture python -m capture.refit_rows build PATH  # the workbooks only
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.refit_rows run
    uv run --group capture python -m capture.refit_rows read  # the saved answer again

Excel for Windows lays a row whose ``ht`` has no ``customHeight`` out at 15 pt beside a Calibri
11 cell, whatever the ``ht`` (``windows/stored_rows``): it refits the row, or gives it the
default height. This suite tells refitted, default and stored apart: every row stores an ``ht``
without ``customHeight``, below, equal to and above the height its cell fits, with a Calibri 11,
20 or 28 cell in column A or no cell. Excel for Mac reads it by AppleScript (``run``), Excel for
Windows by the job group ``refit_rows`` (``capture/job.py``).

- ``refit``: a row a case (``CASES``), from row 1.
- ``macsaved``: rows as Excel for Mac writes a sheet it fitted: Calibri 11, 20 and 28 cells with
  the heights the Mac fits them (15, 26, 37), no ``customHeight``, an unlisted row between, and
  an empty row element at 15.
- ``zero``: a height of 0 without ``customHeight``, as ``ht="0"`` and as ``ht=""`` (which Excel
  reads as 0), on an empty row and beside a Calibri 20 cell; then the same two with
  ``customHeight="1"``, which Excel shows hidden (xlio's capture ``live_reads``).
- ``customsheet``: ``<sheetFormatPr defaultRowHeight="20" customHeight="1">``, where Excel fits no
  row: an ``ht`` without ``customHeight`` beside Calibri 11 and 28 cells and on an empty row, and an
  unlisted row (the default, 20).
- ``zerosheet``: ``<sheetFormatPr defaultRowHeight="15" zeroHeight="1">``, where unlisted rows are
  hidden: an ``ht`` without ``customHeight`` beside Calibri 11 and 20 cells and on an empty row, a
  Calibri 11 row without ``ht``, and an unlisted row (hidden).
- ``excelform``: rows with these attributes (``spans``,
  ``x14ac:dyDescent``, the sheet's ``defaultRowHeight="12.75"`` and ``x14ac:dyDescent``, the
  workbook's ``fileVersion`` and ``AppVersion`` 15.0300) under an Arial 10 Normal font, storing
  heights other than Excel for Windows' fit at 96 dpi: 13.5, 12.75 and 30 beside Arial 10 cells,
  13.5 on an empty row and beside an Arial 12 cell.
- ``excel10``, ``excel11``, ``excel8``: the same form under an Arial 10, 11 and 8 Normal font, each
  row a few cells in its own fonts (``EXCEL_ROWS``), with values or empty and styled, storing a
  height other than sheetpeek's fit: smaller fonts than the Normal font's, Arial bold, Book Antiqua
  and Tahoma, empty styled cells only; and in
  ``excel8`` a column style in Arial 10 from column B on, beside rows with ``customFormat="1"`` and
  no ``s``, with ``customFormat`` and A's style, and with neither.

``run`` saves Excel's answer to ``capture/build/refit_rows/answer.tsv`` and writes
``tests/captures/refit_rows.tsv``.
"""

from __future__ import annotations

import datetime as dt
import re
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

from capture import stored_rows, wrap_band, xmlform

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "refit_rows.tsv"
BUILD_DIR = ROOT / "capture" / "build" / "refit_rows"
ANSWER = BUILD_DIR / "answer.tsv"
CONTAINER_DIR = wrap_band.CONTAINER_DIR.parent / "refit_rows"
COLUMNS = ("workbook", "name", "row", "cell", "stored", "row_height", "height", "top")


@dataclass(frozen=True)
class Case:
    """One row: its cell (``None``, empty; else the Calibri size), stored ``ht`` (``None``, no
    ``ht``; with no cell either, no row element), the ``ht`` as written when it is not the
    number's own text, and whether the row keeps ``customHeight="1"``."""

    workbook: str
    row: int
    size: int | None
    stored: float | None
    text: str | None = None
    custom: bool = False
    font: str = "Calibri"
    label: str | None = None
    cells: tuple[tuple[str, int, bool, bool], ...] = ()
    """For ``EXCEL_ROWS``: the cells from A, each (face, size, bold, has a value)."""
    row_format: str = ""
    """For ``EXCEL_ROWS``: ``<row>`` attributes before ``ht``, such as ``customFormat="1"``; the
    text ``{s}`` stands for A's style index."""

    @property
    def written(self) -> str | None:
        """The ``ht`` attribute's text, or ``None`` for no ``ht``."""
        if self.text is not None:
            return self.text
        return None if self.stored is None else f"{self.stored:g}"

    @property
    def name(self) -> str:
        if self.label is not None:
            return f"{self.workbook} {self.label}"
        cell = "empty" if self.size is None else f"{self.font.lower()} {self.size}"
        written = self.written
        if written is None:
            stored = (
                "no ht" if self.size is not None and self.workbook in SHEET_FORMATS else "unlisted"
            )
        elif written == "":
            stored = 'ht ""'
        else:
            stored = f"ht {written}"
        custom = " custom" if self.custom else ""
        return f"{self.workbook} {cell} {stored}{custom}"


_REFIT = (
    (11, 10.0), (11, 15.0), (11, 30.0),
    (20, 12.0), (20, 26.0), (20, 26.25), (20, 40.0),
    (28, 20.0), (28, 60.0),
    (None, 10.0), (None, 15.0), (None, 30.0),
)  # fmt: skip
_MACSAVED = ((11, 15.0), (20, 26.0), (None, None), (28, 37.0), (None, 15.0))
_ZERO = (
    (None, "0", False), (20, "0", False), (None, "", False), (20, "", False),
    (None, "0", True), (20, "", True),
)  # fmt: skip
_CUSTOMSHEET = ((11, 30.0), (28, 30.0), (None, 30.0), (11, 10.0), (None, None))
_ZEROSHEET = ((11, 30.0), (20, 10.0), (None, 30.0), (11, None), (None, None))
_EXCELFORM = ((10, 13.5), (10, 12.75), (None, 13.5), (12, 13.5), (10, 30.0))
EXCELFORM_DESCENT = {12: "0.25"}
EXCEL_NORMAL = {
    "excelform": ("Arial", 10, "12.75"),
    "excel10": ("Arial", 10, "12.75"),
    "excel11": ("Arial", 11, "14.25"),
    "excel8": ("Arial", 8, "11.25"),
}
"""Each Excel-form workbook's Normal font and its sheet's ``defaultRowHeight``."""
_A10, _A10B, _A11B = ("Arial", 10, False), ("Arial", 10, True), ("Arial", 11, True)
EXCEL_ROWS: dict[str, list[tuple[str, float, tuple[tuple[str, int, bool, bool], ...], str]]] = {
    # (label, stored ht, cells from A: face, size, bold, has a value; row attributes)
    "excel10": [
        ("arial 9 beside empty arial 10", 12.0, (("Arial", 9, False, True), (*_A10, False)), ""),
        (
            "arial 8 and 8 bold beside empty arial 10",
            12.0,
            (("Arial", 8, False, True), ("Arial", 8, True, True), (*_A10, False)),
            "",
        ),
        (
            "empty book antiqua 8, 9, 9 bold",
            15.0,
            (
                ("Book Antiqua", 8, False, False),
                ("Book Antiqua", 9, False, False),
                ("Book Antiqua", 9, True, False),
            ),
            "",
        ),
        (
            "book antiqua 9 beside empty book antiqua 10",
            15.0,
            (("Book Antiqua", 9, False, True), ("Book Antiqua", 10, False, False)),
            "",
        ),
        (
            "empty tahoma 10 and 12 bold",
            15.0,
            (("Tahoma", 10, False, False), ("Tahoma", 12, True, False)),
            "",
        ),
        (
            "tahoma 20 bold beside empty tahoma 10",
            25.5,
            (("Tahoma", 20, True, True), ("Tahoma", 10, False, False)),
            "",
        ),
        (
            "empty tahoma 10 and 11 bold",
            14.25,
            (("Tahoma", 10, False, False), ("Tahoma", 11, True, False)),
            "",
        ),
        ("book antiqua 10", 15.0, (("Book Antiqua", 10, False, True),), ""),
        ("tahoma 12 bold", 15.0, (("Tahoma", 12, True, True),), ""),
    ],
    "excel11": [
        (
            "arial 10, 10 bold, 11 bold",
            15.75,
            ((*_A10, True), (*_A10B, True), (*_A11B, True)),
            "",
        ),
        ("arial 11 bold", 15.75, ((*_A11B, True),), ""),
        ("arial 10 bold", 15.0, ((*_A10B, True),), ""),
        ("arial 12 bold", 16.5, (("Arial", 12, True, True),), ""),
        ("empty arial 10 and 10 bold", 15.0, ((*_A10, False), (*_A10B, False)), ""),
        ("arial 11", 15.0, (("Arial", 11, False, True),), ""),
    ],
    "excel8": [
        ("custom format no s", 11.25, (("Arial", 8, False, True),), 'customFormat="1"'),
        ("no custom format", 11.25, (("Arial", 8, False, True),), ""),
        ("custom format s", 11.25, (("Arial", 8, False, True),), 's="{s}" customFormat="1"'),
        ("empty arial 8", 12.0, (("Arial", 8, False, False),), 'customFormat="1"'),
        ("arial 8 at 12.75", 12.75, (("Arial", 8, False, True),), 'customFormat="1"'),
    ],
}
EXCEL_COLUMN_STYLE = {"excel8": ("Arial", 10)}
"""The font of a style given to columns B to XFD."""
"""Excel's ``x14ac:dyDescent`` of a row by its cell's size, else ``0.2`` (Arial 10's)."""
SHEET_FORMATS = {
    "customsheet": 'baseColWidth="8" defaultRowHeight="20" customHeight="1"',
    "zerosheet": 'baseColWidth="8" defaultRowHeight="15" zeroHeight="1"',
}
"""The ``<sheetFormatPr>`` attributes of a workbook that does not keep openpyxl's."""
CASES = {
    "refit": [Case("refit", k, size, h) for k, (size, h) in enumerate(_REFIT, 1)],
    "macsaved": [Case("macsaved", k, size, h) for k, (size, h) in enumerate(_MACSAVED, 1)],
    "zero": [
        Case("zero", k, size, 0.0, text, custom) for k, (size, text, custom) in enumerate(_ZERO, 1)
    ],
    "customsheet": [Case("customsheet", k, size, h) for k, (size, h) in enumerate(_CUSTOMSHEET, 1)],
    "zerosheet": [Case("zerosheet", k, size, h) for k, (size, h) in enumerate(_ZEROSHEET, 1)],
    "excelform": [
        Case("excelform", k, size, h, font="Arial") for k, (size, h) in enumerate(_EXCELFORM, 1)
    ],
    **{
        name: [
            Case(name, k, cells[0][1], h, font=cells[0][0], label=label, cells=cells, row_format=f)
            for k, (label, h, cells, f) in enumerate(rows, 1)
        ]
        for name, rows in EXCEL_ROWS.items()
    },
}


def workbook_name(name: str) -> str:
    return f"refit_rows_{name}.xlsx"


def build_one(name: str, folder: Path) -> Path:
    """Write one workbook, every ``ht`` without ``customHeight``; return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for case in CASES[name]:
        if case.cells:
            for column, (face, size, bold, valued) in enumerate(case.cells, 1):
                cell = ws.cell(case.row, column, "x" if valued else None)
                cell.font = Font(name=face, sz=size, b=bold, family=2)
        elif case.size is not None:
            ws.cell(case.row, 1, "x").font = Font(name=case.font, sz=case.size, family=2)
        if case.stored is not None:
            # openpyxl leaves out a height of 0; _rewrite writes every ``ht`` as the case does
            ws.row_dimensions[case.row].height = case.stored or 1.0
    column_style = None
    if name in EXCEL_COLUMN_STYLE:
        face, size = EXCEL_COLUMN_STYLE[name]
        spare = ws.cell(1, 50)  # gives the style an index, then goes
        spare.font = Font(name=face, sz=size, family=2)
        column_style = spare.style_id
        del ws._cells[(1, 50)]
    path = folder / workbook_name(name)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    path.write_bytes(stored_rows._without_custom_height(path.read_bytes()))
    if name in EXCEL_NORMAL:
        _excel_form(path, CASES[name], EXCEL_NORMAL[name], column_style)
    elif name not in ("refit", "macsaved"):
        _rewrite(path, CASES[name], SHEET_FORMATS.get(name))
    return path


_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _excel_form(
    path: Path, cases: list[Case], normal: tuple[str, int, str], column_style: int | None
) -> None:
    """The workbook with these attributes: rows with ``spans`` and
    ``x14ac:dyDescent``, the sheet's default height, ``normal``'s face and size as the Normal font,
    a style for columns B to XFD, Excel's ``fileVersion`` and ``AppVersion``; the XML in the
    standard library's form."""
    face, size, default = normal
    edits = {
        "xl/worksheets/sheet1.xml": lambda text: _excel_rows(text, cases, default, column_style),
        "xl/styles.xml": lambda text: _replace_once(
            text,
            '<font><name val="Calibri" /><family val="2" /><color theme="1" /><sz val="11" />'
            '<scheme val="minor" /></font>',
            f'<font><name val="{face}" /><family val="2" /><color theme="1" /><sz val="{size}" />'
            "</font>",
        ),
        "xl/workbook.xml": lambda text: _replace_once(
            text,
            "<workbookPr />",
            '<fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420" />'
            '<workbookPr defaultThemeVersion="153222" />',
        ),
        "docProps/app.xml": lambda text: re.sub(
            r"<Application>[^<]*</Application><AppVersion>[^<]*</AppVersion>",
            "<Application>Microsoft Excel</Application><AppVersion>15.0300</AppVersion>",
            text,
        ),
    }
    with zipfile.ZipFile(path) as archive:
        items = [(info, archive.read(info.filename)) for info in archive.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for info, data in items:
            if info.filename in edits:
                text = xmlform.standard_form(data.decode("utf-8"))
                data = edits[info.filename](text).encode("utf-8")
            archive.writestr(info, data)


def _replace_once(text: str, old: str, new: str) -> str:
    if text.count(old) != 1:
        sys.exit(f"expected one {old!r}, found {text.count(old)}")
    return text.replace(old, new)


def _excel_rows(text: str, cases: list[Case], default: str, column_style: int | None) -> str:
    text = _replace_once(
        text,
        f'<worksheet xmlns="{_MAIN}">',
        f'<worksheet xmlns="{_MAIN}" xmlns:mc="{_MC}" mc:Ignorable="x14ac" xmlns:x14ac="{_X14AC}">',
    )
    text = _replace_once(
        text,
        '<sheetFormatPr baseColWidth="8" defaultRowHeight="15" />',
        f'<sheetFormatPr defaultRowHeight="{default}" x14ac:dyDescent="0.2" />',
    )
    if column_style is not None:
        text = _replace_once(
            text,
            "<sheetData>",
            f'<cols><col min="2" max="16384" width="9.140625" style="{column_style}" /></cols>'
            "<sheetData>",
        )
    for case in cases:
        descent = EXCELFORM_DESCENT.get(case.size or 10, "0.2")
        spans = f"1:{max(1, len(case.cells))}"
        before = ""
        if case.row_format:
            style = re.search(rf'<c r="A{case.row}" s="(\d+)"', text)
            before = " " + case.row_format.replace("{s}", style.group(1) if style else "0")
        text, count = re.subn(
            rf'<row r="{case.row}" ht="[^"]*"',
            f'<row r="{case.row}" spans="{spans}"{before} ht="{case.written}" '
            f'x14ac:dyDescent="{descent}"',
            text,
        )
        if count != 1:
            sys.exit(f"expected one row {case.row} with ht, found {count}")
    return text


def _rewrite(path: Path, cases: list[Case], sheet_format: str | None) -> None:
    """Write each case's ``ht`` as its text, with ``customHeight="1"`` where it keeps it, and
    the sheet's ``<sheetFormatPr>``; the XML in the standard library's form."""
    with zipfile.ZipFile(path) as archive:
        items = [(info, archive.read(info.filename)) for info in archive.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for info, data in items:
            if info.filename == "xl/worksheets/sheet1.xml":
                text = xmlform.standard_form(data.decode("utf-8"))
                for case in cases:
                    if case.written is None:
                        continue
                    custom = ' customHeight="1"' if case.custom else ""
                    text, count = re.subn(
                        rf'<row r="{case.row}" ht="[^"]*"',
                        f'<row r="{case.row}" ht="{case.written}"{custom}',
                        text,
                    )
                    if count != 1:
                        sys.exit(f"{path.name}: expected one row {case.row} with ht, found {count}")
                if sheet_format is not None:
                    text, count = re.subn(
                        r"<sheetFormatPr\b[^>]*/>", f"<sheetFormatPr {sheet_format} />", text
                    )
                    if count != 1:
                        sys.exit(f"{path.name}: expected one <sheetFormatPr />, found {count}")
                data = text.encode("utf-8")
            archive.writestr(info, data)


def build(folder: Path = CONTAINER_DIR) -> list[Path]:
    return [build_one(name, folder) for name in CASES]


def run() -> None:
    """Build, have Excel for Mac read the rows, save its answer, write the TSV."""
    from capture.edition import excel_description

    args = []
    for name, path in zip(CASES, build(), strict=True):
        args += [str(path), ",".join(str(c.row) for c in CASES[name])]
    try:
        result = subprocess.run(
            ["osascript", str(wrap_band.SCRIPT), *args],
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


def records(answer: str) -> list[dict[str, str]]:
    """The records of a Mac answer (``wrap_band.applescript``'s lines), in case order."""
    body = [line for line in answer.strip().split("\n") if line and not line.startswith("# ")]
    every = [case for name in CASES for case in CASES[name]]
    if len(body) != len(every):
        raise ValueError(f"expected {len(every)} lines from Excel, got {len(body)}")
    found = []
    for case, line in zip(every, body, strict=True):
        name, row, _, _, height, _, _ = line.split("\t")
        if (name, int(row)) != (workbook_name(case.workbook), case.row):
            raise ValueError(f"Excel answered for {name} row {row}, expected {case.name}")
        found.append(record(case, height, "", ""))
    return found


def record(case: Case, row_height: str, height: str, top: str) -> dict[str, str]:
    return {
        "workbook": case.workbook,
        "name": case.name,
        "row": str(case.row),
        "cell": "" if case.size is None else f"{case.font} {case.size}",
        "stored": "" if case.written is None else case.written or '""',
        "row_height": wrap_band._number(row_height),
        "height": height,
        "top": top,
    }


def write_records(head: list[str], found: list[dict[str, str]], output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8", newline="\n") as out:
        out.write("".join(f"{line}\n" for line in head))
        out.write("\t".join(COLUMNS) + "\n")
        for found_record in found:
            out.write("\t".join(found_record[c] for c in COLUMNS) + "\n")
    print(f"{len(found)} rows written to {output}")


def write(answer: Path = ANSWER, output: Path = OUTPUT) -> None:
    """Write ``refit_rows.tsv`` from Excel for Mac's saved answer (no Excel)."""
    text = answer.read_text(encoding="utf-8")
    head = [line for line in text.split("\n") if line.startswith("# ")]
    write_records(
        [*head, "# suite: refit_rows (capture.refit_rows); method: row height (AppleScript)"],
        records(text),
        output,
    )


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        for path in build(Path(argv[1])):
            print(path)
        return 0
    if argv == ["run"]:
        run()
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[0], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
