"""Capture the row heights Excel for Mac writes when it saves a workbook it fitted.

    uv run --group capture python -m capture.saved_rows build PATH  # the workbook only
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.saved_rows run

Excel opens a workbook of one cell a row (``ROWS``), each row without ``customHeight``, with or
without an ``ht``, and one with ``customHeight``, and saves a copy (``save_copy.applescript``);
Excel for Windows does the same by the job group ``saved_rows`` (``capture/job.py``). The copy's
``<row>`` elements give, for each row, the ``ht`` Excel wrote and whether it wrote
``customHeight``. It shows what a file Excel saved stores for a row it fitted, and, by a row in
``arial`` beside one in ``Arial``, whether Excel matches a face name without regard to case
(the same fit, or another face's). ``run`` writes ``tests/captures/saved_rows.tsv``; the job,
``tests/captures/windows/saved_rows.tsv``.
"""

from __future__ import annotations

import datetime as dt
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

from capture import wrap_band, xmlform

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "saved_rows.tsv"
CONTAINER_DIR = wrap_band.CONTAINER_DIR.parent / "saved_rows"
SCRIPT = Path(__file__).with_name("save_copy.applescript")
COLUMNS = ("row", "name", "face", "size", "bold", "stored", "custom", "saved_ht", "saved_custom")
ROWS: tuple[tuple[str, str, float, bool, float | None, bool], ...] = (
    # (name, face, size, bold, ht written, customHeight written)
    ("calibri 11", "Calibri", 11, False, None, False),
    ("calibri 20", "Calibri", 20, False, None, False),
    ("calibri 28 ht 60", "Calibri", 28, False, 60.0, False),
    ("arial 20", "Arial", 20, False, None, False),
    ("arial 20 lower case", "arial", 20, False, None, False),
    ("tahoma 12 bold ht 15", "Tahoma", 12, True, 15.0, False),
    ("calibri 11 ht 30 custom", "Calibri", 11, False, 30.0, True),
)


def build(folder: Path = CONTAINER_DIR) -> Path:
    """The workbook Excel opens: a cell a row, the heights as ``ROWS`` writes them."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for row, (_, face, size, bold, height, _custom) in enumerate(ROWS, 1):
        ws.cell(row, 1, "x").font = Font(name=face, sz=size, b=bold, family=2)
        if height is not None:
            ws.row_dimensions[row].height = height
    path = folder / "saved_rows_source.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    _drop_custom_height(path)
    return path


def _drop_custom_height(path: Path) -> None:
    """Take ``customHeight`` off every row ``ROWS`` writes without it."""
    with zipfile.ZipFile(path) as archive:
        items = [(info, archive.read(info.filename)) for info in archive.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for info, data in items:
            if info.filename == "xl/worksheets/sheet1.xml":
                text = xmlform.standard_form(data.decode("utf-8"))
                for row, (*_, custom) in enumerate(ROWS, 1):
                    if not custom:
                        text = re.sub(rf'(<row r="{row}" [^>]*?) customHeight="1"', r"\1", text)
                data = text.encode("utf-8")
            archive.writestr(info, data)


def saved(path: Path) -> dict[int, tuple[str, str]]:
    """Each row of the copy's first sheet: (``ht`` as written, ``customHeight`` as written)."""
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
    found = {}
    for element in re.findall(r"<row [^>]*>", xml):
        attrs = dict(re.findall(r'([\w:]+)="([^"]*)"', element))
        found[int(attrs["r"])] = (attrs.get("ht", ""), attrs.get("customHeight", ""))
    return found


def write(copy: Path, output: Path = OUTPUT, head: list[str] | None = None) -> None:
    """The TSV of ``copy``'s rows, under ``head`` (else this Mac's Excel and today)."""
    if head is None:
        from capture.edition import excel_description

        head = [f"# excel: {excel_description()}, en_US", f"# captured: {dt.date.today()}"]
    rows = saved(copy)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8", newline="\n") as out:
        out.write("".join(f"{line}\n" for line in head))
        out.write(
            "# suite: saved_rows (capture.saved_rows); method: the rows of a copy Excel saved\n"
        )
        out.write("\t".join(COLUMNS) + "\n")
        for row, (name, face, size, bold, height, custom) in enumerate(ROWS, 1):
            saved_ht, saved_custom = rows.get(row, ("", ""))
            fields = [
                str(row), name, face, f"{size:g}", "1" if bold else "0",
                "" if height is None else f"{height:g}", "1" if custom else "0",
                saved_ht, saved_custom,
            ]  # fmt: skip
            out.write("\t".join(fields) + "\n")
    print(f"{len(ROWS)} rows written to {output}")


def run() -> None:
    """Build, have Excel for Mac open the workbook and save a copy, write the TSV."""
    source = build()
    copy = source.with_name("saved_rows_copy.xlsx")
    if copy.exists():
        copy.unlink()
    try:
        result = subprocess.run(
            ["osascript", str(SCRIPT), str(source), str(copy)],
            capture_output=True,
            text=True,
            check=False,
            timeout=120,  # kills osascript, which ignores SIGTERM while it waits on Excel
        )
    except subprocess.TimeoutExpired:
        sys.exit("osascript did not answer in 2 minutes; check Excel before another run")
    if result.returncode != 0:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    write(copy)


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        print(build(Path(argv[1])))
        return 0
    if argv == ["run"]:
        run()
        return 0
    print(__doc__.split("\n\n")[0], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
