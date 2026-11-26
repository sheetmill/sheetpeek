"""How Excel reads a style value outside what it writes: an odd angle, an undefined format id.

    uv run --group capture python -m capture.odd_values build [DIR]  # write the workbooks only
    uv run --group capture python -m capture.odd_values run          # open each in Excel, write TSV

Not a screen suite: Excel's own properties answer, read back as ``capture.package_probe`` reads
them. Each case is one small workbook, written by openpyxl and then changed in ``xl/styles.xml``
where openpyxl cannot write the value: ``Odd value`` in A1, with an alignment whose
``textRotation`` is the case's, and 1234.5 in B1 (column B 20 characters wide, so no format
has to shorten it), whose format's ``numFmtId`` is the case's and is not defined in
``<numFmts>`` unless it is a built-in id. ``odd_values_open.applescript`` opens each one in its
own run, under a time limit, and reads A1's text, A1's ``text orientation``, B1's ``number
format`` and the text B1 shows (``string value``), or records Excel's refusal: a refusal is the
answer for that value. A control (45 degrees, built-in format 4, ``#,##0.00``) opens first and
last: an outcome counts only between two that open. The output is
``tests/captures/odd_values.tsv``. New cases go before the last control.

What the capture settles (``ARCHITECTURE.md``):
- §4.17, the reading of ``textRotation``: a value other than 0 to 180 and 255 reads as 0
  (``sheetpeek._style.alignment``; Excel opens each horizontal). Cases 181, 254, 256 and 1000.
- §5, reading: a custom ``numFmtId`` (164 and up) the file does not define reads as General
  (``sheetpeek._read``; Excel shows each as General). Cases 170 and 300.

``EXPECTED`` is what sheetpeek reads for each case today (taken by running it, see
``sheetpeek_reads``), not a guess at Excel: the capture shows where the two differ.
"""

from __future__ import annotations

import datetime as dt
import io
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment

from capture.edition import excel_description

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "odd_values.tsv"
CONTAINER_DIR = (
    Path.home()
    / "Library"
    / "Containers"
    / "com.microsoft.Excel"
    / "Data"
    / "Documents"
    / "sheetpeek"
    / "odd_values"
)
STYLES = "xl/styles.xml"
TIME_LIMIT = "60"
TEXT = "Odd value"
NUMBER = 1234.5
WIDTH_B = 20
CONTROL_ROTATION = 45
CONTROL_FORMAT_ID = 4
CONTROL_FORMAT = "#,##0.00"
COLUMNS = ("case", "outcome", "A1", "orientation", "number_format", "text")
"""One row a case: ``outcome`` is ``opened``, ``repaired`` (opened under another name) or
``refused``; for a refusal ``A1`` holds Excel's error number and the rest is empty."""


@dataclass(frozen=True)
class Case:
    """One workbook: A1's ``textRotation`` and B1's ``numFmtId``, as stored."""

    name: str
    rotation: int
    format_id: int


CASES: tuple[Case, ...] = (
    Case("control", CONTROL_ROTATION, CONTROL_FORMAT_ID),
    Case("rotation_181", 181, CONTROL_FORMAT_ID),
    Case("rotation_254", 254, CONTROL_FORMAT_ID),
    Case("rotation_256", 256, CONTROL_FORMAT_ID),
    Case("rotation_1000", 1000, CONTROL_FORMAT_ID),
    Case("format_170", CONTROL_ROTATION, 170),
    Case("format_300", CONTROL_ROTATION, 300),
    Case("control_again", CONTROL_ROTATION, CONTROL_FORMAT_ID),
)

EXPECTED: dict[str, tuple[int, bool, str, str]] = {
    "control": (45, False, CONTROL_FORMAT, "1,234.50"),
    "rotation_181": (0, False, CONTROL_FORMAT, "1,234.50"),
    "rotation_254": (0, False, CONTROL_FORMAT, "1,234.50"),
    "rotation_256": (0, False, CONTROL_FORMAT, "1,234.50"),
    "rotation_1000": (0, False, CONTROL_FORMAT, "1,234.50"),
    "format_170": (45, False, "General", "1234.5"),
    "format_300": (45, False, "General", "1234.5"),
    "control_again": (45, False, CONTROL_FORMAT, "1,234.50"),
}
"""What sheetpeek reads for each case today: A1's rotation and stacked flag, B1's format code
and the text it shows."""


def base() -> bytes:
    """The control as openpyxl writes it: A1 at 45 degrees, B1 in built-in format 4."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = TEXT
    sheet["A1"].alignment = Alignment(textRotation=CONTROL_ROTATION)
    sheet["B1"] = NUMBER
    sheet["B1"].number_format = CONTROL_FORMAT
    sheet.column_dimensions["B"].width = WIDTH_B
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _patch(data: bytes, edits: list[tuple[str, str]]) -> bytes:
    """A copy of the package with each piece of ``xl/styles.xml`` replaced; each must be unique."""
    out = io.BytesIO()
    with (
        zipfile.ZipFile(io.BytesIO(data)) as source,
        zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target,
    ):
        for info in source.infolist():
            body = source.read(info.filename)
            if info.filename == STYLES:
                text = body.decode()
                for old, new in edits:
                    if text.count(old) != 1:
                        raise ValueError(f"{old!r} is not in {STYLES} exactly once")
                    text = text.replace(old, new)
                body = text.encode()
            target.writestr(info.filename, body)
    return out.getvalue()


def workbook(case: Case, data: bytes | None = None) -> bytes:
    """The case's workbook: the control with its two values written into the styles."""
    data = base() if data is None else data
    return _patch(
        data,
        [
            (f'textRotation="{CONTROL_ROTATION}"', f'textRotation="{case.rotation}"'),
            (f'numFmtId="{CONTROL_FORMAT_ID}"', f'numFmtId="{case.format_id}"'),
        ],
    )


def sheetpeek_reads(data: bytes) -> tuple[int, bool, str, str]:
    """What sheetpeek reads from a case: A1's rotation and stacked flag, B1's format and text."""
    import sheetpeek

    sheet = sheetpeek.read(io.BytesIO(data), ref="A1:B1", platform="mac")
    a1, b1 = sheet.cell(1, 1), sheet.cell(1, 2)
    shown = sheet.layout("excel").cell(1, 2)
    text = shown.text if shown is not None else ""
    return a1.style.rotation, a1.style.stacked, b1.style.number_format, text


def degrees(orientation: str) -> tuple[int, bool]:
    """Excel's ``text orientation`` as sheetpeek's (rotation, stacked).

    Excel answers an angle as its number and four values by name: horizontal (0), upward (90),
    downward (-90) and vertical (stacked).
    """
    named = {
        "orientation horizontal": (0, False),
        "orientation upward": (90, False),
        "orientation downward": (-90, False),
        "orientation vertical": (0, True),
    }
    if orientation in named:
        return named[orientation]
    try:
        return int(float(orientation)), False
    except ValueError:
        raise ValueError(f"not an orientation Excel reports: {orientation!r}") from None


def build(directory: Path = CONTAINER_DIR) -> list[Path]:
    """Write one workbook per case into ``directory`` (Excel's container unless given)."""
    directory.mkdir(parents=True, exist_ok=True)
    data = base()
    paths = []
    for case in CASES:
        path = directory / f"odd_values_{case.name}.xlsx"
        path.write_bytes(workbook(case, data))
        paths.append(path)
    return paths


def parse(name: str, answer: str) -> dict[str, str]:
    """One row of the TSV from what ``odd_values_open.applescript`` printed for a case."""
    fields = answer.split("\t")
    if fields[0] not in ("opened", "repaired", "refused"):
        raise ValueError(f"{name}: not an answer the script gives: {answer!r}")
    values = [name, *fields][: len(COLUMNS)]
    values += [""] * (len(COLUMNS) - len(values))
    return dict(zip(COLUMNS, values, strict=True))


def _open(path: Path) -> str:
    """One case in its own ``osascript``; alerts are restored even when the limit stops it."""
    script = Path(__file__).with_name("odd_values_open.applescript")
    # -k: osascript waiting for an answer from Excel does not stop for SIGTERM.
    result = subprocess.run(
        ["gtimeout", "-k", "5", TIME_LIMIT, "osascript", str(script), str(path)],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode in (124, 137):
        restore = Path(__file__).with_name("restore_alerts.applescript")
        subprocess.run(["gtimeout", "-k", "5", "20", "osascript", str(restore)], check=False)
        return f"timed out after {TIME_LIMIT} s"
    if result.returncode != 0:
        raise SystemExit(f"{path.name}: {result.stderr.strip()}")
    return result.stdout.strip()


def write(rows: list[dict[str, str]], description: str, path: Path | None = None) -> None:
    """Write the rows with the header lines ``capture.package_probe`` writes."""
    path = OUTPUT if path is None else path
    with path.open("w", encoding="utf-8") as out:
        out.write(f"# suite: odd_values (capture.odd_values); {description}, en_US\n")
        out.write(f"# captured: {dt.date.today().isoformat()}; method: open with alerts off, ")
        out.write("read A1's text and text orientation, B1's number format and the text B1 ")
        out.write("shows (string value); a refusal's error number in A1\n")
        out.write("\t".join(COLUMNS) + "\n")
        for row in rows:
            out.write("\t".join(row[c] for c in COLUMNS) + "\n")


def run() -> None:
    """Open every case, print the outcome and write the TSV. A case that hangs ends the run."""
    paths = build()
    rows = []
    for case, path in zip(CASES, paths, strict=True):
        answer = _open(path)
        print(f"{case.name}\t{answer}", flush=True)
        if answer.startswith("timed out"):
            sys.exit(
                "a case hung: stopped, and nothing written. Look into it before running again."
            )
        rows.append(parse(case.name, answer))
    for control in (rows[0], rows[-1]):
        if control["outcome"] != "opened":
            sys.exit(f"the control did not open ({control}): nothing written")
    write(rows, excel_description())
    print(f"{len(rows)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if argv[:1] == ["build"]:
        for path in build(Path(argv[1]) if len(argv) > 1 else CONTAINER_DIR):
            print(path)
    elif argv[:1] == ["run"]:
        run()
    else:
        print(__doc__)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
