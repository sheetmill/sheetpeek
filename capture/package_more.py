"""Capture what Excel for Mac does with the sizes no column, row or font can have (§5.5), further.

    export EXCEL_EDITION="Microsoft 365"
    uv run --group capture python -m capture.package_more run
    uv run --group capture python -m capture.package_more read          # write the TSV again
    uv run --group capture python -m capture.package_more build FOLDER   # the workbooks only

The ``package`` suite (``capture/package_probe.py``) left four questions of ARCHITECTURE §5.5,
"Numbers no column, row or font can have", unverified; this suite asks them in the same way.
Each case is one small workbook, ``package_probe.base()`` (111 in A1 of its only sheet), with
one thing changed in its worksheet or styles part, built by ``package_probe``'s own edits:

- ``width_1e100``, ``width_1e200``, ``width_1e300``: a ``<col width>`` between 1e30 (one point,
  captured) and 1e308 (0, a hidden column, captured): where 0 starts.
- ``default_width_*``, ``base_width_*``: ``defaultColWidth`` and ``baseColWidth`` of 255, 256,
  299 and 300 (300 is refused for both, captured): where the refusal starts.
- ``font_size_*``: the Normal font's size 409, 409.5 and 410 (410 is refused, captured), and -2
  and -10 (whether a size below -1 is its magnitude, as -1 is).
- ``arial_10_*``, ``calibri_14_*``: the Normal font Arial 10 or Calibri 14 (``styles.xml``'s
  first font, without ``scheme``), alone and with a ``<col width>`` of -1, 1e6 and 1e30, whose
  widths with Calibri 11 were 10921, 6036.5 and 0.09 units: whether a width is kept in points
  modulo 65,536 at 6 pt a digit whatever the Normal font's digit, or in its own digits.

No window and no screenshot: ``package_more_open.applescript`` opens each workbook in its own
``osascript`` under a time limit, with alerts off, and reads A1, the number of sheets, the
column width (characters) and width (points) of columns A and B, the heights of rows 1 and 2,
and A1's font size and name; or the error number of a refusal, which is the answer for that
case. A control opens first and last: an outcome counts only between two that open. A case that
hangs ends the run and nothing is saved. Numbers go into the XML as written here, as text. ``run``
saves Excel's answers with the environment (``capture/build/package_more/excel.tsv``); ``read``
(:func:`write`) turns them into ``tests/captures/package_more.tsv`` without Excel. New cases go
before the last control.
"""

from __future__ import annotations

import datetime as dt
import os
import subprocess
import sys
from collections.abc import Callable
from pathlib import Path

from capture.package_probe import _edit, _font_size, _sheet_format, _width, base, control

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "package_more.tsv"
BUILD_DIR = ROOT / "capture" / "build" / "package_more"
SAVED = BUILD_DIR / "excel.tsv"
CONTAINER_DIR = (
    Path.home() / "Library" / "Containers" / "com.microsoft.Excel" / "Data" / "Documents"
    / "sheetpeek" / "package_more"
)  # fmt: skip
SCRIPT = Path(__file__).with_name("package_more_open.applescript")
RESTORE = Path(__file__).with_name("restore_alerts.applescript")
TIME_LIMIT = "60"
STYLES = "xl/styles.xml"
NORMAL_FONT = (
    '<name val="Calibri" /><family val="2" /><color theme="1" /><sz val="11" />'
    '<scheme val="minor" />'
)
"""openpyxl's first font, the Normal style's."""
OPENED = (
    "A1", "sheets", "width A", "width A pt", "height 1", "width B", "width B pt", "height 2",
    "font size", "font name",
)  # fmt: skip
"""What the script reads of a workbook Excel opens, in its order."""
COLUMNS = ("case", "outcome", *OPENED)


def _normal(name: str, size: str) -> Callable[[bytes], bytes]:
    """The Normal font ``name`` at ``size``, without a theme ``scheme`` that would replace it."""
    new = f'<name val="{name}" /><family val="2" /><color theme="1" /><sz val="{size}" />'
    return _edit(STYLES, NORMAL_FONT, new)


def _then(*edits: Callable[[bytes], bytes]) -> Callable[[bytes], bytes]:
    """The edits applied in turn."""

    def case(data: bytes) -> bytes:
        for edit in edits:
            data = edit(data)
        return data

    return case


def _base_width(value: str) -> Callable[[bytes], bytes]:
    return _edit("xl/worksheets/sheet1.xml", 'baseColWidth="8"', f'baseColWidth="{value}"')


CASES: tuple[tuple[str, Callable[[bytes], bytes]], ...] = (
    ("control", control),
    ("width_1e100", _width("1e100")),
    ("width_1e200", _width("1e200")),
    ("width_1e300", _width("1e300")),
    ("default_width_255", _sheet_format("defaultColWidth", "255")),
    ("default_width_256", _sheet_format("defaultColWidth", "256")),
    ("default_width_299", _sheet_format("defaultColWidth", "299")),
    ("default_width_300", _sheet_format("defaultColWidth", "300")),
    ("base_width_255", _base_width("255")),
    ("base_width_256", _base_width("256")),
    ("base_width_299", _base_width("299")),
    ("base_width_300", _base_width("300")),
    ("font_size_409", _font_size("409")),
    ("font_size_409.5", _font_size("409.5")),
    ("font_size_410", _font_size("410")),
    ("font_size_minus_2", _font_size("-2")),
    ("font_size_minus_10", _font_size("-10")),
    ("arial_10", _normal("Arial", "10")),
    ("arial_10_width_minus_1", _then(_normal("Arial", "10"), _width("-1"))),
    ("arial_10_width_1e6", _then(_normal("Arial", "10"), _width("1e6"))),
    ("arial_10_width_1e30", _then(_normal("Arial", "10"), _width("1e30"))),
    ("calibri_14", _normal("Calibri", "14")),
    ("calibri_14_width_minus_1", _then(_normal("Calibri", "14"), _width("-1"))),
    ("calibri_14_width_1e6", _then(_normal("Calibri", "14"), _width("1e6"))),
    ("calibri_14_width_1e30", _then(_normal("Calibri", "14"), _width("1e30"))),
    ("control_again", control),
)

EXPECTED: dict[str, tuple[float, float, float, float, float] | None] = {
    # What sheetpeek takes for each case today, by running it on the built workbook (in excel
    # mode, hidden columns shown): None where it refuses the file, else the widths of columns A
    # and B and the heights of rows 1 and 2 in points, and A1's font size. The capture shows
    # where Excel differs.
    "control": (53.0, 53.0, 15.0, 15.0, 11.0),
    "width_1e100": (1.0, 53.0, 15.0, 15.0, 11.0),
    "width_1e200": (1.0, 53.0, 15.0, 15.0, 11.0),
    "width_1e300": (1.0, 53.0, 15.0, 15.0, 11.0),
    "default_width_255": (1530.0, 1530.0, 15.0, 15.0, 11.0),
    "default_width_256": None,
    "default_width_299": None,
    "default_width_300": None,
    "base_width_255": (65.0, 65.0, 15.0, 15.0, 11.0),
    "base_width_256": None,
    "base_width_299": None,
    "base_width_300": None,
    "font_size_409": (1735.0, 1735.0, 31.0, 31.0, 409.0),
    "font_size_409.5": (1745.0, 1745.0, 31.0, 31.0, 409.5),
    "font_size_410": None,
    "font_size_minus_2": (11.0, 11.0, 12.0, 12.0, 2.0),
    "font_size_minus_10": (45.0, 45.0, 14.0, 14.0, 10.0),
    "arial_10": (45.0, 45.0, 14.0, 14.0, 10.0),
    "arial_10_width_minus_1": (54609.0, 45.0, 14.0, 14.0, 10.0),
    "arial_10_width_1e6": (30187.0, 45.0, 14.0, 14.0, 10.0),
    "arial_10_width_1e30": (0.0, 45.0, 14.0, 14.0, 10.0),
    "calibri_14": (61.0, 61.0, 19.0, 19.0, 14.0),
    "calibri_14_width_minus_1": (65530.0, 61.0, 19.0, 19.0, 14.0),
    "calibri_14_width_1e6": (53184.0, 61.0, 19.0, 19.0, 14.0),
    "calibri_14_width_1e30": (1.0, 61.0, 19.0, 19.0, 14.0),
    "control_again": (53.0, 53.0, 15.0, 15.0, 11.0),
}


def build(folder: Path = CONTAINER_DIR) -> list[Path]:
    """Write one workbook per case into ``folder``."""
    folder.mkdir(parents=True, exist_ok=True)
    data = base()
    paths = []
    for name, case in CASES:
        path = folder / f"package_more_{name}.xlsx"
        path.write_bytes(case(data))
        paths.append(path)
    return paths


def _open(path: Path) -> str:
    """One case in its own ``osascript``; alerts are restored even when the limit stops it."""
    # -k: osascript waiting for an answer from Excel does not stop for SIGTERM.
    result = subprocess.run(
        ["gtimeout", "-k", "5", TIME_LIMIT, "osascript", str(SCRIPT), str(path)],
        capture_output=True, text=True, check=False,
    )  # fmt: skip
    if result.returncode in (124, 137):
        subprocess.run(["gtimeout", "-k", "5", "20", "osascript", str(RESTORE)], check=False)
        return f"timed out after {TIME_LIMIT} s"
    if result.returncode != 0:
        raise SystemExit(f"{path.name}: {result.stderr.strip()}")
    return result.stdout.strip()


def run() -> None:
    """Open every case, save Excel's answers, write the TSV. A case that hangs ends the run."""
    from capture.screen.session import environment

    edition = os.environ.get("EXCEL_EDITION")
    if not edition:
        raise SystemExit("state the Excel edition (Excel > About Excel) in EXCEL_EDITION")
    found = environment(edition)
    paths = build()
    lines = []
    try:
        for (name, _), path in zip(CASES, paths, strict=True):
            outcome = _open(path)
            print(f"{name}\t{outcome}", flush=True)
            if outcome.startswith("timed out"):
                sys.exit(f"{name} hung: stopped, and nothing saved. Look into it first.")
            lines.append(f"{name}\t{outcome}")
    finally:
        for path in paths:
            path.unlink(missing_ok=True)
    answer = "\n".join(lines)
    measure(answer)  # stops on an answer that does not fit before anything is saved
    found["date"] = dt.date.today().isoformat()
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    with SAVED.open("w", encoding="utf-8", newline="\n") as out:
        for key, value in found.items():
            out.write(f"# {key}: {value}\n")
        out.write(answer + "\n")
    write()


def measure(answer: str) -> list[dict[str, str]]:
    """The records for Excel's answers: one line per case, ``case<TAB>opened|refused<TAB>...``.

    A record of a refusal holds its error number under ``A1`` and nothing else.

    Raises:
        ValueError: the answers are not one per case in order, a line has the wrong number of
            fields, or a control did not open (then no outcome counts).
    """
    lines = [line for line in answer.strip().split("\n") if line and not line.startswith("# ")]
    if len(lines) != len(CASES):
        raise ValueError(f"expected {len(CASES)} cases, got {len(lines)}")
    records = []
    for (name, _), line in zip(CASES, lines, strict=True):
        got, outcome, *fields = line.split("\t")
        if got != name:
            raise ValueError(f"Excel answered for {got}, expected {name}")
        if outcome == "opened" and len(fields) == len(OPENED):
            values = fields
        elif outcome == "refused" and len(fields) == 1:
            values = [fields[0], *[""] * (len(OPENED) - 1)]
        else:
            raise ValueError(f"{name}: not an answer: {line!r}")
        if name.startswith("control") and outcome != "opened":
            raise ValueError(f"{name} did not open: no outcome counts")
        records.append({"case": name, "outcome": outcome, **dict(zip(OPENED, values, strict=True))})
    return records


def write(saved: Path = SAVED, output: Path = OUTPUT) -> list[dict[str, str]]:
    """Measure the saved answers again (no Excel) and write ``package_more.tsv``."""
    text = saved.read_text(encoding="utf-8")
    first = [line for line in text.splitlines() if line.startswith("# ")]
    records = measure(text)
    with output.open("w", encoding="utf-8", newline="\n") as out:
        for line in first:
            out.write(line + "\n")
        out.write("# suite: package_more\n")
        out.write("# one workbook a case, opened with alerts off; opened: A1 (111), the number\n")
        out.write("# of sheets, column width (characters) and width (points) of A and B, the\n")
        out.write("# heights of rows 1 and 2, A1's font size and name; refused: the error number\n")
        out.write("# under A1; method: AppleScript, no window\n")
        out.write("\t".join(COLUMNS) + "\n")
        for record in records:
            out.write("\t".join(record[c] for c in COLUMNS).rstrip("\t") + "\n")
    print(f"{len(records)} rows written to {output}")
    return records


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        paths = build(Path(argv[1]))
        print(f"{len(paths)} workbooks in {argv[1]}")
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
