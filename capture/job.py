"""Capture suites as Excel capture jobs, and the suite files their results give.

    uv run python -m capture.job write GROUP JOB_ID OUT_DIR
    uv run python -m capture.job read GROUP JOB_DIR RESULT_DIR

The job format is sheetmill's, version 1:
https://github.com/sheetmill/sheetmill/blob/e3a5207c251929a8a19d713777b70ace3506fb95/docs/capture-jobs.md
A group (``GROUPS``) is one workbook and the suites read from it. ``write`` makes the job folder:
the manifest (``job.json``) and the workbook in ``files/``, built as the suite's module builds it
for Excel for Mac, with its zip entries' times and document dates fixed, so that writing it again
gives the same bytes and the SHA-256 the result records. ``read`` checks the manifest's and the
input's SHA-256 against the result's and writes each suite's file in
``tests/captures/windows/`` in the schema of the Mac's, under a header naming the platform, the
Excel build and the job, with the result's ``header`` whole as one line of JSON
(``job_header``); the manifest is copied beside them, byte for byte, as ``<group>.job.json``.

Groups:

- ``cells``: ``capture/cell_text.py``'s workbook (suites ``merge_width`` and ``error_width``),
  opened, and the text each cell shows at its width read (``read`` of ``text``, one cell at a
  time), as the Mac's ``string value`` reads it.
- ``borders_fills``: the windows of the screen suites ``borders`` and ``fills``, a workbook each,
  built as the Mac's run builds them (the suite's ``fill`` and the two magenta markers). Each is
  opened, its cells' geometry and displayed fill read with the selection parked off the picture,
  and Excel's window captured at each of the suite's zooms (``view``, then ``capture_window``).
  ``read`` measures each picture with the suite's own ``measure``, as the Mac's ``read`` does, on
  a grid from the two markers at the picture's device pixels a point. The gridline's gray is
  measured in each picture (``capture.screen.run.gridline_gray``, on the control cell) and
  written in the file's header, so that what needs it uses the gray the picture shows.
- ``text_style``: the suite ``text_style``, a workbook a group (``text_style_<group>.xlsx``),
  captured at the group's zooms as ``text_steps`` is, with the displayed font and alignment read
  for every cell; the gray is measured on D1.
- ``shrink_fonts``: ``text_steps``' group of that name alone, as ``text_steps`` captures its
  groups, read into ``tests/captures/windows/shrink_fonts.tsv`` (``text_steps``' columns).
- ``font_cut``: the suite ``font_cut``, a workbook a group (``font_cut_<group>.xlsx``, the
  theme groups' themes rewritten before the copy), at zoom 100, with the displayed font and its
  color read for every cell; Excel's standard font is in the result's header.
- ``line_budget``: ``text_style``' group of that name alone, as ``text_style`` captures its
  groups, read into ``tests/captures/windows/line_budget.tsv`` (``text_style``' columns).
- ``text_steps``: the suite ``text_steps``, a workbook a group (``text_steps_<group>.xlsx``, each
  under the group's Normal font), captured at zoom 100 as ``borders_fills``' windows are and
  measured with the suite's own code into one file; the gray is measured on D1 in each picture.
"""

from __future__ import annotations

import hashlib
import json
import re
import shutil
import sys
import tempfile
import zipfile
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from capture import borders, cell_text, fills, font_cut, text_steps, text_style

ROOT = Path(__file__).resolve().parents[1]
WINDOWS = ROOT / "tests" / "captures" / "windows"
FIXED_TIME = "2026-01-01T00:00:00Z"
ZIP_TIME = (1980, 1, 1, 0, 0, 0)
GROUPS = {
    "cells": ("merge_width", "error_width"),
    "borders_fills": ("borders", "fills"),
    "text_steps": ("text_steps",),
    "text_style": ("text_style",),
    "shrink_fonts": ("shrink_fonts",),
    "line_budget": ("line_budget",),
    "font_cut": ("font_cut",),
}
PURPOSES = {
    "cells": "sheetpeek capture group cells: the text Excel shows in merged cells and for "
    "error values at set column widths (suites merge_width and error_width)",
    "borders_fills": "sheetpeek capture group borders_fills: window pictures of each border "
    "line style and each pattern fill at zoom 100 and 200 (suites borders and fills)",
    "text_steps": "sheetpeek capture group text_steps: window pictures of indents under seven "
    "Normal fonts and of shrink to fit by single characters, at zoom 100 (suite text_steps)",
    "shrink_fonts": "sheetpeek capture group shrink_fonts: window pictures of shrink to fit "
    "under cell fonts the group text_steps does not hold, at zoom 100 (text_steps' group "
    "shrink_fonts)",
    "font_cut": "sheetpeek capture group font_cut: window pictures of cells whose fonts lie "
    "past the workbook's font table's cut, under other themes and Normal fonts, at zoom 100 "
    "(suite font_cut)",
    "line_budget": "sheetpeek capture group line_budget: window pictures of two-word texts as "
    "wide as the cell less a few pixels, wrapped, justified and distributed, at zoom 100 "
    "(text_style's group line_budget)",
    "text_style": "sheetpeek capture group text_style: window pictures of alignment, shrink to "
    "fit, underlines and font effects, wrapping and distributed text, at zoom 100 and 200 "
    "(suite text_style)",
}
PARK = "AZ300"
"""The active cell while a window is captured: off every picture."""
SHEET = "Sheet"


@dataclass(frozen=True)
class Window:
    """One screen suite's window: how its workbook is filled, its cells and how it is measured."""

    suite: str
    fill: Callable[[Any], object]
    cells: tuple[str, ...]
    markers: tuple[str, str]
    zooms: tuple[int, ...]
    measure: Callable[[Any], list[dict[str, str]]]
    columns: tuple[str, ...]
    control: str
    """An unstyled cell above an unstyled one, whose bottom gridline gives the gray."""
    name: str = ""
    """The workbook's and its pictures' stem, by default the suite's: a suite of several windows
    names each."""
    display: tuple[str, ...] = ("interior.color", "interior.pattern")
    """The properties of the format Excel displays read for every cell (``DISPLAYED``)."""
    postprocess: Callable[[Path], None] | None = None
    """An edit of the saved workbook before it is copied into the job (``capture.screen.run``'s
    ``postprocess``)."""

    @property
    def stem(self) -> str:
        return self.name or self.suite


def _text_style_window(group: str, suite: str = "text_style") -> Window:
    return Window(
        suite,
        lambda ws: text_style.fill(ws, group),
        tuple(text_style._cells(text_style.cases(group))),
        ("B2", "J20"),
        text_style.ZOOMS[group],
        lambda cap: text_style.group_records(cap, group),
        text_style.COLUMNS,
        "D1",
        f"text_style_{group}",
        ("font.name", "font.size", "font.underline", "align.horizontal", "align.vertical"),
    )


def _font_cut_window(group: str) -> Window:
    return Window(
        "font_cut",
        lambda ws: font_cut.fill(ws, group),
        tuple(font_cut.references(group)),
        font_cut.MARKERS,
        font_cut.ZOOMS,
        lambda cap: font_cut.group_records(cap, group),
        font_cut.FIELDS,
        "D1",
        f"font_cut_{group}",
        ("font.name", "font.size", "font.color"),
        font_cut.postprocess(group),
    )


def _text_steps_window(group: str, suite: str = "text_steps") -> Window:
    return Window(
        suite,
        lambda ws: text_steps.fill(ws, group),
        tuple(text_steps.references(group)),
        text_steps.MARKERS[group],
        (100,),
        lambda cap: text_steps.group_records(cap, group),
        text_steps.COLUMNS,
        "D1",
        f"text_steps_{group}",
    )


WINDOWS_OF = {
    "borders_fills": (
        Window(
            "borders",
            borders.fill,
            tuple(borders.CELLS.values()),
            borders.MARKERS,
            borders.ZOOMS,
            borders.measure,
            borders.COLUMNS,
            borders.CELLS["none"],
        ),
        Window(
            "fills",
            fills.fill,
            tuple(cell for _, _, cell in fills.cases()),
            fills.MARKERS,
            fills.ZOOMS,
            fills.measure,
            fills.COLUMNS,
            next(c for p, s, c in fills.cases() if p == "none" and s == "bw"),
        ),
    ),
    "text_steps": tuple(
        _text_steps_window(group) for group in text_steps.GROUPS if group != "shrink_fonts"
    ),
    "shrink_fonts": (_text_steps_window("shrink_fonts", "shrink_fonts"),),
    "text_style": tuple(
        _text_style_window(group) for group in text_style.GROUPS if group != "line_budget"
    ),
    "line_budget": (_text_style_window("line_budget", "line_budget"),),
    "font_cut": tuple(_font_cut_window(group) for group in font_cut.GROUPS),
}
DISPLAYED = {
    "interior.color": "display.fill",
    "interior.pattern": "display.pattern",
    "font.name": "display.font_name",
    "font.color": "display.font_color",
    "font.size": "display.font_size",
    "font.underline": "display.underline",
    "align.horizontal": "display.h_align",
    "align.vertical": "display.v_align",
}
"""Each displayed property's key in the Mac's reads."""


def fixed_copy(raw: Path, path: Path) -> None:
    """``raw`` rewritten to ``path`` with fixed zip times and document dates."""
    with zipfile.ZipFile(raw) as source, zipfile.ZipFile(path, "w") as out:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "docProps/core.xml":
                text = re.sub(
                    r"(<dcterms:(?:created|modified)[^>]*>)[^<]*(<)",
                    rf"\g<1>{FIXED_TIME}\2",
                    data.decode("utf-8"),
                )
                data = text.encode("utf-8")
            entry = zipfile.ZipInfo(info.filename, date_time=ZIP_TIME)
            entry.compress_type = zipfile.ZIP_DEFLATED
            entry.external_attr = 0o600 << 16
            out.writestr(entry, data)


def _cells_workbook(files: Path) -> dict[str, Any]:
    with tempfile.TemporaryDirectory() as tmp:
        raw = cell_text.build(Path(tmp) / "raw.xlsx")
        fixed_copy(raw, files / "cell_text.xlsx")
    steps: list[dict[str, Any]] = [{"op": "open"}]
    for k, spec in enumerate(cell_text.specs()):
        steps.append({"op": "read", "id": f"text_{k}", "range": spec, "props": ["text"]})
    return {"file": "cell_text.xlsx", "steps": steps}


def _bounds(cells: list[str]) -> str:
    """``A1:<last column><last row>`` over ``cells``."""
    from openpyxl.utils.cell import (
        column_index_from_string,
        coordinate_from_string,
        get_column_letter,
    )

    found = [coordinate_from_string(c) for c in cells]
    last_column = max(column_index_from_string(letter) for letter, _ in found)
    return f"A1:{get_column_letter(last_column)}{max(row for _, row in found)}"


def _window_workbook(window: Window, files: Path) -> dict[str, Any]:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill

    from capture.screen.window import MARKER

    name = f"{window.stem}.xlsx"
    with tempfile.TemporaryDirectory() as tmp:
        raw = Path(tmp) / name
        wb = Workbook()
        window.fill(wb.active)
        for cell in window.markers:
            wb.active[cell].fill = PatternFill("solid", fgColor="FF" + MARKER)
        wb.save(raw)
        if window.postprocess is not None:
            window.postprocess(raw)
        fixed_copy(raw, files / name)
    span = f"{SHEET}!{_bounds([*window.markers, *window.cells])}"
    park = {"op": "view", "sheet": SHEET, "select": f"{SHEET}!{PARK}"}
    steps: list[dict[str, Any]] = [
        {"op": "open"},
        {**park, "zoom": window.zooms[0], "scroll_row": 1, "scroll_column": 1},
        {"op": "read", "id": "geometry", "range": span,
         "props": ["left", "top", "width", "height", "text"]},
        {"op": "read_format", "id": "display", "range": span,
         "props": list(window.display), "display": True},
    ]  # fmt: skip
    for zoom in window.zooms:
        steps.append({**park, "id": f"view_z{zoom}", "zoom": zoom, "scroll_row": 1,
                      "scroll_column": 1})  # fmt: skip
        steps.append({"op": "capture_window", "id": f"shot_z{zoom}",
                      "name": f"{window.stem}_z{zoom}.png"})  # fmt: skip
    steps.append({"op": "close", "save": False})
    return {"file": name, "steps": steps}


def manifest(group: str, job_id: str, workbooks: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "schema": 1,
        "job_id": job_id,
        "package": "sheetpeek",
        "purpose": PURPOSES[group],
        "limits": {"job_seconds": 1800, "step_seconds": 600},
        "audit": False,
        "workbooks": workbooks,
    }


def write(group: str, job_id: str, out: Path) -> Path:
    """The job folder ``out/<job_id>``."""
    if group not in GROUPS:
        sys.exit(f"unknown group {group!r}; expected one of {', '.join(GROUPS)}")
    folder = out / job_id
    files = folder / "files"
    files.mkdir(parents=True)
    if group in WINDOWS_OF:
        books = [_window_workbook(w, files) for w in WINDOWS_OF[group]]
    else:
        books = [_cells_workbook(files)]
    doc = manifest(group, job_id, books)
    (folder / "job.json").write_text(json.dumps(doc, indent=1) + "\n", encoding="utf-8")
    return folder


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _results(workbook: dict[str, Any]) -> dict[str, Any]:
    if workbook.get("status") != "ok":
        raise ValueError(f"{workbook.get('file')}: workbook status {workbook.get('status')!r}")
    found = {}
    for step in workbook["steps"]:
        if step.get("status") != "ok":
            raise ValueError(f"{workbook.get('file')}: step {step.get('op')}: {step.get('error')}")
        if "id" in step:
            found[step["id"]] = step.get("result")
    return found


def _shown(prop: str, value: Any) -> str:
    """A displayed property as the Mac's reads keep it: a color as ``r,g,b`` (Windows returns one
    number, red in the low byte), a size as a number, an enumeration as Windows' number; a
    property the cell holds several of (a rich-text cell's fonts), which Windows returns as
    nothing, as an empty field."""
    if value is None:
        return ""
    if prop in ("interior.color", "font.color"):
        v = int(value)
        return f"{v & 0xFF},{(v >> 8) & 0xFF},{(v >> 16) & 0xFF}"
    if prop == "font.name":
        return str(value)
    if prop == "font.size":
        return f"{float(value):g}"
    return str(int(value))


def _cell_rows(steps: dict[str, Any]) -> list[tuple[str, str, str]]:
    """The geometry and display reads as the Mac's run keeps them: ``(cell, property, value)``.

    Displayed properties are keyed as the Mac's (``DISPLAYED``); ``display.pattern``,
    ``display.underline``, ``display.h_align`` and ``display.v_align`` are Windows' numbers
    (``XlPattern``, ``XlUnderlineStyle``, ``XlHAlign``, ``XlVAlign``) as returned.
    """
    from openpyxl.utils.cell import get_column_letter

    geometry, display = steps["geometry"], steps["display"]
    rows: list[tuple[str, str, str]] = []
    for r in range(int(geometry["rows"])):
        for c in range(int(geometry["columns"])):
            name = f"{get_column_letter(c + 1)}{r + 1}"
            rows.extend(
                (name, f"geometry.{key}", repr(float(geometry[key][r][c])))
                for key in ("left", "top", "width", "height")
            )
            # as the Mac's run writes a read (write_tsv): a tab or line break as a space
            text = str(geometry["text"][r][c] or "")
            rows.append((name, "geometry.text", re.sub(r"[\t\r\n]", " ", text)))
            cell = display["cells"][r][c]
            rows.extend(
                (name, DISPLAYED[prop], _shown(prop, value))
                for prop, value in cell.items()
                if prop in DISPLAYED
            )
    return rows


def window_captures(window: Window, steps: dict[str, Any], artifacts: Path) -> list[Any]:
    """The window's ``Capture`` at each zoom, its gridline gray measured in the picture."""
    import dataclasses

    from capture.screen import png
    from capture.screen.run import Capture, _geometry, gridline_gray
    from capture.screen.window import locate, nominal_scale

    rows = _cell_rows(steps)
    geometry = _geometry(rows)
    found = []
    for zoom in window.zooms:
        seen, shot = steps[f"view_z{zoom}"], steps[f"shot_z{zoom}"]
        if float(seen["zoom"]) != zoom:
            raise ValueError(f"{window.suite}: zoom {seen['zoom']} where {zoom} was asked")
        display = float(shot["dpi"]) / 96
        image = png.read(artifacts / shot["artifact"])
        nominal = nominal_scale(zoom, display, "windows")
        grid, check = locate(
            image, geometry[window.markers[0]], geometry[window.markers[1]], nominal
        )
        # two markers: their scale must be the nominal one; one (the second below the window at
        # zoom 200): the grid is the first marker's at the nominal scale (CAPTURING.md)
        if check.get("scale") != "nominal" and float(str(check.get("nominal_off", 1))) > 0.01:
            raise ValueError(f"{window.suite} zoom {zoom}: the markers' scale is not nominal")
        cap = Capture(zoom, image, grid, check, geometry, rows, "windows")
        gray = gridline_gray(cap, window.control)
        if gray is None:
            raise ValueError(f"{window.suite} zoom {zoom}: no gridline gray at {window.control}")
        found.append(dataclasses.replace(cap, measured_gridline=gray))
    return found


def _job_lines(
    group: str, outcome: dict[str, Any], manifest_doc: dict[str, Any], method: str
) -> tuple[str, tuple[str, ...]]:
    header = outcome["header"]
    build = str(header.get("environment", {}).get("excel", {}).get("build", ""))
    lines = (
        "# platform: windows",
        f"# excel_build: {build}",
        f"# job: {group}.job.json",
        f"# job_id: {manifest_doc['job_id']}",
        f"# captured: {str(outcome.get('started', ''))[:10]}; method: {method}",
        f"# job_header: {json.dumps(header, sort_keys=True, ensure_ascii=False)}",
    )
    return build, lines


def _read_windows(
    group: str,
    outcome: dict[str, Any],
    manifest_doc: dict[str, Any],
    result: Path,
    folder: Path,
) -> list[Path]:
    by_file = {wb["file"]: wb for wb in outcome["workbooks"]}
    suites: dict[str, list[Window]] = {}
    for window in WINDOWS_OF[group]:
        suites.setdefault(window.suite, []).append(window)
    written = []
    for suite, windows in suites.items():
        records: list[dict[str, str]] = []
        grays, zooms = set(), {}
        for window in windows:
            steps = _results(by_file[f"{window.stem}.xlsx"])
            for cap in window_captures(window, steps, result / "artifacts"):
                grays.add("".join(f"{c:02X}" for c in cap.gridline))
                records += window.measure(cap)
            zooms.update(dict.fromkeys(window.zooms))
        if len(grays) != 1:
            raise ValueError(f"{suite}: the pictures show different gridline grays {grays}")
        _, lines = _job_lines(group, outcome, manifest_doc, "capture_window (capture job)")
        controls = ", ".join(dict.fromkeys(w.control for w in windows))
        columns = windows[0].columns
        path = folder / f"{suite}.tsv"
        with path.open("w", encoding="utf-8", newline="\n") as out:
            out.write("\n".join(lines) + "\n")
            out.write(f"# zooms: {','.join(str(z) for z in zooms)}\n")
            out.write(f"# gridline: {grays.pop()} (measured: {controls}'s bottom edge)\n")
            out.write(f"# suite: {suite}\n" + "\t".join(columns) + "\n")
            for record in records:
                out.write("\t".join(record.get(c, "") for c in columns) + "\n")
        written.append(path)
    return written


def read(group: str, job: Path, result: Path, folder: Path = WINDOWS) -> list[Path]:
    """The suite files from a job folder and its result folder (``result.json``)."""
    outcome = json.loads((result / "result.json").read_text(encoding="utf-8"))
    if outcome.get("status") != "ok":
        raise ValueError(f"the job did not end ok: {outcome.get('status')!r}")
    header = outcome["header"]
    if header["manifest_sha256"] != _sha256(job / "job.json"):
        raise ValueError("the result's manifest_sha256 is not this job's manifest")
    for name, entry in header["inputs"].items():
        if entry["sha256"] != _sha256(job / "files" / name):
            raise ValueError(f"the result's input {name} is not this job's")
    manifest_doc = json.loads((job / "job.json").read_text(encoding="utf-8"))
    folder.mkdir(parents=True, exist_ok=True)
    if group in WINDOWS_OF:
        written = _read_windows(group, outcome, manifest_doc, result, folder)
        shutil.copyfile(job / "job.json", folder / f"{group}.job.json")
        return written
    by_file = {wb["file"]: wb for wb in outcome["workbooks"]}
    results = _results(by_file["cell_text.xlsx"])
    shown: dict[tuple[str, int, int], str] = {}
    for k, spec in enumerate(cell_text.specs()):
        for r, row in enumerate(results[f"text_{k}"]["text"], 1):
            for c, text in enumerate(row, 1):
                shown[(spec, r, c)] = str(text)
    build, lines = _job_lines(group, outcome, manifest_doc, "read text (capture job)")
    method = "\n".join(lines)
    written = cell_text.write_tsvs(shown, f"Excel {build} for Windows", method, folder)
    shutil.copyfile(job / "job.json", folder / f"{group}.job.json")
    return written


def main(argv: list[str]) -> int:
    if len(argv) == 4 and argv[0] == "write":
        print(write(argv[1], argv[2], Path(argv[3])))
        return 0
    if len(argv) == 4 and argv[0] == "read":
        for path in read(argv[1], Path(argv[2]), Path(argv[3])):
            print(path)
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
