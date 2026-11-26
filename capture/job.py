"""Capture suites as Excel capture jobs, and the suite files their results give.

    uv run python -m capture.job write GROUP JOB_ID OUT_DIR
    uv run python -m capture.job read GROUP JOB_DIR RESULT_DIR

The job format is sheetmill's, version 1:
https://github.com/sheetmill/sheetmill/blob/3adb865500f4d2223e5c1a758b466e081d3fdc79/docs/capture-jobs.md
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
- ``shrink_fonts``: ``text_steps``'s group of that name alone, as ``text_steps`` captures its
  groups, read into ``tests/captures/windows/shrink_fonts.tsv`` (``text_steps``'s columns).
- ``font_cut``: the suite ``font_cut``, a workbook a group (``font_cut_<group>.xlsx``, the
  theme groups' themes rewritten before the copy), at zoom 100, with the displayed font and its
  color read for every cell; Excel's standard font is in the result's header.
- ``font_faces``: ``font_cut``'s ``faces_*`` groups, as ``font_cut`` writes its groups, read
  into ``tests/captures/windows/font_faces.tsv``.
- ``font_aptos``: ``font_cut``'s group ``faces_aptos`` alone, read into
  ``tests/captures/windows/font_aptos.tsv``.
- ``merge_rows``: the suite ``merge_rows``'s workbook opened, and every case row's height and
  column A's font read from Excel (``read`` of ``row_height``, ``read_format`` of ``font.name``
  and ``font.size``), into ``tests/captures/windows/merge_rows.tsv`` (the suite's columns).
- ``line_budget``: ``text_style``'s group of that name alone, as ``text_style`` captures its
  groups, read into ``tests/captures/windows/line_budget.tsv`` (``text_style``'s columns).
- ``fitted_rows``: ``row_height_fitted``'s groups ``styled`` and ``scripts`` and the suite
  ``wrap_rows``, a workbook each, opened, and every case row's height and column A's font read
  (as ``merge_rows``), into ``tests/captures/windows/row_height_fitted.tsv`` (those groups'
  records, the Mac suite's columns) and ``tests/captures/windows/wrap_rows.tsv``.
- ``wrap_band``: the suite ``wrap_band``, a workbook a column width, opened, and every case
  row's height, column A's width (``column_width`` and ``width``) and column A's font read, into
  ``tests/captures/windows/wrap_band.tsv`` (the suite's columns).
- ``row_labels``: ``row_bottoms``' group ``labels``, as the Mac's run captures it, read into
  ``tests/captures/windows/row_labels.tsv`` (``row_bottoms``' columns).
- ``refit_rows``: the suite ``refit_rows``, as ``stored_rows``, into
  ``tests/captures/windows/refit_rows.tsv``.
- ``stored_rows``: the suite ``stored_rows``, a workbook each, every case row's ``row_height``,
  ``height`` and ``top`` read, into ``tests/captures/windows/stored_rows.tsv``.
- ``wrap_steps``: the suite ``wrap_steps`` as ``wrap_band``, into
  ``tests/captures/windows/wrap_steps.tsv``.
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

from capture import (
    borders,
    cell_text,
    fills,
    font_cut,
    merge_rows,
    refit_rows,
    row_bottoms,
    row_height_fitted,
    stored_rows,
    text_steps,
    text_style,
    wrap_band,
    wrap_rows,
    wrap_steps,
)

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
    "row_labels": ("row_labels",),
    "stored_rows": ("stored_rows",),
    "refit_rows": ("refit_rows",),
    "font_cut": ("font_cut",),
    "font_faces": ("font_faces",),
    "font_aptos": ("font_aptos",),
    "merge_rows": ("merge_rows",),
    "fitted_rows": ("row_height_fitted", "wrap_rows"),
    "wrap_band": ("wrap_band",),
    "wrap_steps": ("wrap_steps",),
}
FITTED_GROUPS = ("styled", "scripts")
"""``row_height_fitted``'s groups in the job ``fitted_rows``."""
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
    "font_faces": "sheetpeek capture group font_faces: window pictures of cells past the font "
    "table's cut beside reference cells in five faces, under two Normal fonts, at zoom 100 "
    "(font_cut's faces groups)",
    "font_aptos": "sheetpeek capture group font_aptos: window pictures of cells past the font "
    "table's cut beside reference cells, under an Aptos Narrow theme and Normal font, at zoom "
    "100 (font_cut's group faces_aptos)",
    "merge_rows": "sheetpeek capture group merge_rows: the heights Excel fits to rows around "
    "merges across rows, and the font it applies in each merge, read from Excel.",
    "fitted_rows": "sheetpeek capture group fitted_rows: the heights Excel fits to rows of "
    "superscript, subscript and accounting-underlined text and of wrapped, justified and "
    "distributed text of one to six lines, and the font it applies, read from Excel.",
    "wrap_band": "sheetpeek capture group wrap_band: the heights Excel fits to rows of "
    "wrapped two-word texts whose width steps by one unit across the column's last pixels or "
    "points, at several column widths, with each column's width and the applied font, read "
    "from Excel.",
    "wrap_steps": "sheetpeek capture group wrap_steps: the heights Excel fits to rows of "
    "wrapped two-word texts across a column's last pixels or points, at columns placing "
    "autofit's later long-text steps and under a Courier 10 Normal font, with each column's "
    "width and the applied font, read from Excel.",
    "refit_rows": "sheetpeek capture group refit_rows: the heights Excel lays out for rows "
    "storing a height without customHeight, below, at and above the height their cell fits, "
    'for a sheet as Excel for Mac saves it, for heights of 0 written as ht="0" and ht="", '
    "and in sheets with a custom default height or zeroHeight, read from Excel.",
    "stored_rows": "sheetpeek capture group stored_rows: the heights and tops Excel lays out for "
    "rows whose height a file stores in points, across the fractions of a pixel, as a row's ht "
    "with customHeight and without and as the sheet's defaultRowHeight, read from Excel.",
    "row_labels": "sheetpeek capture group row_labels: a window picture at zoom 100 of text "
    "beside a merge across columns and beside empty cells styled angled, in rows of a fixed "
    "height, and of turned text, merged or not, in fitted rows (row_bottoms' group labels)",
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


def _font_cut_window(group: str, suite: str = "font_cut") -> Window:
    return Window(
        suite,
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
    "row_labels": (
        Window(
            "row_labels",
            lambda ws: row_bottoms.fill(ws, "labels"),
            tuple(row_bottoms.references("labels")),
            row_bottoms.MARKERS["labels"],
            (100,),
            lambda cap: row_bottoms.group_records(cap, "labels"),
            row_bottoms.COLUMNS,
            "D1",
            "row_bottoms_labels",
        ),
    ),
    "font_cut": tuple(
        _font_cut_window(group) for group in font_cut.GROUPS if group not in font_cut.FACES
    ),
    "font_faces": tuple(
        _font_cut_window(group, "font_faces") for group in ("faces_calibri", "faces_times")
    ),
    "font_aptos": (_font_cut_window("faces_aptos", "font_aptos"),),
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


def _merge_rows_workbook(files: Path) -> dict[str, Any]:
    with tempfile.TemporaryDirectory() as tmp:
        raw = merge_rows.build(Path(tmp))
        fixed_copy(raw, files / merge_rows.WORKBOOK)
    last = max(case.row for case in merge_rows.CASES)
    return {
        "file": merge_rows.WORKBOOK,
        "steps": [
            {"op": "open"},
            {"op": "read", "id": "heights", "range": f"A1:A{last}", "props": ["row_height"]},
            {"op": "read_format", "id": "fonts", "range": f"A1:A{last}",
             "props": ["font.name", "font.size"]},
        ],
    }  # fmt: skip


def _rows_workbook(name: str, last: int) -> dict[str, Any]:
    """Open ``name`` and read rows 1 to ``last``: each row's height and column A's font."""
    return {
        "file": name,
        "steps": [
            {"op": "open"},
            {"op": "read", "id": "heights", "range": f"A1:A{last}", "props": ["row_height"]},
            {"op": "read_format", "id": "fonts", "range": f"A1:A{last}",
             "props": ["font.name", "font.size"]},
        ],
    }  # fmt: skip


def _fitted_rows_workbooks(files: Path) -> list[dict[str, Any]]:
    books = []
    with tempfile.TemporaryDirectory() as tmp:
        for group in FITTED_GROUPS:
            raw = row_height_fitted.build_one(group, Path(tmp))
            fixed_copy(raw, files / raw.name)
            books.append(_rows_workbook(raw.name, len(row_height_fitted.rows(group))))
        raw = wrap_rows.build(Path(tmp))
        fixed_copy(raw, files / wrap_rows.WORKBOOK)
        books.append(_rows_workbook(wrap_rows.WORKBOOK, len(wrap_rows.CASES)))
    return books


def _wrap_band_workbooks(files: Path) -> list[dict[str, Any]]:
    books = []
    with tempfile.TemporaryDirectory() as tmp:
        for column in wrap_band.COLUMN_KEYS:
            raw = wrap_band.build_one(column, Path(tmp))
            fixed_copy(raw, files / raw.name)
            last = len(wrap_band.cases(column))
            books.append(
                {
                    "file": raw.name,
                    "steps": [
                        {"op": "open"},
                        {"op": "read", "id": "heights", "range": f"A1:A{last}",
                         "props": ["row_height", "column_width", "width"]},
                        {"op": "read_format", "id": "fonts", "range": f"A1:A{last}",
                         "props": ["font.name", "font.size"]},
                    ],
                }
            )  # fmt: skip
    return books


def _rows_widths_workbook(name: str, last: int) -> dict[str, Any]:
    """Open ``name`` and read rows 1 to ``last``: height, column A's width and font."""
    return {
        "file": name,
        "steps": [
            {"op": "open"},
            {"op": "read", "id": "heights", "range": f"A1:A{last}",
             "props": ["row_height", "column_width", "width"]},
            {"op": "read_format", "id": "fonts", "range": f"A1:A{last}",
             "props": ["font.name", "font.size"]},
        ],
    }  # fmt: skip


def _wrap_steps_workbooks(files: Path) -> list[dict[str, Any]]:
    books = []
    with tempfile.TemporaryDirectory() as tmp:
        for column in wrap_steps.COLUMN_LIST:
            raw = wrap_steps.build_one(column, Path(tmp))
            fixed_copy(raw, files / raw.name)
            books.append(_rows_widths_workbook(raw.name, len(wrap_steps.cases(column))))
    return books


def _width_answer(workbook: dict[str, Any], name: str, count: int) -> list[str]:
    """``name``'s rows 1 to ``count`` as the suites' AppleScript answers them, widths included."""
    results = _results(workbook)
    heights = results["heights"]
    fonts = results["fonts"]["cells"]
    return [
        "\t".join(
            (
                name,
                str(row),
                str(fonts[row - 1][0]["font.name"]),
                f"{float(fonts[row - 1][0]['font.size']):g}",
                f"{float(heights['row_height'][row - 1][0]):g}",
                repr(float(heights["column_width"][row - 1][0])),
                repr(float(heights["width"][row - 1][0])),
            )
        )
        for row in range(1, count + 1)
    ]


def _stored_rows_workbooks(files: Path) -> list[dict[str, Any]]:
    books = []
    with tempfile.TemporaryDirectory() as tmp:
        for form, stored in stored_rows.workbooks():
            raw = stored_rows.build_one(form, stored, Path(tmp))
            fixed_copy(raw, files / raw.name)
            last = len(stored_rows.cases(form, stored))
            books.append(
                {
                    "file": raw.name,
                    "steps": [
                        {"op": "open"},
                        {"op": "read", "id": "rows", "range": f"A1:A{last}",
                         "props": ["row_height", "height", "top"]},
                    ],
                }
            )  # fmt: skip
    return books


def _refit_rows_workbooks(files: Path) -> list[dict[str, Any]]:
    books = []
    with tempfile.TemporaryDirectory() as tmp:
        for name, cases in refit_rows.CASES.items():
            raw = refit_rows.build_one(name, Path(tmp))
            fixed_copy(raw, files / raw.name)
            books.append(
                {
                    "file": raw.name,
                    "steps": [
                        {"op": "open"},
                        {"op": "read", "id": "rows", "range": f"A1:A{len(cases)}",
                         "props": ["row_height", "height", "top"]},
                    ],
                }
            )  # fmt: skip
    return books


def _read_refit_rows(
    outcome: dict[str, Any], manifest_doc: dict[str, Any], folder: Path
) -> list[Path]:
    by_file = {wb["file"]: wb for wb in outcome["workbooks"]}
    _, lines = _job_lines("refit_rows", outcome, manifest_doc, "read row_height (capture job)")
    found = []
    for name, cases in refit_rows.CASES.items():
        rows = _results(by_file[refit_rows.workbook_name(name)])["rows"]
        for k, case in enumerate(cases):
            values = [repr(float(rows[key][k][0])) for key in ("row_height", "height", "top")]
            found.append(refit_rows.record(case, *values))
    path = folder / "refit_rows.tsv"
    refit_rows.write_records([*lines, "# suite: refit_rows"], found, path)
    return [path]


def _read_stored_rows(
    outcome: dict[str, Any], manifest_doc: dict[str, Any], folder: Path
) -> list[Path]:
    """``stored_rows.tsv``: each case row's reported height, height and top, in points."""
    by_file = {wb["file"]: wb for wb in outcome["workbooks"]}
    _, lines = _job_lines("stored_rows", outcome, manifest_doc, "read row_height (capture job)")
    path = folder / "stored_rows.tsv"
    with path.open("w", encoding="utf-8", newline="\n") as out:
        out.write("\n".join(lines) + "\n# suite: stored_rows\n")
        out.write("\t".join(stored_rows.COLUMNS) + "\n")
        for form, stored in stored_rows.workbooks():
            found = _results(by_file[stored_rows.workbook_name(form, stored)])["rows"]
            for k, case in enumerate(stored_rows.cases(form, stored)):
                values = [repr(float(found[key][k][0])) for key in ("row_height", "height", "top")]
                fields = [form, case.name, f"{case.stored:g}", str(case.row), *values]
                out.write("\t".join(fields) + "\n")
    return [path]


def _read_wrap_steps(
    outcome: dict[str, Any], manifest_doc: dict[str, Any], folder: Path
) -> list[Path]:
    """``wrap_steps.tsv`` in the Mac suite's columns."""
    by_file = {wb["file"]: wb for wb in outcome["workbooks"]}
    answer = []
    for column in wrap_steps.COLUMN_LIST:
        name = column.workbook
        answer += _width_answer(by_file[name], name, len(wrap_steps.cases(column)))
    _, found = wrap_steps.records("\n".join(answer))
    _, lines = _job_lines("wrap_steps", outcome, manifest_doc, "read row_height (capture job)")
    path = folder / "wrap_steps.tsv"
    wrap_steps.write_records([*lines, "# suite: wrap_steps"], found, path)
    return [path]


def _read_wrap_band(
    outcome: dict[str, Any], manifest_doc: dict[str, Any], folder: Path
) -> list[Path]:
    """``wrap_band.tsv`` in the Mac suite's columns, from the answer lines its AppleScript gives."""
    by_file = {wb["file"]: wb for wb in outcome["workbooks"]}
    answer = []
    for column in wrap_band.COLUMN_KEYS:
        name = wrap_band.workbook_name(column)
        results = _results(by_file[name])
        heights = results["heights"]
        fonts = results["fonts"]["cells"]
        for row in range(1, len(wrap_band.cases(column)) + 1):
            font = fonts[row - 1][0]
            answer.append(
                "\t".join(
                    (
                        name,
                        str(row),
                        str(font["font.name"]),
                        f"{float(font['font.size']):g}",
                        f"{float(heights['row_height'][row - 1][0]):g}",
                        repr(float(heights["column_width"][row - 1][0])),
                        repr(float(heights["width"][row - 1][0])),
                    )
                )
            )
    _, found = wrap_band.records("\n".join(answer))
    _, lines = _job_lines("wrap_band", outcome, manifest_doc, "read row_height (capture job)")
    path = folder / "wrap_band.tsv"
    wrap_band.write_records([*lines, "# suite: wrap_band"], found, path)
    return [path]


def _answer_lines(workbook: dict[str, Any], name: str, count: int) -> list[str]:
    """``name``'s rows 1 to ``count`` as the Mac's AppleScript answers them."""
    results = _results(workbook)
    heights = results["heights"]["row_height"]
    fonts = results["fonts"]["cells"]
    return [
        "\t".join(
            (
                name,
                str(row),
                str(fonts[row - 1][0]["font.name"]),
                f"{float(fonts[row - 1][0]['font.size']):g}",
                f"{float(heights[row - 1][0]):g}",
            )
        )
        for row in range(1, count + 1)
    ]


def _read_fitted_rows(
    outcome: dict[str, Any], manifest_doc: dict[str, Any], folder: Path
) -> list[Path]:
    """``row_height_fitted.tsv`` (groups ``styled`` and ``scripts``) and ``wrap_rows.tsv``."""
    by_file = {wb["file"]: wb for wb in outcome["workbooks"]}
    _, lines = _job_lines("fitted_rows", outcome, manifest_doc, "read row_height (capture job)")
    path = folder / "row_height_fitted.tsv"
    columns = row_height_fitted.COLUMNS
    with path.open("w", encoding="utf-8", newline="\n") as out:
        out.write("\n".join(lines) + "\n")
        out.write(f"# suite: row_height_fitted, groups {', '.join(FITTED_GROUPS)}\n")
        out.write("\t".join(columns) + "\n")
        for group in FITTED_GROUPS:
            name = row_height_fitted.workbook_name(group)
            cases = row_height_fitted.rows(group)
            answers = _answer_lines(by_file[name], name, len(cases))
            for (row, case), answer in zip(cases, answers, strict=True):
                _, _, applied_name, applied_size, height = answer.split("\t")
                record = {
                    "group": group,
                    "name": case.name,
                    "row": str(row),
                    "rotation": str(case.rotation),
                    "size": f"{case.size:g}",
                    "applied": f"{applied_name} {applied_size}",
                    "height": height,
                    "text": case.text,
                }
                out.write("\t".join(record[c] for c in columns) + "\n")
    answer = _answer_lines(by_file[wrap_rows.WORKBOOK], wrap_rows.WORKBOOK, len(wrap_rows.CASES))
    _, found = wrap_rows.records("\n".join(answer))
    wrapped = folder / "wrap_rows.tsv"
    with wrapped.open("w", encoding="utf-8", newline="\n") as out:
        out.write("\n".join(lines) + "\n")
        out.write("# suite: wrap_rows\n" + "\t".join(wrap_rows.COLUMNS) + "\n")
        for record in found:
            out.write("\t".join(record[c] for c in wrap_rows.COLUMNS) + "\n")
    return [path, wrapped]


def _read_merge_rows(
    outcome: dict[str, Any], manifest_doc: dict[str, Any], folder: Path
) -> list[Path]:
    """``merge_rows.tsv`` in the Mac suite's columns: each case row's height and applied font."""
    by_file = {wb["file"]: wb for wb in outcome["workbooks"]}
    results = _results(by_file[merge_rows.WORKBOOK])
    heights = results["heights"]["row_height"]
    fonts = results["fonts"]["cells"]
    answer = []
    for case in merge_rows.CASES:
        font = fonts[case.row - 1][0]
        answer.append(
            "\t".join(
                (
                    merge_rows.WORKBOOK,
                    str(case.row),
                    str(font["font.name"]),
                    f"{float(font['font.size']):g}",
                    f"{float(heights[case.row - 1][0]):g}",
                )
            )
        )
    _, found = merge_rows.records("\n".join(answer))
    _, lines = _job_lines("merge_rows", outcome, manifest_doc, "read row_height (capture job)")
    path = folder / "merge_rows.tsv"
    with path.open("w", encoding="utf-8", newline="\n") as out:
        out.write("\n".join(lines) + "\n")
        out.write("# suite: merge_rows\n" + "\t".join(merge_rows.COLUMNS) + "\n")
        for record in found:
            out.write("\t".join(record[c] for c in merge_rows.COLUMNS) + "\n")
    return [path]


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
    elif group == "merge_rows":
        books = [_merge_rows_workbook(files)]
    elif group == "fitted_rows":
        books = _fitted_rows_workbooks(files)
    elif group == "wrap_band":
        books = _wrap_band_workbooks(files)
    elif group == "wrap_steps":
        books = _wrap_steps_workbooks(files)
    elif group == "stored_rows":
        books = _stored_rows_workbooks(files)
    elif group == "refit_rows":
        books = _refit_rows_workbooks(files)
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
    if group in WINDOWS_OF or group in (
        "merge_rows",
        "fitted_rows",
        "wrap_band",
        "wrap_steps",
        "stored_rows",
        "refit_rows",
    ):
        if group == "merge_rows":
            written = _read_merge_rows(outcome, manifest_doc, folder)
        elif group == "fitted_rows":
            written = _read_fitted_rows(outcome, manifest_doc, folder)
        elif group == "wrap_band":
            written = _read_wrap_band(outcome, manifest_doc, folder)
        elif group == "wrap_steps":
            written = _read_wrap_steps(outcome, manifest_doc, folder)
        elif group == "stored_rows":
            written = _read_stored_rows(outcome, manifest_doc, folder)
        elif group == "refit_rows":
            written = _read_refit_rows(outcome, manifest_doc, folder)
        else:
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
