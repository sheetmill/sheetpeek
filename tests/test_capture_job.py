"""``capture/job.py``: the jobs it writes, and the suite files a result gives, without Excel.

The reader runs on a result made from the Mac's committed captures, so that its files must hold
the Mac's rows exactly: the grid a job returns and the text the Mac's AppleScript returns are
written by the same code.
"""

import hashlib
import json

import pytest

pytest.importorskip("openpyxl")

from capture import cell_text, job

CAPTURES = job.ROOT / "tests" / "captures"


def _sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_writing_a_job_again_gives_the_same_bytes(tmp_path):
    first = job.write("cells", "sheetpeek-test-cells", tmp_path / "a")
    second = job.write("cells", "sheetpeek-test-cells", tmp_path / "b")
    for name in ("job.json", "files/cell_text.xlsx"):
        assert _sha(first / name) == _sha(second / name), name
    doc = json.loads((first / "job.json").read_text(encoding="utf-8"))
    assert doc["package"] == "sheetpeek"
    assert doc["purpose"].startswith("sheetpeek capture group cells")
    (workbook,) = doc["workbooks"]
    assert workbook["steps"][0] == {"op": "open"}
    assert [s["range"] for s in workbook["steps"][1:]] == cell_text.specs()


def _mac_grids():
    """The text the Mac read, as the grids a job's read steps return, from the committed files."""
    rows = {}
    for suite in ("merge_width", "error_width"):
        lines = (CAPTURES / f"{suite}.tsv").read_text(encoding="utf-8").splitlines()
        rows[suite] = [line.split("\t") for line in lines if not line.startswith("# ")][1:]
    # The Mac's file goes width by width: each width's merges, then its hidden-column rows.
    by_sheet: dict[str, list[list[str]]] = {}
    for r in rows["merge_width"]:
        sheet = ("h" if r[1] == "A:C, B hidden" else "w") + r[0]
        by_sheet.setdefault(sheet, []).append([r[4]])
    grids = {}
    for k, spec in enumerate(cell_text.specs()):
        if spec.startswith("errors!"):
            width = len(cell_text.ERROR_WIDTHS)
            texts = [r[2] for r in rows["error_width"]]
            grids[f"text_{k}"] = [texts[i : i + width] for i in range(0, len(texts), width)]
        else:
            grids[f"text_{k}"] = by_sheet[spec.split("!")[0]]
    return grids


def _result(folder, job_folder, grids, *, manifest_sha=None):
    steps = [{"op": "open", "status": "ok"}]
    steps += [
        {"op": "read", "id": key, "status": "ok", "result": {"text": grid}}
        for key, grid in grids.items()
    ]
    header = {
        "manifest_sha256": manifest_sha or _sha(job_folder / "job.json"),
        "inputs": {"cell_text.xlsx": {"sha256": _sha(job_folder / "files" / "cell_text.xlsx")}},
        "environment": {"excel": {"build": "16.0.0.0"}},
    }
    outcome = {
        "status": "ok",
        "started": "2026-11-23T00:00:00Z",
        "header": header,
        "workbooks": [{"file": "cell_text.xlsx", "status": "ok", "steps": steps}],
    }
    folder.mkdir(parents=True)
    (folder / "result.json").write_text(json.dumps(outcome), encoding="utf-8")
    return folder


def test_a_result_reads_into_the_macs_schema_and_rows(tmp_path):
    job_folder = job.write("cells", "sheetpeek-test-cells", tmp_path / "jobs")
    result = _result(tmp_path / "result", job_folder, _mac_grids())
    out = tmp_path / "windows"
    written = job.read("cells", job_folder, result, out)
    assert [p.name for p in written] == ["merge_width.tsv", "error_width.tsv"]
    for path in written:
        ours = [x for x in path.read_text(encoding="utf-8").splitlines() if not x.startswith("# ")]
        theirs = (CAPTURES / path.name).read_text(encoding="utf-8").splitlines()
        assert ours == [x for x in theirs if not x.startswith("# ")], path.name
        assert "# platform: windows" in path.read_text(encoding="utf-8")
        assert "# excel_build: 16.0.0.0" in path.read_text(encoding="utf-8")
    assert (out / "cells.job.json").read_bytes() == (job_folder / "job.json").read_bytes()


def test_a_result_of_another_job_is_refused(tmp_path):
    job_folder = job.write("cells", "sheetpeek-test-cells", tmp_path / "jobs")
    result = _result(tmp_path / "result", job_folder, _mac_grids(), manifest_sha="0" * 64)
    with pytest.raises(ValueError, match="manifest_sha256"):
        job.read("cells", job_folder, result, tmp_path / "windows")


# --- the screen group borders_fills ---------------------------------------------------------------

COLUMN_PT, ROW_PT = 30.0, 24.0
"""The hand-built result's grid: every column and row alike, tall enough to scan an edge clear
of its corners at zoom 100."""


def test_writing_a_screen_job_again_gives_the_same_bytes(tmp_path):
    first = job.write("borders_fills", "sheetpeek-test-bf", tmp_path / "a")
    second = job.write("borders_fills", "sheetpeek-test-bf", tmp_path / "b")
    for name in ("job.json", "files/borders.xlsx", "files/fills.xlsx"):
        assert _sha(first / name) == _sha(second / name), name
    doc = json.loads((first / "job.json").read_text(encoding="utf-8"))
    assert [wb["file"] for wb in doc["workbooks"]] == ["borders.xlsx", "fills.xlsx"]
    for wb in doc["workbooks"]:
        ops = [s["op"] for s in wb["steps"]]
        assert ops == ["open", "view", "read", "read_format"] + ["view", "capture_window"] * 2 + [
            "close"
        ]
        assert [s["zoom"] for s in wb["steps"] if s["op"] == "view"] == [100, 100, 200]


def _geometry_grid(rows, columns, text=""):
    return {
        "rows": rows,
        "columns": columns,
        "left": [[c * COLUMN_PT for c in range(columns)] for _ in range(rows)],
        "top": [[r * ROW_PT for _ in range(columns)] for r in range(rows)],
        "width": [[COLUMN_PT] * columns for _ in range(rows)],
        "height": [[ROW_PT] * columns for _ in range(rows)],
        "text": [[text] * columns for _ in range(rows)],
    }


def _picture(path, window, zoom, gray):
    """A Windows window at 96 dpi: white, gridlines on every cell's bottom and right edge, the
    two markers filled magenta; nothing else drawn."""
    from capture.screen import png
    from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

    scale = 96 / 72 * zoom / 100
    width, height = round(15 * COLUMN_PT * scale), round(24 * ROW_PT * scale)
    template = bytearray(b"\xff" * (3 * width))
    for c in range(1, 16):
        x = round(c * COLUMN_PT * scale)
        if x < width:
            template[3 * x : 3 * x + 3] = bytes(gray)
    rows = [bytearray(template) for _ in range(height)]
    for r in range(1, 25):
        y = round(r * ROW_PT * scale)
        if y < height:
            rows[y] = bytearray(bytes(gray) * width)
    for marker in window.markers:
        letter, row = coordinate_from_string(marker)
        c = column_index_from_string(letter) - 1
        x0, x1 = round(c * COLUMN_PT * scale), round((c + 1) * COLUMN_PT * scale)
        y0, y1 = round((row - 1) * ROW_PT * scale), round(row * ROW_PT * scale)
        for y in range(y0, y1):
            rows[y][3 * x0 : 3 * x1] = b"\xff\x00\xff" * (x1 - x0)
    image = png.Image(width, height, 3, rows, "sRGB")
    png.write_crop(image, (0, 0, width, height), path)


def _screen_result(folder, job_folder, gray=(224, 224, 224)):
    (folder / "artifacts").mkdir(parents=True)
    books = []
    for window in job.WINDOWS_OF["borders_fills"]:
        white = {"interior.color": 16777215.0, "interior.pattern": -4142}
        steps = [
            {"op": "open", "status": "ok"},
            {"op": "read", "id": "geometry", "status": "ok", "result": _geometry_grid(24, 15)},
            {"op": "read_format", "id": "display", "status": "ok",
             "result": {"cells": [[white] * 15 for _ in range(24)]}},
        ]  # fmt: skip
        for zoom in window.zooms:
            name = f"{window.suite}_z{zoom}.png"
            _picture(folder / "artifacts" / name, window, zoom, gray)
            steps.append({"op": "view", "id": f"view_z{zoom}", "status": "ok",
                          "result": {"zoom": float(zoom)}})  # fmt: skip
            steps.append({"op": "capture_window", "id": f"shot_z{zoom}", "status": "ok",
                          "result": {"artifact": name, "dpi": 96}})  # fmt: skip
        books.append({"file": f"{window.suite}.xlsx", "status": "ok", "steps": steps})
    inputs = {wb["file"]: {"sha256": _sha(job_folder / "files" / wb["file"])} for wb in books}
    header = {
        "manifest_sha256": _sha(job_folder / "job.json"),
        "inputs": inputs,
        "environment": {"excel": {"build": "16.0.0.0"}},
    }
    outcome = {"status": "ok", "started": "2026-11-23T00:00:00Z", "header": header,
               "workbooks": books}  # fmt: skip
    (folder / "result.json").write_text(json.dumps(outcome), encoding="utf-8")
    return folder


def test_a_screen_result_reads_into_the_macs_schema_with_its_measured_gray(tmp_path):
    from capture import borders, fills

    job_folder = job.write("borders_fills", "sheetpeek-test-bf", tmp_path / "jobs")
    result = _screen_result(tmp_path / "result", job_folder, gray=(201, 202, 203))
    out = tmp_path / "windows"
    written = job.read("borders_fills", job_folder, result, out)
    assert [p.name for p in written] == ["borders.tsv", "fills.tsv"]
    for path, columns in zip(written, (borders.COLUMNS, fills.COLUMNS), strict=True):
        lines = path.read_text(encoding="utf-8").splitlines()
        head = [x for x in lines if x.startswith("# ")]
        assert "# platform: windows" in head
        assert "# excel_build: 16.0.0.0" in head
        # the gray is the picture's, whatever the platform's recorded one
        assert any(x.startswith("# gridline: C9CACB ") for x in head)
        body = [x.split("\t") for x in lines if not x.startswith("# ")]
        assert tuple(body[0]) == columns
    rows = (out / "borders.tsv").read_text(encoding="utf-8").splitlines()
    controls = [x.split("\t") for x in rows if "\tnone\t" in x]
    # the control's four edges at each zoom: one gridline row each, the gray's own
    assert len(controls) == 8
    assert {x[borders.COLUMNS.index("ink")] for x in controls} == {"C9CACB"}
    assert (out / "borders_fills.job.json").read_bytes() == (job_folder / "job.json").read_bytes()


def test_a_screen_result_without_a_gridline_is_refused(tmp_path):
    job_folder = job.write("borders_fills", "sheetpeek-test-bf", tmp_path / "jobs")
    result = _screen_result(tmp_path / "result", job_folder, gray=(255, 255, 255))
    with pytest.raises(ValueError, match="no gridline gray"):
        job.read("borders_fills", job_folder, result, tmp_path / "windows")


def test_display_reads_keep_the_macs_form():
    steps = {
        "geometry": _geometry_grid(1, 1, "x"),
        "display": {"cells": [[{"interior.color": 0xFF0000 + 0x80, "interior.pattern": -4125.0}]]},
    }
    rows = {(cell, key): value for cell, key, value in job._cell_rows(steps)}
    assert rows[("A1", "display.fill")] == "128,0,255"  # red in the low byte
    assert rows[("A1", "display.pattern")] == "-4125"
    assert rows[("A1", "geometry.width")] == repr(COLUMN_PT)
    assert rows[("A1", "geometry.text")] == "x"
