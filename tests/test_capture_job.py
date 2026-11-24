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


def _picture(path, window, zoom, gray, rows_n=24):
    """A Windows window at 96 dpi: white, gridlines on every cell's bottom and right edge, the
    two markers filled magenta; nothing else drawn."""
    from capture.screen import png
    from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

    scale = 96 / 72 * zoom / 100
    width, height = round(15 * COLUMN_PT * scale), round(rows_n * ROW_PT * scale)
    template = bytearray(b"\xff" * (3 * width))
    for c in range(1, 16):
        x = round(c * COLUMN_PT * scale)
        if x < width:
            template[3 * x : 3 * x + 3] = bytes(gray)
    rows = [bytearray(template) for _ in range(height)]
    for r in range(1, rows_n + 1):
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


def _screen_result(folder, job_folder, gray=(224, 224, 224), group="borders_fills", rows_n=24):
    (folder / "artifacts").mkdir(parents=True)
    books = []
    for window in job.WINDOWS_OF[group]:
        white = {"interior.color": 16777215.0, "interior.pattern": -4142}
        steps = [
            {"op": "open", "status": "ok"},
            {"op": "read", "id": "geometry", "status": "ok",
             "result": _geometry_grid(rows_n, 15)},
            {"op": "read_format", "id": "display", "status": "ok",
             "result": {"cells": [[white] * 15 for _ in range(rows_n)]}},
        ]  # fmt: skip
        for zoom in window.zooms:
            name = f"{window.stem}_z{zoom}.png"
            _picture(folder / "artifacts" / name, window, zoom, gray, rows_n)
            steps.append({"op": "view", "id": f"view_z{zoom}", "status": "ok",
                          "result": {"zoom": float(zoom)}})  # fmt: skip
            steps.append({"op": "capture_window", "id": f"shot_z{zoom}", "status": "ok",
                          "result": {"artifact": name, "dpi": 96}})  # fmt: skip
        books.append({"file": f"{window.stem}.xlsx", "status": "ok", "steps": steps})
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


def test_a_text_steps_job_has_a_window_a_group_and_reads_into_one_file(tmp_path):
    from capture import text_steps

    first = job.write("text_steps", "sheetpeek-test-ts", tmp_path / "a")
    second = job.write("text_steps", "sheetpeek-test-ts", tmp_path / "b")
    doc = json.loads((first / "job.json").read_text(encoding="utf-8"))
    names = [wb["file"] for wb in doc["workbooks"]]
    # shrink_fonts, added later, is a job of its own, so that this one is written as it ran
    groups = [g for g in text_steps.GROUPS if g != "shrink_fonts"]
    assert names == [f"text_steps_{group}.xlsx" for group in groups]
    for name in ["job.json", *(f"files/{n}" for n in names)]:
        assert _sha(first / name) == _sha(second / name), name
    # a hand-built result: pictures with markers and gridlines only, so no case has ink
    result = _screen_result(tmp_path / "result", first, group="text_steps", rows_n=32)
    (path,) = job.read("text_steps", first, result, tmp_path / "windows")
    lines = path.read_text(encoding="utf-8").splitlines()
    assert path.name == "text_steps.tsv"
    assert "# gridline: E0E0E0 (measured: D1's bottom edge)" in lines
    assert "# zooms: 100" in lines
    columns = next(x for x in lines if not x.startswith("# "))
    assert columns.split("\t") == list(text_steps.COLUMNS)


def test_a_text_style_job_reads_the_displayed_font_and_alignment(tmp_path):
    from capture import text_style

    folder = job.write("text_style", "sheetpeek-test-tst", tmp_path / "a")
    doc = json.loads((folder / "job.json").read_text(encoding="utf-8"))
    groups = [g for g in text_style.GROUPS if g != "line_budget"]  # a job of its own
    assert [wb["file"] for wb in doc["workbooks"]] == [
        f"text_style_{group}.xlsx" for group in groups
    ]
    for wb, group in zip(doc["workbooks"], groups, strict=True):
        (read,) = [s for s in wb["steps"] if s["op"] == "read_format"]
        assert read["props"] == [
            "font.name", "font.size", "font.underline", "align.horizontal", "align.vertical"
        ]  # fmt: skip
        zooms = [s["zoom"] for s in wb["steps"] if s.get("id", "").startswith("view_z")]
        assert zooms == list(text_style.ZOOMS[group])


def test_displayed_fonts_and_alignments_keep_the_macs_keys():
    cell = {"font.name": "Calibri", "font.size": 11.0, "font.underline": -4142.0,
            "align.horizontal": 1.0, "align.vertical": -4107.0}  # fmt: skip
    steps = {"geometry": _geometry_grid(1, 1, "aaa\nbbb\r\nc\td"), "display": {"cells": [[cell]]}}
    rows = {(c, key): value for c, key, value in job._cell_rows(steps)}
    assert rows[("A1", "geometry.text")] == "aaa bbb  c d"  # as the Mac's reads are written
    assert rows[("A1", "display.font_name")] == "Calibri"
    assert rows[("A1", "display.font_size")] == "11"
    assert rows[("A1", "display.underline")] == "-4142"
    assert rows[("A1", "display.h_align")] == "1"
    assert rows[("A1", "display.v_align")] == "-4107"


def test_the_shrink_fonts_job_reads_into_a_file_of_its_own(tmp_path):
    from capture import text_steps

    first = job.write("shrink_fonts", "sheetpeek-test-sf", tmp_path / "a")
    doc = json.loads((first / "job.json").read_text(encoding="utf-8"))
    assert [wb["file"] for wb in doc["workbooks"]] == ["text_steps_shrink_fonts.xlsx"]
    result = _screen_result(tmp_path / "result", first, group="shrink_fonts", rows_n=16)
    (path,) = job.read("shrink_fonts", first, result, tmp_path / "windows")
    assert path.name == "shrink_fonts.tsv"
    lines = path.read_text(encoding="utf-8").splitlines()
    columns = next(x for x in lines if not x.startswith("# "))
    assert columns.split("\t") == list(text_steps.COLUMNS)


def test_the_line_budget_job_reads_into_a_file_of_its_own(tmp_path):
    from capture import text_style

    first = job.write("line_budget", "sheetpeek-test-lb", tmp_path / "a")
    doc = json.loads((first / "job.json").read_text(encoding="utf-8"))
    assert [wb["file"] for wb in doc["workbooks"]] == ["text_style_line_budget.xlsx"]
    result = _screen_result(tmp_path / "result", first, group="line_budget", rows_n=22)
    (path,) = job.read("line_budget", first, result, tmp_path / "windows")
    assert path.name == "line_budget.tsv"
    lines = path.read_text(encoding="utf-8").splitlines()
    columns = next(x for x in lines if not x.startswith("# "))
    assert columns.split("\t") == list(text_style.COLUMNS)
    assert len(text_style.cases("line_budget")) == 136  # 34 texts, four alignments each


def test_the_font_cut_job_rewrites_its_themes_and_reads_font_colors(tmp_path):
    import zipfile

    from capture import font_cut

    folder = job.write("font_cut", "sheetpeek-test-fc", tmp_path / "a")
    doc = json.loads((folder / "job.json").read_text(encoding="utf-8"))
    assert [wb["file"] for wb in doc["workbooks"]] == [
        f"font_cut_{group}.xlsx" for group in font_cut.GROUPS if group not in font_cut.FACES
    ]
    for wb in doc["workbooks"]:
        (read,) = [s for s in wb["steps"] if s["op"] == "read_format"]
        assert read["props"] == ["font.name", "font.size", "font.color"]
    with zipfile.ZipFile(folder / "files" / "font_cut_theme_16.xlsx") as z:
        theme = z.read("xl/theme/theme1.xml").decode()
    assert '<a:minorFont>\n        <a:latin typeface="Arial"/>' in theme
    cell = {"font.name": "Arial", "font.size": 24.0, "font.color": 0x0000FF + 0.0}
    steps = {"geometry": _geometry_grid(1, 1), "display": {"cells": [[cell]]}}
    rows = {(c, key): value for c, key, value in job._cell_rows(steps)}
    assert rows[("A1", "display.font_color")] == "255,0,0"  # red in the low byte
    mixed = {"font.name": None, "font.size": None, "font.color": None}  # a rich-text cell
    steps = {"geometry": _geometry_grid(1, 1), "display": {"cells": [[mixed]]}}
    rows = {(c, key): value for c, key, value in job._cell_rows(steps)}
    assert rows[("A1", "display.font_size")] == rows[("A1", "display.font_color")] == ""
