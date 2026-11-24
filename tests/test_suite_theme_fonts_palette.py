"""The screen suite ``theme_fonts_palette``: its workbooks, its measuring, and the comparison.

ARCHITECTURE.md §4.14 (theme fonts, the workbook palette, number-format colors) and §3.3. The
workbooks are built and read back with sheetpeek; the measuring runs on a synthetic capture
(``Canvas``); the comparison with ``tests/captures/theme_fonts_palette.tsv`` is skipped until
the suite has been captured.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import openpyxl
import pytest
from capture import theme_fonts_palette as suite
from capture.screen.run import write_tsv
from capture.screen.window import MARKER

import sheetpeek
from tests.capture_report import REPORT
from tests.screen_canvas import Canvas

CAPTURE = Path(__file__).with_name("captures") / "theme_fonts_palette.tsv"
FILL_TOLERANCE = 8
"""Per channel, for a fill's pixels: the screen's color conversion moves saturated colors by up
to 6 (``00B050`` showed as ``06B050`` in the suite ``tables``)."""
INK_TOLERANCE = 16
"""Per channel, for the core of 18 pt bold strokes (about 5 device pixels wide): the conversion's
6, and what smoothing leaves in the pixels at 90 % of the farthest from the background."""
DEFAULT_COLUMN = 53.0
"""Excel for Mac's default column at Calibri 11, in points (``layout``, ARCHITECTURE §4.1)."""
KNOWN: dict[tuple[str, str], str] = {}
"""Width cases where sheetpeek is known to differ from Excel, with the reason."""


def _built(tmp_path, group):
    wb = openpyxl.Workbook()
    suite.fill(wb.active, group)
    path = tmp_path / f"{group}.xlsx"
    wb.save(path)
    post = suite.postprocess(group)
    if post is not None:
        post(path)
    return path


# --- the workbooks


def test_every_group_builds_and_reads_back_as_built(tmp_path):
    placed = suite.build(tmp_path)
    assert len(placed) == sum(len(suite.cases(g)) for g in suite.GROUPS)
    for group in suite.GROUPS:
        path = tmp_path / f"theme_fonts_palette_{group}.xlsx"
        assert suite.sheetpeek_view(path, group) == suite.EXPECTED[group], group
    with zipfile.ZipFile(tmp_path / "theme_fonts_palette_no_theme.xlsx") as package:
        assert not [n for n in package.namelist() if "theme" in n]
        assert "theme" not in package.read("[Content_Types].xml").decode()
    with zipfile.ZipFile(tmp_path / "theme_fonts_palette_palette_bad.xlsx") as package:
        styles = package.read("xl/styles.xml").decode()
    assert len(re.findall("<rgbColor ", styles)) == 64
    assert re.findall(r'<rgbColor rgb="([^"]*)"', styles)[17] == "zz"
    with zipfile.ZipFile(tmp_path / "theme_fonts_palette_palette_long.xlsx") as package:
        assert package.read("xl/styles.xml").decode().count("<rgbColor ") == 70


def test_the_fonts_theme_is_as_the_cases_say(tmp_path):
    sheet = sheetpeek.read(_built(tmp_path, "fonts"))
    b4, b5, b6, b8 = (sheet.cell(r, 2).font for r in (4, 5, 6, 8))
    assert (b4.stored_name, b4.scheme) == ("Calibri", "minor")
    assert (b5.stored_name, b5.scheme) == ("Times New Roman", "minor")
    assert (b6.stored_name, b6.scheme) == ("Calibri", "major")
    assert b8 == sheet.normal_font
    with zipfile.ZipFile(_built(tmp_path, "fonts")) as package:
        theme = package.read("xl/theme/theme1.xml").decode()
    minor = theme[theme.index("<a:minorFont>") : theme.index("</a:minorFont>")]
    major = theme[theme.index("<a:majorFont>") : theme.index("</a:majorFont>")]
    assert '<a:latin typeface="Courier New"' in minor
    assert f'<a:font script="Jpan" typeface="{suite.MINCHO}"' in minor
    assert '<a:latin typeface="Georgia"' in major
    assert f'<a:font script="Jpan" typeface="{suite.MARU}"' in major


def test_every_case_names_a_distinct_cell_away_from_the_markers():
    for group in suite.GROUPS:
        found = suite.cases(group)
        cells = [case.cell for case in found]
        assert len(set(cells)) == len(cells), group
        assert set(cells) == set(suite.EXPECTED[group]), group
        if group in suite.WINDOWED:
            assert not {"B2", suite.MARKERS[group]} & set(cells), group


def test_no_case_draws_a_color_near_the_marker():
    marker = [int(MARKER[i : i + 2], 16) for i in (0, 2, 4)]
    drawn = {v for g in suite.WINDOWED for v in suite.EXPECTED[g].values() if isinstance(v, str)}
    for color in drawn | set(suite.RECOLORED.values()):
        if not re.fullmatch("[0-9A-F]{6}", color):
            continue
        rgb = [int(color[i : i + 2], 16) for i in (0, 2, 4)]
        assert max(abs(a - b) for a, b in zip(rgb, marker, strict=True)) > 96, color


def test_every_text_case_has_a_control_in_its_face():
    for group, controls in suite.CONTROLS.items():
        for case in suite.cases(group):
            if case.kind == "text" and "Japanese" not in case.case:
                face = suite.EXPECTED[group][case.cell]
                # no_theme draws Excel's own minor font, which the window has no control of
                assert face in controls or (group, face) == ("no_theme", "Aptos Narrow"), face


# --- the measuring, on a synthetic capture


def _canvas(group, reads=()):
    ws = openpyxl.Workbook().active
    suite.fill(ws, group)
    return Canvas(ws, [c.cell for c in suite.cases(group)], rows=reads)


def test_text_ink_reads_as_a_bitmap_with_its_size_in_points():
    canvas = _canvas("fonts")
    for cell in ("B4", "B9"):  # the same mark in two cells of one column: the same ink
        x0, y0, _, _ = canvas.box(cell)
        canvas.paint(x0 + 6, y0 + 8, x0 + 46, y0 + 22, (0, 0, 0))
    x0, y0, _, _ = canvas.box("B10")
    canvas.paint(x0 + 6, y0 + 8, x0 + 50, y0 + 22, (0, 0, 0))
    records = {r["cell"]: r for r in suite.measure(canvas.capture(), "fonts")}
    assert (records["B4"]["ink_w"], records["B4"]["ink_h"]) == ("20", "7")
    assert records["B4"]["ink"] == records["B9"]["ink"]
    assert records["B10"]["ink"] != records["B4"]["ink"]
    assert records["B4"]["ink_color"] == "000000"
    assert "ink" not in records["B5"]  # nothing drawn
    assert set(suite.COLUMNS) >= set(records["B4"])


def test_a_strokes_color_is_its_core_not_its_smoothed_edge():
    canvas = _canvas("palette")
    x0, y0, _, _ = canvas.box("F4")
    canvas.paint(x0 + 10, y0 + 10, x0 + 30, y0 + 30, (0xC0, 0xA0, 0x80))  # smoothing
    canvas.paint(x0 + 12, y0 + 12, x0 + 28, y0 + 28, (0x80, 0x40, 0x00))  # the core
    record = next(r for r in suite.measure(canvas.capture(), "palette") if r["cell"] == "F4")
    assert record["ink_color"] == suite.BROWN
    assert (record["ink_w"], record["ink_h"]) == ("10", "10")


def test_a_fill_reads_as_its_center_color():
    canvas = _canvas("palette")
    canvas.paint(*canvas.box("B6"), (0x00, 0x80, 0x80))
    records = {r["cell"]: r for r in suite.measure(canvas.capture(), "palette")}
    assert records["B6"]["pixel"] == suite.TEAL
    assert records["B4"]["pixel"] == "FFFFFF"


def test_excels_reads_are_written_as_hex_and_text():
    reads = [
        ("B6", "display.fill", "0,128,128"),
        ("B6", "display.font_color", "0,0,0"),
        ("B6", "display.font_name", "Calibri"),
        ("B6", "geometry.width", "120.0"),
        ("B6", "geometry.height", "15.0"),
    ]
    record = next(
        r
        for r in suite.measure(_canvas("palette", reads).capture(), "palette")
        if r["cell"] == "B6"
    )
    assert (record["fill"], record["font_color"], record["font_name"]) == (
        suite.TEAL,
        "000000",
        "Calibri",
    )
    assert (record["width"], record["height"]) == ("120.0", "15.0")


def test_write_takes_read_only_groups_and_refusals(tmp_path, monkeypatch):
    monkeypatch.setattr(suite, "BUILD_DIR", tmp_path)
    monkeypatch.setattr(suite, "OUTPUT", tmp_path / "out.tsv")
    header = {"excel": "16.115", "appearance": "dark", "suite": "x", "markers": "B2,J20"}
    (tmp_path / "palette_short").mkdir()
    write_tsv(
        tmp_path / "palette_short" / "formats.tsv",
        header,
        [("B4", "display.fill", "0,128,128"), ("B6", "display.fill", "0,102,204")],
    )
    suite.record_refusal("palette_bad", "osascript: execution error: Excel refused (-50)")
    suite.write()
    lines = (tmp_path / "out.tsv").read_text(encoding="utf-8").splitlines()
    assert lines[:3] == ["# excel: 16.115", "# appearance: dark", "# suite: theme_fonts_palette"]
    rows = [dict(zip(suite.COLUMNS, line.split("\t"), strict=True)) for line in lines[4:]]
    assert [(r["group"], r["cell"], r["fill"]) for r in rows[:2]] == [
        ("palette_short", "B4", suite.TEAL),
        ("palette_short", "B5", ""),
    ]
    assert rows[2]["fill"] == "0066CC"
    assert rows[-1]["group"] == "palette_bad"
    assert rows[-1]["case"] == "refused: Excel refused (-50)"


# --- the comparison with Excel


@pytest.fixture(scope="module")
def captured():
    if not CAPTURE.exists():
        pytest.skip("not captured yet: tests/captures/theme_fonts_palette.tsv")
    lines = [ln for ln in CAPTURE.read_text(encoding="utf-8").splitlines() if ln[:2] != "# "]
    names = lines[0].split("\t")
    rows = [dict(zip(names, line.split("\t"), strict=False)) for line in lines[1:]]
    refused = {r["group"]: r["case"] for r in rows if r["case"].startswith("refused")}
    if refused:
        REPORT["theme_fonts_palette refused"] = "theme_fonts_palette: " + "; ".join(
            f"{group} {reason}" for group, reason in refused.items()
        )
    return {(r["group"], r["cell"]): r for r in rows if r["group"] not in refused}


def _near(a, b, tolerance):
    return all(abs(int(a[i : i + 2], 16) - int(b[i : i + 2], 16)) <= tolerance for i in (0, 2, 4))


@pytest.mark.excel
def test_text_draws_in_the_face_sheetpeek_names(captured):
    # A scheme font draws the theme's Latin face whatever it stores (Calibri-minor and
    # Times-New-Roman-minor as Courier New, Calibri-major as Georgia); the Normal style too;
    # without a theme Excel's own minor font, Aptos Narrow; with an empty typeface the stored
    # name. A case draws its control's ink
    # exactly: the same face, size and place in the column.
    compared = reported = 0
    for group, controls in suite.CONTROLS.items():
        for case in suite.cases(group):
            row = captured.get((group, case.cell))
            if row is None or case.kind != "text" or "Japanese" in case.case:
                continue
            face = suite.EXPECTED[group][case.cell]
            assert row["ink"], (group, case.cell)
            if face not in controls:
                # no control of the face (no_theme's Aptos Narrow): Excel's own read names it,
                # and every case of the face draws one ink
                assert row["font_name"] == face, (group, case.cell, row["font_name"])
                same = [c.cell for c in suite.cases(group) if suite.EXPECTED[group][c.cell] == face]
                assert {captured[(group, c)]["ink"] for c in same} == {row["ink"]}, same
                compared += 1
                reported += 1
                continue
            control = captured[(group, controls[face])]
            assert row["ink"] == control["ink"], (group, case.cell, face)
            compared += 1
            reported += row["font_name"] == face
    if ("fonts", "B7") in captured:  # Calibri without a scheme keeps its own face
        assert captured[("fonts", "B7")]["ink"] != captured[("fonts", "B9")]["ink"]
    REPORT["theme_fonts_palette faces"] = (
        f"theme_fonts_palette: {compared} text cells draw their face's ink; "
        f"{reported} of them have Excel report that face's name"
    )


@pytest.mark.excel
@pytest.mark.xfail(
    strict=True,
    reason="16.115 draws Japanese text in a scheme font in the theme's Jpan face; sheetpeek "
    "reads only the theme's Latin face (a known deviation, ARCHITECTURE §4.14)",
)
def test_japanese_text_draws_in_the_themes_script_font(captured):
    # The minor font's Jpan face (Hiragino Mincho ProN) and the major's (Hiragino Maru Gothic
    # ProN), not the system's fallback for the Latin face.
    if ("fonts", "B13") not in captured:
        pytest.skip("group fonts not captured")
    for cell, control in suite.JAPANESE_CONTROLS.items():
        row, other = captured[("fonts", cell)], captured[("fonts", control)]
        assert row["ink"], cell
        assert row["ink"] == other["ink"], cell
    assert captured[("fonts", "B15")]["ink"] != captured[("fonts", "B16")]["ink"]


@pytest.mark.excel
def test_column_widths_follow_the_normal_font(captured):
    # Exact: Excel's widths are whole pixels at 96 dpi (0.75 pt steps), as sheetpeek's are.
    deviations = []
    for (group, cell), row in captured.items():
        if row["kind"] != "width":
            continue
        if (group, cell) in KNOWN:
            deviations.append((group, cell))
            continue
        assert float(row["width"]) == suite.EXPECTED[group][cell], (group, cell)
    if ("fonts", "D4") in captured:  # the default column follows the theme's Courier New
        assert float(captured[("fonts", "D4")]["width"]) != DEFAULT_COLUMN
    if deviations:
        REPORT["theme_fonts_palette widths"] = (
            f"theme_fonts_palette: {len(deviations)} known width deviations ("
            + "; ".join(f"{g} {c}: {KNOWN[(g, c)]}" for g, c in deviations)
            + ")"
        )


@pytest.mark.excel
def test_indexed_colors_follow_the_workbook_palette(captured):
    # Fills (read exactly, and on screen) and font colors: 0 to 7 no longer mirror 8 to 15, a
    # short list keeps the default past its end, a non-color entry and an index past 64 as
    # sheetpeek reads them.
    for (group, cell), row in captured.items():
        expected = suite.EXPECTED[group][cell]
        if row["kind"] == "fill":
            if (group, cell) == ("palette_long", "B6"):
                # indexed 66 of a list of 70: Excel draws 646464, neither the list's entry nor
                # sheetpeek's black for an index past the palette (a known deviation)
                assert row["fill"] == "646464", (group, cell)
                continue
            assert row["fill"] == expected, (group, cell)
            if row.get("pixel"):
                assert _near(row["pixel"], expected, FILL_TOLERANCE), (group, cell)
        elif row["kind"] == "font":
            if (group, cell) == ("palette", "D5"):
                # indexed 3, recolored: Excel's display format reports the workbook's entry
                # (996600) and draws the default's (00FF00), as sheetpeek draws it
                assert row["font_color"] == suite.OLIVE, (group, cell)
            else:
                assert row["font_color"] == expected, (group, cell)
            assert _near(row["ink_color"], expected, INK_TOLERANCE), (group, cell)


@pytest.mark.excel
def test_number_format_colors_take_the_palette_but_named_colors_do_not(captured):
    # [ColorN] draws palette N + 7; [Red] and [Blue] draw red and blue though the palette
    # recolors their slots. The screen is the truth: Excel's display format reports palette
    # indices for the named colors.
    found = [row for row in captured.values() if row["kind"] == "format"]
    for row in found:
        expected = suite.EXPECTED[row["group"]][row["cell"]]
        assert _near(row["ink_color"], expected, INK_TOLERANCE), row
