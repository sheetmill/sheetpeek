"""Fonts past Excel's font table: drawn and measured as Excel for Mac draws them, reported.

The workbooks are the screen suite ``font_cut``'s (``capture/font_cut.py``); the comparison with
its capture is at the end.
"""

from pathlib import Path

import pytest

import sheetpeek

pytest.importorskip("openpyxl")

from capture import font_cut

CALIBRI_12 = ("Calibri", 12.0)


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    return font_cut.build(tmp_path_factory.mktemp("font_cut"))


def _at(cell):
    return int(cell[1:]), 3  # column C


def _cut_cells(sheet):
    return {
        (row, column)
        for n in sheet.notices
        if n.code == "font.cut"
        for r in n.cells
        for row in range(r.min_row, r.max_row + 1)
        for column in range(r.min_col, r.max_col + 1)
    }


def test_a_list_font_past_the_cut_is_drawn_in_calibri_12_on_the_mac(built):
    sheet = sheetpeek.read(built["list"], platform="mac")
    for cell, entry in font_cut.probes("list"):
        font = sheet.cell(*_at(cell)).style.font
        if entry >= 1000:
            assert (font.name, font.size, font.color) == (*CALIBRI_12, None), cell
        else:
            assert (font.name, font.size) == ("Arial", 24.0), cell
    assert _cut_cells(sheet) == {_at(c) for c, e in font_cut.probes("list") if e >= 1000}
    (notice,) = {(n.outcome, n.reason) for n in sheet.notices if n.code == "font.cut"}
    assert notice == ("approximated", "file")


def test_a_run_past_the_cut_keeps_its_color_on_the_mac(built):
    sheet = sheetpeek.read(built["rich"], platform="mac")
    for cell, entry in font_cut.probes("rich"):
        (span,) = sheet.cell(*_at(cell)).runs
        if entry >= 1000:
            assert (span.font.name, span.font.size) == CALIBRI_12, cell
        else:
            assert (span.font.name, span.font.size) == ("Arial", 24.0), cell
        assert span.font.color is not None, cell
        assert span.font.color.upper().endswith("0000"), cell  # its own red
    assert _cut_cells(sheet) == {_at(c) for c, e in font_cut.probes("rich") if e >= 1000}


def test_windows_draws_a_font_past_the_cut_in_aptos_narrow_11(built):
    # The cut is at 1,000 under openpyxl's Calibri theme and Calibri 11 Normal font (group list)
    # and at 999 under a Times New Roman Normal font, not the theme's body font (theme_10).
    for group, first in (("list", 1000), ("theme_10", 999), ("faces_aptos", 1000)):
        sheet = sheetpeek.read(built[group], platform="windows")
        found = font_cut.probes(group)
        for cell, entry in found:
            font = sheet.cell(*_at(cell)).style.font
            if entry >= first:
                assert (font.name, font.size, font.color) == ("Aptos Narrow", 11.0, None), cell
        assert _cut_cells(sheet) == {_at(c) for c, e in found if e >= first}, group
        notices = {(n.outcome, n.reason) for n in sheet.notices if n.code == "font.cut"}
        assert notices == (
            {("approximated", "file")} if any(e >= first for _, e in found) else set()
        )


def test_a_run_past_the_cut_keeps_its_color_on_windows(built):
    sheet = sheetpeek.read(built["rich"], platform="windows")
    for cell, entry in font_cut.probes("rich"):
        (span,) = sheet.cell(*_at(cell)).runs
        if entry >= 1000:
            assert (span.font.name, span.font.size) == ("Aptos Narrow", 11.0), cell
        assert span.font.color is not None, cell
        assert span.font.color.upper().endswith("0000"), cell  # its own red


def test_a_workbook_short_of_the_cut_reports_nothing(tmp_path):
    import openpyxl
    from openpyxl.styles import Font

    wb = openpyxl.Workbook()
    for k in range(1, 60):
        wb.active.cell(k, 1, "x").font = Font(name="Arial", size=8 + k % 20, color=f"FF{k:06X}")
    path = tmp_path / "few.xlsx"
    wb.save(path)
    for platform in ("mac", "windows"):
        assert _cut_cells(sheetpeek.read(path, platform=platform)) == set()


@pytest.mark.excel
def test_rows_fit_as_excel_for_mac_fits_them(built):
    # The capture's fitted rows (font_cut.tsv) against sheetpeek's layout with platform="mac":
    # 30 pt for Arial 24, from entry 1,000 the larger of Calibri 12's 16 pt and the Normal row.
    path = Path(__file__).with_name("captures") / "font_cut.tsv"
    lines = path.read_text(encoding="utf-8").split("\n")
    rows = [line.split("\t") for line in lines if line and not line.startswith("#")]
    records = [dict(zip(rows[0], row, strict=True)) for row in rows[1:]]
    layouts = {g: sheetpeek.read(built[g], platform="mac").layout("excel") for g in font_cut.GROUPS}
    wrong = {}
    for r in records:
        heights = {row.index: row.height_pt for row in layouts[r["group"]].rows}
        ours = heights[int(r["cell"][1:])]
        if ours != float(r["row_height"]):
            wrong[(r["group"], r["cell"])] = (float(r["row_height"]), ours)
    assert len(records) == 50
    assert wrong == {}
