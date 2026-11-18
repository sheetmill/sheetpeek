"""The places a feature extends the pipeline, so features can be built side by side."""

import openpyxl

import sheetpeek
from sheetpeek import _detect, _layout, _structure, _style
from tests.workbooks import replace_once, rewrite, save

SHEET1 = "xl/worksheets/sheet1.xml"


def _workbook(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "x"
    wb.active.row_dimensions[1].outlineLevel = 1
    return save(wb, tmp_path / "seams.xlsx")


def test_a_feature_keeps_what_it_reads_in_extra(tmp_path):
    # Each feature adds what it reads of the sheet to Structure.extra under its own key, so
    # features meet only in the reader's pipeline (the outline here).
    from sheetpeek._notices import Collector
    from sheetpeek._read_outline import KEY, with_levels

    with sheetpeek.open(_workbook(tmp_path)) as book:
        source = book._book.sheet(0)
        found, _ = book._stream(source, None, Collector("sheet"))
        structure = with_levels(_structure.fill(found, source), source)
    assert structure.extra[KEY].rows == {1: 1}
    assert structure.used_range == sheetpeek.CellRange(1, 1, 1, 1)


def test_a_layout_pass_sees_the_draft(tmp_path, monkeypatch):
    seen = []
    monkeypatch.setattr(_layout, "PASSES", [*_layout.PASSES, lambda d: seen.append(len(d.rows))])
    sheetpeek.read(_workbook(tmp_path)).layout("excel")
    assert seen == [1]


def test_a_rule_replaces_the_unknown_element_report(tmp_path, monkeypatch):
    path = _workbook(tmp_path)
    rewrite(path, SHEET1, replace_once("</worksheet>", "<futureFeature /></worksheet>"))
    assert "sheet.unknown_element" in {n.code for n in sheetpeek.read(path).notices}
    monkeypatch.setitem(_detect.RULES, "futureFeature", lambda context, child: None)
    assert "sheet.unknown_element" not in {n.code for n in sheetpeek.read(path).notices}


def test_style_aspects_each_own_their_findings():
    assert [a.ATTRIBUTE for a in _style.ASPECTS] == [
        "number_format", "border", "fill", "font", "alignment"
    ]  # fmt: skip
    for aspect in _style.ASPECTS:
        assert callable(aspect.read)
        assert callable(aspect.unread)
