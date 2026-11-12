"""The places a feature extends the pipeline, so features can be built side by side."""

import io

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


def test_a_scan_handler_keeps_what_it_finds_in_extra(monkeypatch):
    found = []

    def outline(state, element):
        found.append(element.get("outlineLevel"))
        state.extra["outline"] = found

    monkeypatch.setitem(
        _structure.HANDLERS, _structure._ROW, [*_structure.HANDLERS[_structure._ROW], outline]
    )
    xml = (
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        b'<row r="1" outlineLevel="2"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'
    )
    structure = _structure.scan(io.BytesIO(xml))
    assert structure.extra == {"outline": ["2"]}
    assert structure.used_range == sheetpeek.CellRange(1, 1, 1, 1)  # the core handler still ran


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
