"""The HTML backend: static markup drawn from excel mode's geometry (ARCHITECTURE.md §7.3)."""

import re

import openpyxl
import pytest
from openpyxl.styles import Alignment, Font, PatternFill

import sheetpeek
from sheetpeek.backends import HTMLBackend
from tests.workbooks import save

LONG = "a long text value that spills"


def page(tmp_path, build, ref=None, **options):
    wb = openpyxl.Workbook()
    wb.active.title = "S & P"
    build(wb.active)
    sheet = sheetpeek.read(save(wb, tmp_path / "book.xlsx"), ref=ref)
    return sheetpeek.render(sheet, HTMLBackend(**options))


def tds(text, row=0):
    """The <td> elements of one body row."""
    rows = re.findall(r"<tr style[^>]*>(.*?)</tr>", text)
    return re.findall(r"<td.*?</td>", rows[row])


def test_page_and_fragment(tmp_path):
    whole = page(tmp_path, lambda ws: ws.cell(1, 1, 1))
    assert whole.startswith("<!doctype html>\n")
    assert "<title>S &amp; P</title>" in whole
    assert whole.endswith("</body>\n</html>\n")
    assert "<script" not in whole
    part = page(tmp_path, lambda ws: ws.cell(1, 1, 1), fragment=True)
    assert part.startswith('<style>\n@font-face { font-family: "sheetpeek-calibri";')
    assert '<div class="sheetpeek" data-color-scheme="light">' in part
    assert "<html" not in part


def test_geometry_in_whole_pixels(tmp_path):
    # Positions are rounded from points (x 4/3) and sizes are their differences: 53 pt columns
    # are 71, 70 and 71 px, so nothing accumulates (ARCHITECTURE.md §7.3).
    text = page(tmp_path, lambda ws: ws.cell(1, 1, 1), ref="A1:C2")
    assert re.findall(r'<col style="width:(\d+)px">', text) == ["71", "70", "71"]
    assert (
        "style=\"width:213px;font:11pt 'sheetpeek-calibri', 'sheetpeek-calibri-fallback', "
        'sans-serif"'
    ) in text
    assert re.findall(r'<tr style="height:(\d+)px">', text) == ["20", "20"]


def test_spill_fills_and_gridlines(tmp_path):
    def build(ws):
        ws["A1"] = LONG
        ws["B1"].fill = PatternFill("solid", fgColor="FFFFFF00")

    cells = tds(page(tmp_path, build, ref="A1:D2"))
    # B is crossed by A's text and filled: its fill covers its own and C's left gridline.
    assert (
        cells[1] == '<td class="sp-crossed" style="background:#FFFF00;border-color:#FFFF00"></td>'
    )
    assert cells[2].startswith('<td class="sp-crossed" style="border-left-color:#FFFF00">')
    assert cells[3] == "<td></td>"
    below = tds(page(tmp_path, build, ref="A1:D2"), row=1)
    assert below[1] == '<td style="border-top-color:#FFFF00"></td>'
    # The text is a run inside A at A's left inset (3 pt, 4 px). It lies inside its clip window
    # over B and C, so it needs no clip span; the baseline is 12 pt (16 px) down, less
    # Calibri's 0.8418 em.
    assert cells[0].startswith('<td><span class="sp-run" style="left:4px;top:3.654px;')


def test_anchors(tmp_path):
    def build(ws):
        ws["A1"] = 7  # right-aligned, inside its window: anchored by its end from the page's edge
        ws["B1"] = "c"
        ws["B1"].alignment = Alignment(horizontal="center")
        ws["A2"] = 1
        ws["B2"] = LONG  # right-aligned and blocked by A2: clipped, anchored by its end
        ws["B2"].alignment = Alignment(horizontal="right")
        ws["A3"] = "x"
        ws["A3"].font = Font(name="Arial")  # not spaced from Carlito: always clipped

    text = page(tmp_path, build, ref="A1:B3")
    a, b = tds(text)
    # The page is 142 px wide; the 7 ends 2 pt (to 68 px) before B's gridline: 142 - 68 = 74.
    assert a.startswith('<td><span class="sp-run" style="right:74px;')
    assert 'class="sp-run sp-center" style="left:' in b
    blocked = tds(text, 1)[1]
    assert '<span class="sp-clip" style="left:72px;top:21px;width:69px;height:19px">' in blocked
    # 2 pt before the next gridline: 141 px less the rounded end, 139 px
    assert 'class="sp-run" style="right:2px;' in blocked
    assert '<span class="sp-clip"' in tds(text, 2)[0]


def test_calibri_spacing_from_carlito(tmp_path):
    # Excel draws a Calibri 11 digit 6 pt wide; Calibri's own advance is 1038/2048 em,
    # 5.575 pt, so eight digits get (48 - 44.60) / 8 pt each.
    def build(ws):
        ws["A1"] = "00000000"
        ws["A2"] = "00000000"
        ws["A2"].font = Font(name="Arial")
        ws["A3"] = "<é&>"

    text = page(tmp_path, build, ref="A1:A3")
    assert "letter-spacing:0.425pt" in tds(text)[0]
    assert "letter-spacing" not in tds(text, 1)[0]  # not Calibri
    third = tds(text, 2)[0]
    assert "letter-spacing" not in third  # not printable ASCII
    assert "&lt;é&amp;&gt;</span>" in third
    assert "font:11pt 'Arial', sans-serif" in tds(text, 1)[0]


def test_merges_span(tmp_path):
    def build(ws):
        ws["A1"] = "merged"
        ws.merge_cells("A1:B2")

    text = page(tmp_path, build, ref="A1:C2")
    first, second = tds(text), tds(text, 1)
    assert first[0].startswith('<td colspan="2" rowspan="2">')
    assert len(first) == 2
    assert len(second) == 1  # A2 and B2 are covered


def test_headings_and_color_scheme(tmp_path):
    text = page(
        tmp_path,
        lambda ws: ws.cell(1, 1, 1),
        ref="A1:B1",
        column_headings=True,
        row_headings=True,
        color_scheme="dark",
    )
    assert 'data-color-scheme="dark"' in text
    assert '<th class="sp-corner"></th><th class="sp-colhead" scope="col">A</th>' in text
    assert '<th class="sp-rowhead" scope="row">1</th>' in text
    assert "#ffffff" in text  # the paper stays white
    with pytest.raises(ValueError, match="color_scheme"):
        HTMLBackend(color_scheme="auto")  # type: ignore[arg-type]


def test_hidden_gridlines(tmp_path):
    def build(ws):
        ws["A1"] = 1
        ws.sheet_view.showGridLines = False

    assert '<div class="sheetpeek sp-nogrid"' in page(tmp_path, build)


def test_needs_excel_mode(tmp_path):
    sheet = sheetpeek.read(save(openpyxl.Workbook(), tmp_path / "e.xlsx"))
    with pytest.raises(ValueError, match="excel mode"):
        HTMLBackend().render(sheet.layout("characters"), __import__("io").StringIO())
    assert sheetpeek.render(sheet.layout("characters"), "html").startswith("<!doctype html>")


def test_right_aligned_runs_are_anchored_by_their_end(tmp_path):
    # Another font than Calibri changes where right-aligned text starts, never where it ends:
    # every right-aligned run is placed with right:, clipped or not.
    def build(ws):
        for row in range(1, 6):
            ws.cell(row, 2, 10 ** (row + 3))  # numbers of 5 to 9 digits
        ws["A6"] = 1
        ws["B6"] = LONG
        ws["B6"].alignment = Alignment(horizontal="right")  # blocked by A6: clipped

    text = page(tmp_path, build, ref="A1:C6")
    styles = re.findall(r'<span class="sp-run" style="([^"]*)"', text)
    assert len(styles) == 7
    assert all(style.startswith("right:") for style in styles[:5] + styles[6:])


def test_calibri_faces(tmp_path):
    # Spaced runs use a face that is only a local Calibri or Carlito, else a local Arial or
    # Helvetica scaled to Calibri's widths with Calibri's vertical metrics.
    text = page(tmp_path, lambda ws: ws.cell(1, 1, 1))
    assert '@font-face { font-family: "sheetpeek-calibri"; src: local("Calibri"), ' in text
    assert 'local("Carlito")' in text
    assert re.search(
        r'@font-face \{ font-family: "sheetpeek-calibri-fallback"; src: local\("Arial"\).*'
        r"size-adjust: 91.6%; ascent-override: 103.93%; descent-override: 29.37%",
        text,
    )
