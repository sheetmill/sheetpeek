"""The sheet view as drawn: zoom, frozen panes, the print-area option and page-break marks.

ARCHITECTURE.md §4.18. The HTML backend draws the zoom and the panes; every output takes the
print-area option.
"""

import re

import openpyxl
import pytest
from openpyxl.styles import Border, Font, PatternFill, Side

import sheetpeek
from sheetpeek.backends import HTMLBackend
from tests.workbooks import rewrite, save

SHEET1 = "xl/worksheets/sheet1.xml"


def _workbook(tmp_path, attributes="", pane=""):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet["A1"] = "Total"
    sheet["A1"].font = Font(underline="single", size=14)
    sheet["B2"] = 1234
    sheet["B2"].fill = PatternFill("solid", fgColor="FFFF00")
    sheet["C3"].border = Border(left=Side("thick"), bottom=Side("thin"))
    sheet.column_dimensions["B"].width = 20
    sheet.row_dimensions[2].height = 30
    path = save(workbook, tmp_path / "view.xlsx")
    if attributes or pane:
        rewrite(
            path,
            SHEET1,
            lambda text: text.replace(
                '<sheetView workbookViewId="0">',
                f'<sheetView workbookViewId="0" {attributes}>{pane}',
                1,
            ),
        )
    return path


def _html(path, **options):
    options.setdefault("column_headings", True)
    options.setdefault("row_headings", True)
    return sheetpeek.render(path, HTMLBackend(**options), on_notice="ignore")


def _sizes(page):
    """The page's sheet-sized quantities: column widths, row heights, the table's width."""
    columns = [int(w) for w in re.findall(r'<col style="width:(\d+)px">', page)]
    rows = [int(h) for h in re.findall(r'<tr style="height:(\d+)px">', page)]
    table = int(re.search(r'<table style="width:(\d+)px', page).group(1))
    return columns, rows, table


def _fonts(page):
    return [float(s) for s in re.findall(r"font:(?:\w+ )*?([\d.]+)pt", page)]


def test_the_zoom_saved_in_the_file_is_the_default(tmp_path):
    saved = _html(_workbook(tmp_path, 'zoomScale="200"'))
    assert saved == _html(_workbook(tmp_path), zoom=200)
    assert 'style="--sheetpeek-zoom:2"' in saved


def test_an_explicit_zoom_overrides_the_saved_one(tmp_path):
    path = _workbook(tmp_path, 'zoomScale="200"')
    assert _html(path, zoom=100) == _html(_workbook(tmp_path))


def test_sheet_sizes_scale_with_the_zoom(tmp_path):
    path = _workbook(tmp_path)
    (c1, r1, t1), (c2, r2, t2) = _sizes(_html(path)), _sizes(_html(path, zoom=200))
    # Positions are rounded to whole pixels, so a size can differ from twice by a pixel.
    assert all(abs(b - 2 * a) <= 1 for a, b in zip(c1 + r1, c2 + r2, strict=True))
    assert abs((t2 - 1) - 2 * (t1 - 1)) <= 1  # the table's width less its last gridline
    assert [2 * f for f in _fonts(_html(path))] == _fonts(_html(path, zoom=200))


def test_lines_tiles_and_symbols_keep_their_screen_size(tmp_path):
    path = _workbook(tmp_path)
    at_100, at_200 = _html(path), _html(path, zoom=200)
    # Borders, rules and gridlines are drawn in CSS px that do not depend on the zoom.
    for pattern in (r"border-left:(\d+)px", r"border-bottom:(\d+)px", r'class="sp-rule"'):
        assert re.findall(pattern, at_100) == re.findall(pattern, at_200)
    assert ".sp-rule { position: absolute; height: 1px;" in at_200
    assert "border-left: 1px solid var(--sheetpeek-grid)" in at_200


def test_the_headings_font_scales_with_the_zoom(tmp_path):
    page = _html(_workbook(tmp_path), zoom=50)
    assert "font-size: calc(9pt * var(--sheetpeek-zoom, 1))" in page
    assert 'style="--sheetpeek-zoom:0.5"' in page


@pytest.mark.parametrize("zoom", [9, 401, 0, 100.0, "100", True])
def test_a_zoom_excel_does_not_allow_is_refused(zoom):
    with pytest.raises(ValueError, match="zoom"):
        HTMLBackend(zoom=zoom)


@pytest.mark.parametrize("zoom", [10, 400])
def test_the_ends_of_the_zoom_range_are_drawn(tmp_path, zoom):
    assert "<table" in _html(_workbook(tmp_path), zoom=zoom)


def test_the_cli_takes_a_zoom_for_html_only(tmp_path, capsys):
    from sheetpeek.cli import main

    path = str(_workbook(tmp_path))
    assert main(["view", path, "-t", "html", "--zoom", "150", "--notices", "none"]) == 0
    assert "--sheetpeek-zoom:1.5" in capsys.readouterr().out
    assert main(["view", path, "-t", "text", "--zoom", "150"]) == 2
    assert "HTML only" in capsys.readouterr().err
    assert main(["view", path, "-t", "html", "--zoom", "5"]) == 2
    assert main(["view", path, "--all-sheets", "--zoom", "150", "--notices", "none"]) == 0
    assert "--sheetpeek-zoom:1.5" in capsys.readouterr().out
