"""The sheet view as drawn: zoom, frozen panes, the print-area option and page-break marks.

ARCHITECTURE.md §4.18. The HTML backend draws the zoom and the panes; every output takes the
print-area option.
"""

import re

import openpyxl
import pytest
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.hyperlink import Hyperlink

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


def _frozen(tmp_path, x=1, y=2, top_left="B3", view_top_left="", state="frozen", rows=8):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for row in range(1, rows + 1):
        for column in range(1, 5):
            sheet.cell(row, column).value = f"r{row}c{column}"
    sheet["D2"].hyperlink = Hyperlink(ref="D2", location="A1")
    path = save(workbook, tmp_path / "frozen.xlsx")
    pane = f'<pane xSplit="{x}" ySplit="{y}" topLeftCell="{top_left}" state="{state}"/>'
    attributes = f'topLeftCell="{view_top_left}"' if view_top_left else ""
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


def _panes(page):
    """Each pane copy's markup by its class."""
    pattern = r'<div class="sp-pane sp-pane-(\w+)"[^>]*>(.*?)</div>\n</div>'
    return dict(re.findall(pattern, page, re.S))


def test_a_frozen_pane_holds_its_rows_and_columns_on_a_page(tmp_path):
    page = _html(_frozen(tmp_path))
    assert page.count('<div class="sp-panes">') == 1
    panes = _panes(page)
    assert set(panes) == {"left", "top", "corner"}
    # The copies hold only the frozen rows (with the headings' row) or the frozen columns.
    assert panes["top"].count("<tr") == 3
    assert panes["corner"].count("<tr") == 3
    assert "r1c1" in panes["corner"]
    assert "r1c2" not in panes["corner"]
    assert "r8c1" in panes["left"]
    assert "r8c2" not in panes["left"]
    assert "r3c1" not in panes["top"]
    # Cell ids stay unique: only the sheet itself carries them.
    assert page.count('id="sp-A1"') == 1
    assert "position: sticky" in page
    assert "<script" not in page.lower()


def test_the_freeze_lines_are_drawn_where_the_split_is(tmp_path):
    page = _html(_frozen(tmp_path), column_headings=False, row_headings=False)
    pattern = r'class="sp-freeze" style="left:(\d+)px;top:(\d+)px;width:(\d+)px;height:(\d+)px"'
    lines = re.findall(pattern, page)
    # Two lines in the sheet and in each of its three copies.
    assert len(lines) == 8
    horizontal, vertical = lines[0], lines[1]
    assert horizontal[1] == str(round(2 * 15 * 4 / 3))  # below two default rows
    assert horizontal[3] == "1"
    first_column = re.search(r'<col style="width:(\d+)px">', page).group(1)
    assert vertical[0] == first_column  # after column A
    assert vertical[2] == "1"
    assert "--sheetpeek-freeze: #9b9b9b" in page


def test_a_fragment_draws_only_the_freeze_lines_by_default(tmp_path):
    path = _frozen(tmp_path)
    fragment = _html(path, fragment=True)
    assert "sp-panes" not in fragment.split("</style>")[1]
    assert fragment.count('class="sp-freeze"') == 2
    assert 'class="sp-panes"' in _html(path, fragment=True, freeze="panes")
    assert 'class="sp-panes"' not in _html(path, freeze="static")


def test_the_saved_scroll_position_is_not_followed(tmp_path):
    # Frozen two rows from row 3 of the view: rows 1 to 4 are before the split, all drawn.
    page = _html(_frozen(tmp_path, x=0, y=2, top_left="A5", view_top_left="A3"))
    panes = _panes(page)
    assert set(panes) == {"top"}
    assert panes["top"].count("<tr") == 5  # the headings' row and rows 1 to 4
    assert "r1c1" in page.split('<div class="sp-pane ')[0]


@pytest.mark.parametrize(
    ("x", "y", "top_left", "state"),
    [(0, 9, "A10", "frozen"), (5, 0, "F1", "frozen"), (2400, 1800, "C4", "split")],
    ids=["split below the range", "split past the range", "split pane"],
)
def test_nothing_is_frozen_without_a_split_inside_the_range(tmp_path, x, y, top_left, state):
    page = _html(_frozen(tmp_path, x=x, y=y, top_left=top_left, state=state))
    assert "sp-freeze" not in page


def test_a_frozen_split_pane_is_drawn_as_frozen(tmp_path):
    assert 'class="sp-panes"' in _html(_frozen(tmp_path, state="frozenSplit"))


def test_text_reaching_into_the_frozen_columns_is_in_their_copy(tmp_path):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet["C1"] = "right-aligned text long enough to spill"
    sheet["C1"].alignment = Alignment(horizontal="right")
    sheet["A3"] = 1
    path = save(workbook, tmp_path / "spill.xlsx")
    rewrite(
        path,
        SHEET1,
        lambda text: text.replace(
            '<sheetView workbookViewId="0">',
            '<sheetView workbookViewId="0"><pane xSplit="2" topLeftCell="C1" state="frozen"/>',
            1,
        ),
    )
    panes = _panes(_html(path))
    assert "right-aligned text" in panes["left"]


def test_an_unknown_freeze_is_refused():
    with pytest.raises(ValueError, match="freeze"):
        HTMLBackend(freeze="sticky")


def test_other_outputs_draw_and_report_nothing_for_a_frozen_pane(tmp_path):
    path = _frozen(tmp_path)
    for to in ("text", "markdown", "csv"):
        assert not [n for n in sheetpeek.notices(path, to) if n.code.startswith("view")]


def _with_area(tmp_path, area, rows=5, name="area.xlsx"):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for row in range(1, rows + 1):
        for column in range(1, 5):
            sheet.cell(row, column).value = f"r{row}c{column}"
    path = save(workbook, tmp_path / name)
    if area:
        rewrite(
            path,
            "xl/workbook.xml",
            lambda text: text.replace(
                "</sheets>",
                '</sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'
                f"{area}</definedName></definedNames>",
                1,
            ),
        )
    return path


def test_each_print_area_is_a_block_in_the_file_order(tmp_path):
    path = _with_area(tmp_path, "Sheet!$C$4:$D$5,Sheet!$A$1:$B$2")
    text = sheetpeek.render(path, "text", print_area=True, on_notice="ignore")
    assert text == "r4c3     r4c4\nr5c3     r5c4\n\nr1c1     r1c2\nr2c1     r2c2\n"
    csv = sheetpeek.render(path, "csv", print_area=True, on_notice="ignore")
    assert csv == "r4c3,r4c4\nr5c3,r5c4\n\nr1c1,r1c2\nr2c1,r2c2\n"


def test_print_areas_are_blocks_of_one_html_page(tmp_path):
    path = _with_area(tmp_path, "Sheet!$C$4:$D$5,Sheet!$A$1:$B$2")
    page = sheetpeek.render(path, "html", print_area=True, on_notice="ignore")
    assert page.count("<html") == 1
    assert page.count('<div class="sp-area">') == 2
    assert page.index("r4c3") < page.index("r1c1")


def test_whole_columns_and_rows_are_cut_to_the_used_range(tmp_path):
    path = _with_area(tmp_path, "Sheet!$B:$C")
    assert sheetpeek.render(path, "csv", print_area=True, on_notice="ignore").count("\n") == 5
    path = _with_area(tmp_path, "Sheet!$2:$3", name="rows.xlsx")
    csv = sheetpeek.render(path, "csv", print_area=True, on_notice="ignore")
    assert csv == "r2c1,r2c2,r2c3,r2c4\nr3c1,r3c2,r3c3,r3c4\n"


def test_whole_columns_stay_within_a_small_cell_limit(tmp_path):
    path = _with_area(tmp_path, "Sheet!$A:$D")
    limits = sheetpeek.ReadLimits(max_cells=100)
    assert sheetpeek.render(path, "csv", print_area=True, limits=limits, on_notice="ignore")


def test_a_sheet_without_a_print_area_is_shown_whole(tmp_path):
    path = _with_area(tmp_path, "")
    found = []
    assert sheetpeek.render(path, "csv", print_area=True, on_notice=found.append) == (
        sheetpeek.render(path, "csv", on_notice="ignore")
    )
    assert found == []


def test_the_print_area_is_off_by_default(tmp_path):
    path = _with_area(tmp_path, "Sheet!$A$1:$A$1")
    assert sheetpeek.render(path, "csv", on_notice="ignore").count("\n") == 5


def test_an_unusable_part_is_reported_only_with_the_option(tmp_path):
    path = _with_area(tmp_path, "Sheet!$A$1:$A$2,Sheet!#REF!")
    found = []
    csv = sheetpeek.render(path, "csv", print_area=True, on_notice=found.append)
    assert csv == "r1c1\nr2c1\n"
    assert [(n.code, n.reason) for n in found] == [("print.area_unusable", "file")]
    found.clear()
    sheetpeek.render(path, "csv", on_notice=found.append)
    assert found == []


def test_the_print_area_option_takes_a_workbook_and_no_range(tmp_path):
    path = _with_area(tmp_path, "Sheet!$A$1:$A$2")
    with pytest.raises(ValueError, match="ref"):
        sheetpeek.render(path, "csv", ref="A1:B2", print_area=True)
    with pytest.raises(ValueError, match="workbook"):
        sheetpeek.render(sheetpeek.read(path), "csv", print_area=True)
    with path.open("rb") as handle:
        assert (
            sheetpeek.render(handle, "csv", print_area=True, on_notice="ignore") == "r1c1\nr2c1\n"
        )


def test_the_cli_shows_the_print_area(tmp_path, capsys):
    from sheetpeek.cli import main

    path = str(_with_area(tmp_path, "Sheet!$C$4:$D$5,Sheet!$A$1:$B$2"))
    assert main(["view", path, "-t", "csv", "--print-area", "--notices", "none"]) == 0
    assert capsys.readouterr().out == "r4c3,r4c4\nr5c3,r5c4\n\nr1c1,r1c2\nr2c1,r2c2\n"
    assert main(["view", path, "--print-area", "-r", "A1:B2"]) == 2
    assert main(["view", path, "--print-area", "--all-sheets"]) == 2
    assert "--print-area" in capsys.readouterr().err
