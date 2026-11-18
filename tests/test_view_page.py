"""The sheet view and page layout as read: zoom, zeros, frozen panes, print areas, page breaks.

ARCHITECTURE.md §4.18. These tests check the values the file states, taken as Excel takes them
where that is captured; ``tests/test_view_drawn.py`` tests how they are drawn.
"""

import openpyxl
import pytest

import sheetpeek
from sheetpeek import CellRange, SheetView
from tests.test_limits import SHEET1, book  # noqa: F401 - the fixture
from tests.workbooks import rewrite, save


def _view(book, attributes, pane=""):  # noqa: F811
    rewrite(
        book,
        SHEET1,
        lambda text: text.replace(
            '<sheetView workbookViewId="0">',
            f'<sheetView workbookViewId="0" {attributes}>{pane}',
            1,
        ),
    )
    return sheetpeek.read(book).view


def test_a_sheet_without_view_settings_reads_the_defaults(book):  # noqa: F811
    assert sheetpeek.read(book).view == SheetView()


def test_the_saved_zoom_and_zero_display_are_read(book):  # noqa: F811
    view = _view(book, 'zoomScale="85" showZeros="0" topLeftCell="C5"')
    assert (view.zoom, view.show_zeros, view.top_left) == (85, False, "C5")


@pytest.mark.parametrize("value", ["9", "401", "-100", "0"])
def test_a_zoom_excel_does_not_allow_is_100(book, value):  # noqa: F811
    assert _view(book, f'zoomScale="{value}"').zoom == 100


@pytest.mark.parametrize("value", ["85.5", "abc", ""])
def test_a_zoom_that_is_not_a_whole_number_is_100(book, value):  # noqa: F811
    # As xlio reads it; what Excel does is not captured yet (unverified).
    assert _view(book, f'zoomScale="{value}"').zoom == 100


def test_a_frozen_pane_gives_its_rows_and_columns(book):  # noqa: F811
    view = _view(book, "", '<pane xSplit="2" ySplit="3" topLeftCell="C4" state="frozen"/>')
    assert (view.frozen_rows, view.frozen_columns, view.pane_state) == (3, 2, "frozen")


def test_a_split_pane_freezes_nothing(book):  # noqa: F811
    # A split pane's xSplit and ySplit are positions in twips, not counts.
    view = _view(book, "", '<pane xSplit="2400" ySplit="1800" topLeftCell="C4"/>')
    assert (view.frozen_rows, view.frozen_columns, view.pane_state) == (0, 0, "split")


@pytest.mark.parametrize(
    ("text", "areas", "problem"),
    [
        ("Sheet!$A$1:$C$5", [CellRange(1, 1, 5, 3)], None),
        ("sheet!$B$2", [CellRange(2, 2, 2, 2)], None),
        ("Sheet!$A$1:$B$2,Sheet!$D$4:$E$5", [CellRange(1, 1, 2, 2), CellRange(4, 4, 5, 5)], None),
        ("Sheet!$A:$C", [CellRange(1, 1, 1_048_576, 3)], None),
        ("Sheet!$2:$4", [CellRange(2, 1, 4, 16_384)], None),
        ("Sheet!#REF!", [], "#REF!"),
        ("Other!$A$1:$B$2", [], "another sheet"),
        ("OFFSET(Sheet!$A$1,0,0,5,5)", [], "formula"),
        ("Sheet!$A$1:$B$2,Sheet!#REF!", [CellRange(1, 1, 2, 2)], "#REF!"),
    ],
)
def test_a_print_area_is_read_part_by_part(book, text, areas, problem):  # noqa: F811
    # xlio reads the defined name; its problem is worded by xlio.
    rewrite(
        book,
        "xl/workbook.xml",
        lambda t: t.replace(
            "</sheets>",
            '</sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'
            f"{text}</definedName></definedNames>",
            1,
        ),
    )
    sheet = sheetpeek.read(book)
    assert sheet.print_areas == tuple(areas)
    if problem is None:
        assert sheet.print_area_problem is None
    else:
        assert problem in sheet.print_area_problem


def test_the_print_area_and_manual_breaks_are_read_from_the_file(book):  # noqa: F811
    rewrite(
        book,
        "xl/workbook.xml",
        lambda text: text.replace(
            "</workbook>",
            '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'
            "Sheet!$A$1:$B$4</definedName></definedNames></workbook>",
            1,
        ),
    )
    rewrite(
        book,
        SHEET1,
        lambda text: text.replace(
            "</worksheet>",
            '<rowBreaks count="2" manualBreakCount="1"><brk id="10" max="16383" man="1"/>'
            '<brk id="20" max="16383"/></rowBreaks>'
            '<colBreaks count="1" manualBreakCount="1"><brk id="6" max="1048575" man="1"/>'
            "</colBreaks></worksheet>",
            1,
        ),
    )
    sheet = sheetpeek.read(book)
    assert (sheet.print_areas, sheet.print_area_problem) == ((CellRange(1, 1, 4, 2),), None)
    # id="10" lies between rows 10 and 11; only the manual break is kept.
    assert (sheet.row_breaks, sheet.column_breaks) == ((10,), (6,))


def test_quoted_sheet_names_keep_their_commas_and_quotes(tmp_path):
    workbook = openpyxl.Workbook()
    workbook.active.title = "It's, here"
    workbook.active["A1"] = 1
    workbook.active.print_area = ["A1:B2", "C3"]
    path = save(workbook, tmp_path / "quoted.xlsx")
    sheet = sheetpeek.read(path)
    assert sheet.print_areas == (CellRange(1, 1, 2, 2), CellRange(3, 3, 3, 3))
    assert sheet.print_area_problem is None


def test_breaks_outside_the_sheet_are_dropped(book):  # noqa: F811
    rows = (
        '<rowBreaks count="5" manualBreakCount="5"><brk id="0" man="1"/><brk id="1048576" man="1"/>'
        '<brk id="99999999999" man="1"/><brk id="-3" man="1"/><brk id="5" man="1"/></rowBreaks>'
        '<colBreaks count="1" manualBreakCount="1"><brk id="16384" man="1"/></colBreaks>'
    )
    rewrite(book, SHEET1, lambda t: t.replace("</worksheet>", rows + "</worksheet>", 1))
    sheet = sheetpeek.read(book)
    assert (sheet.row_breaks, sheet.column_breaks) == ((5,), ())


def test_a_print_area_and_page_breaks_are_not_reported(book):  # noqa: F811
    # Excel's Normal view draws neither, so the sheet is shown as Excel shows it: no notice,
    # and --strict passes.
    from sheetpeek.cli import main

    rewrite(
        book,
        "xl/workbook.xml",
        lambda text: text.replace(
            "</workbook>",
            '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'
            "Sheet!$A$1:$B$4</definedName></definedNames></workbook>",
            1,
        ),
    )
    rewrite(
        book,
        SHEET1,
        lambda text: text.replace(
            "</worksheet>",
            '<rowBreaks count="1" manualBreakCount="1"><brk id="10" man="1"/></rowBreaks>'
            '<colBreaks count="1" manualBreakCount="1"><brk id="6" man="1"/></colBreaks>'
            "</worksheet>",
            1,
        ),
    )
    sheet = sheetpeek.read(book)
    assert sheet.print_areas
    assert not [n for n in sheet.notices if n.code.startswith("print.") or "element" in n.code]
    assert main(["view", str(book), "--to", "text", "--strict", "--notices", "none"]) == 0


@pytest.mark.parametrize("state", ["frozen", "frozenSplit"])
def test_a_frozen_pane_is_not_reported(book, state):  # noqa: F811
    _view(book, "", f'<pane xSplit="2" ySplit="3" topLeftCell="C4" state="{state}"/>')
    assert "view.panes" not in {n.code for n in sheetpeek.read(book).notices}


def test_a_split_pane_is_reported(book):  # noqa: F811
    _view(book, "", '<pane xSplit="2400" ySplit="1800" topLeftCell="C4"/>')
    assert "view.panes" in {n.code for n in sheetpeek.read(book).notices}


def test_print_titles_are_not_reported(book):  # noqa: F811
    # Excel's Normal view draws nothing for print titles, as for a print area.
    rewrite(
        book,
        "xl/workbook.xml",
        lambda text: text.replace(
            "</sheets>",
            '</sheets><definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">'
            "Sheet!$1:$1</definedName></definedNames>",
            1,
        ),
    )
    assert sheetpeek.read(book).notices == ()


@pytest.mark.parametrize("show_zeros", [True, False])
def test_a_sheet_that_hides_zeros_shows_them_as_excel_does(tmp_path, show_zeros):
    # xlformat's capture "a sheet that hides zeros": a value equal to 0 under a format without a
    # zero section shows nothing; a zero section still shows; only the stored value counts. The
    # hidden zero's cell still blocks spill (unverified).
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.sheet_view.showZeros = show_zeros
    sheet["A1"] = 0
    sheet["B1"] = 0
    sheet["B1"].number_format = '0;-0;"zero"'
    sheet["C1"] = 0.4
    sheet["C1"].number_format = "0"
    sheet["D1"] = "0"
    sheet["A2"] = "a long text to spill"
    sheet["B2"] = 0
    path = save(workbook, tmp_path / "zeros.xlsx")
    read = sheetpeek.read(path)
    assert read.view.show_zeros is show_zeros
    laid = read.layout("excel")
    first, second = laid.rows
    zero = "0" if show_zeros else ""
    assert [cell.text for cell in first.cells] == [zero, "zero", "0", "0"]
    assert second.cells[1].text == zero
    assert second.runs[0].text == "a long text to spill"
    assert second.runs[0].clip[2] <= (laid.columns[1].x_pt or 0) + 0.01  # stops at B2
    assert not {n.code for n in read.notices} & {"view.hide_zeros"}
