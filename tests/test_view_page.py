"""The sheet view and page layout as read: zoom, zeros, frozen panes, print areas, page breaks.

ARCHITECTURE.md §4.18. Nothing here is drawn yet; the values are what the file states, read as
Excel reads them where that is captured.
"""

import pytest

import sheetpeek
from sheetpeek import CellRange, SheetView
from sheetpeek._page import print_areas
from tests.test_limits import SHEET1, book  # noqa: F401 - the fixture
from tests.workbooks import rewrite


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
def test_a_zoom_that_is_not_a_whole_number_refuses_the_workbook(book, value):  # noqa: F811
    # openpyxl refuses it (measured, 3.1.5); what Excel does is not captured.
    with pytest.raises(sheetpeek.PackageError):
        _view(book, f'zoomScale="{value}"')


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
        ("Sheet!$A:$C", [CellRange(1, 1, 1048576, 3)], None),
        ("Sheet!$2:$4", [CellRange(2, 1, 4, 16384)], None),
        ("Sheet!#REF!", [], "a deleted range (#REF!)"),
        ("#REF!", [], "a deleted range (#REF!)"),
        ("Other!$A$1:$B$2", [], "a range on another sheet"),
        ("OFFSET(Sheet!$A$1,0,0,5,5)", [], "a formula"),
        ("Sheet!$A$1:$B$2,Sheet!#REF!", [CellRange(1, 1, 2, 2)], "a deleted range (#REF!)"),
        ("Sheet!$A$1:$XFE$2", [], "a formula"),
    ],
)
def test_a_print_area_is_parsed_part_by_part(text, areas, problem):
    assert print_areas(text, "Sheet") == (tuple(areas), problem)


def test_quoted_sheet_names_keep_their_commas_and_quotes():
    text = "'It''s, here'!$A$1:$B$2,'It''s, here'!$C$3"
    assert print_areas(text, "It's, here") == ((CellRange(1, 1, 2, 2), CellRange(3, 3, 3, 3)), None)


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


def test_breaks_outside_the_sheet_are_dropped():
    import xml.etree.ElementTree as ET

    from sheetpeek._page import page_breaks

    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rows = ET.fromstring(
        f'<rowBreaks xmlns="{ns}"><brk id="0" man="1"/><brk id="1048576" man="1"/>'
        '<brk id="99999999999" man="1"/><brk id="-3" man="1"/><brk id="5" man="1"/></rowBreaks>'
    )
    columns = ET.fromstring(f'<colBreaks xmlns="{ns}"><brk id="16384" man="1"/></colBreaks>')
    assert page_breaks([rows, columns]) == ((5,), ())


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
