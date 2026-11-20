"""A sheet's structure (the reader's cell stream and xlio's reading), against openpyxl and XML."""

import openpyxl
import pytest
from openpyxl.styles import Alignment

import sheetpeek
from sheetpeek import CellRange
from sheetpeek._notices import Collector
from sheetpeek._structure import ColumnSpec, fill, width_at_digit
from tests.workbooks import raw_sheet, save


def scan_file(path):
    """The first sheet's structure, as the reader finds it before reading its cells."""
    with sheetpeek.open(path) as book:
        source = book._book.sheet(0)
        found, _, _ = book._stream(source, None, None, Collector("sheet"))
        return fill(found, source)


@pytest.fixture
def scan_xml(tmp_path):
    def scan_body(body: str):
        return scan_file(raw_sheet(tmp_path / "raw.xlsx", body))

    return scan_body


def test_matches_openpyxl_normal_mode(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "x"
    ws["F9"] = 3
    ws.merge_cells("A2:C3")
    ws.merge_cells("E5:E8")
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["D"].hidden = True
    ws.column_dimensions.group("G", "I", hidden=True)
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[5].hidden = True
    ws.sheet_format.defaultColWidth = 12
    ws.sheet_view.showGridLines = False
    ws["A2"].alignment = Alignment(horizontal="center")
    path = save(wb, tmp_path / "s.xlsx")

    found = scan_file(path)
    reference = openpyxl.load_workbook(path).active
    assert {str(m) for m in found.merges} == {str(m) for m in reference.merged_cells.ranges}
    assert {(c.min, c.max, c.width, c.hidden) for c in found.columns} == {
        (d.min, d.max, d.width, d.hidden) for d in reference.column_dimensions.values()
    }
    rows = {i: (d.height, bool(d.hidden)) for i, d in reference.row_dimensions.items()}
    assert found.rows == {i: r for i, r in rows.items() if r != (None, False)}
    assert found.default_column_width == 12
    assert found.base_column_width == reference.sheet_format.baseColWidth
    assert found.default_row_height == reference.sheet_format.defaultRowHeight
    assert found.show_gridlines is False
    assert found.used_range == CellRange.parse(reference.dimensions)


def test_used_range_counts_every_cell_element(scan_xml):
    found = scan_xml(
        '<dimension ref="A1"/><sheetData>'
        '<row r="3"><c r="C3" s="0"/></row>'
        '<row r="7"><c r="B7"><v>1</v></c><c r="H7" t="inlineStr"><is><t>x</t></is></c></row>'
        '<row r="9" ht="30" customHeight="1"/>'
        "</sheetData>"
    )
    # A cell element without a value counts; a row with only a height does not.
    assert found.used_range == CellRange.parse("B3:H7")
    assert found.rows == {9: (30.0, False)}


def test_positions_without_references(scan_xml):
    # The schema allows omitting r on rows and cells; positions then follow their predecessors.
    found = scan_xml(
        "<sheetData>"
        "<row><c><v>1</v></c><c><v>2</v></c></row>"
        '<row r="4"><c r="C4"><v>3</v></c><c><v>4</v></c></row>'
        "<row><c><v>5</v></c></row>"
        "</sheetData>"
    )
    assert found.used_range == CellRange.parse("A1:D5")


def test_used_range_includes_merges(scan_xml):
    found = scan_xml(
        "<sheetData><row r='2'><c r='B2'><v>1</v></c></row></sheetData>"
        "<mergeCells count='1'><mergeCell ref='C1:E3'/></mergeCells>"
    )
    assert found.used_range == CellRange.parse("B1:E3")


def test_empty_sheet(scan_xml):
    found = scan_xml("<sheetData/>")
    assert found.used_range is None
    assert found.merges == ()


def test_column_specs(scan_xml):
    found = scan_xml(
        '<cols><col min="2" max="4" width="3.5" customWidth="1"/>'
        '<col min="6" max="6" hidden="true"/></cols><sheetData/>'
    )
    assert found.columns == (ColumnSpec(2, 4, 3.5, False, "3.5"), ColumnSpec(6, 6, None, True))


def test_formulas_returning_the_empty_string(tmp_path):
    # An empty cached string is the empty string (xlio); a t="str" formula without <v> has no
    # cached value, as xlio reads it, and is computed (here the empty string, and 2).
    path = raw_sheet(
        tmp_path / "raw.xlsx",
        "<sheetData><row r='1'>"
        "<c r='A1' t='str'><f>\"\"</f><v></v></c>"
        "<c r='B1' t='str'><f>\"\"</f></c>"
        "<c r='C1'><f>1+1</f><v></v></c>"
        "</row></sheetData>",
    )
    assert scan_file(path).uncached_formulas == 2
    sheet = sheetpeek.read(path)
    assert [(c.value, c.kind) for c in sheet.cells[0]] == [
        ("", "text"),
        ("", "text"),
        (2, "number"),
    ]


def test_formulas_without_cached_values(scan_xml):
    found = scan_xml(
        "<sheetData><row r='1'>"
        "<c r='A1'><f>1+1</f><v>2</v></c>"
        "<c r='B1'><f>1+1</f></c>"
        "<c r='C1' t='e'><f>\"a\"</f></c>"
        "</row></sheetData>"
    )
    assert found.uncached_formulas == 2
    assert found.first_uncached_formula == "B1"


def test_first_sheet_view_wins(scan_xml):
    found = scan_xml(
        '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/>'
        '<sheetView workbookViewId="1"/></sheetViews><sheetData/>'
    )
    assert found.show_gridlines is False


def test_raw_sheet_helper(tmp_path):
    path = raw_sheet(
        tmp_path / "r.xlsx", '<sheetData><row r="2"><c r="B2"><v>1</v></c></row></sheetData>'
    )
    assert scan_file(path).used_range == CellRange.parse("B2")


@pytest.mark.parametrize(
    ("raw", "xlio_width", "points"),
    [
        ("100", 100.0, 700.0),  # under 65,536 points, as written
        ("1e6", 6036.5, 53184.0),  # 7e6 modulo 65,536 (captured package_more, Calibri 14)
        ("-1", 10921.0, 65530.0),  # -7 modulo 65,536, and one more point (captured)
        ("9400", 9400.0, 264.0),  # past 65,536 points at 7 pt, though not at 6
        ("1e30", 0.1, 1.0),  # past 2**31 points: one point (captured)
        ("-1e30", 0.1, 1.0),
        ("nan", 0.0, 0.0),  # xlio's 0, a hidden column, stands
        ("1e308", 0.0, 0.0),
        ("abc", 0.0, 0.0),
    ],
)
def test_a_width_wraps_at_the_normal_fonts_digit(raw, xlio_width, points):
    spec = ColumnSpec(1, 1, xlio_width, False, raw)
    assert width_at_digit(spec, 7.0) * 7.0 == pytest.approx(points)
    assert width_at_digit(spec, 6.0) == xlio_width  # xlio's own digit: its width


def test_a_width_xlio_does_not_give_is_kept():
    assert width_at_digit(ColumnSpec(1, 1, None, True), 7.0) is None
    assert width_at_digit(ColumnSpec(1, 1, 2.0, False, None), 7.0) == 2.0
    # Text xlio took as a number that Python does not parse keeps xlio's width.
    assert width_at_digit(ColumnSpec(1, 1, 5.0, False, "5 pt"), 7.0) == 5.0
