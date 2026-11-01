"""The structure scan, checked against openpyxl's normal mode and against hand-written XML."""

import io
import zipfile

import openpyxl
from openpyxl.styles import Alignment

from sheetpeek import CellRange
from sheetpeek._structure import ColumnSpec, scan
from tests.workbooks import SHEET1, raw_sheet, save


def scan_file(path, rels=None):
    with zipfile.ZipFile(path) as archive, archive.open(SHEET1) as handle:
        return scan(handle, rels)


def scan_xml(body: str, rels=None):
    xml = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"{body}</worksheet>"
    )
    return scan(io.BytesIO(xml.encode()), rels)


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


def test_used_range_counts_every_cell_element():
    found = scan_xml(
        '<dimension ref="A1"/><sheetData>'
        '<row r="3"><c r="C3" s="1"/></row>'
        '<row r="7"><c r="B7"><v>1</v></c><c r="H7" t="s"><v>0</v></c></row>'
        '<row r="9" ht="30" customHeight="1"/>'
        "</sheetData>"
    )
    # A styled cell without a value counts; a row with only a height does not.
    assert found.used_range == CellRange.parse("B3:H7")
    assert found.rows == {9: (30.0, False)}


def test_positions_without_references():
    # The schema allows omitting r on rows and cells; positions then follow their predecessors.
    found = scan_xml(
        "<sheetData>"
        "<row><c><v>1</v></c><c><v>2</v></c></row>"
        '<row r="4"><c r="C4"><v>3</v></c><c><v>4</v></c></row>'
        "<row><c><v>5</v></c></row>"
        "</sheetData>"
    )
    assert found.used_range == CellRange.parse("A1:D5")


def test_used_range_includes_merges():
    found = scan_xml(
        "<sheetData><row r='2'><c r='B2'><v>1</v></c></row></sheetData>"
        "<mergeCells count='1'><mergeCell ref='C1:E3'/></mergeCells>"
    )
    assert found.used_range == CellRange.parse("B1:E3")


def test_empty_sheet():
    found = scan_xml("<sheetData/>")
    assert found.used_range is None
    assert found.merges == ()


def test_column_specs():
    found = scan_xml(
        '<cols><col min="2" max="4" width="3.5" customWidth="1"/>'
        '<col min="6" max="6" hidden="true"/></cols><sheetData/>'
    )
    assert found.columns == (ColumnSpec(2, 4, 3.5, False), ColumnSpec(6, 6, None, True))


def test_formulas_returning_the_empty_string():
    found = scan_xml(
        "<sheetData><row r='1'>"
        "<c r='A1' t='str'><f>\"\"</f><v></v></c>"
        "<c r='B1' t='str'><f>\"\"</f></c>"
        "<c r='C1'><f>1+1</f><v></v></c>"
        "</row></sheetData>"
    )
    assert found.empty_strings == {(1, 1), (1, 2)}
    assert found.uncached_formulas == 1


def test_formulas_without_cached_values():
    found = scan_xml(
        "<sheetData><row r='1'>"
        "<c r='A1'><f>1+1</f><v>2</v></c>"
        "<c r='B1'><f>1+1</f></c>"
        "<c r='C1' t='e'><f>\"a\"</f></c>"
        "</row></sheetData>"
    )
    assert found.uncached_formulas == 2
    assert found.first_uncached_formula == "B1"


def test_hyperlinks():
    found = scan_xml(
        "<sheetData/><hyperlinks>"
        '<hyperlink ref="A1" r:id="rId1"/>'
        '<hyperlink ref="B2:B3" location="Other!A1"/>'
        '<hyperlink ref="C1" r:id="rId9"/>'
        "</hyperlinks>",
        {"rId1": ("https://example.com/", True)},
    )
    # An unknown relationship id and no location: no link.
    assert found.hyperlinks == (
        (CellRange.parse("A1"), "https://example.com/"),
        (CellRange.parse("B2:B3"), "#Other!A1"),
    )


def test_first_sheet_view_wins():
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
