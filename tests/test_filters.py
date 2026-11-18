"""Filter and sort buttons: what the reader keeps and which header cells draw which glyph.

ARCHITECTURE.md §4.19. The screen suite ``filters`` is to capture these rules; until then
they are unverified.
"""

from __future__ import annotations

import pytest

import sheetpeek
from sheetpeek import AutoFilter, CellRange, FilterColumn, SortCondition
from tests.workbooks import replace_once, rewrite, save

openpyxl = pytest.importorskip("openpyxl")
filters = pytest.importorskip("openpyxl.worksheet.filters")
tables = pytest.importorskip("openpyxl.worksheet.table")

SHEET = "xl/worksheets/sheet1.xml"


def grid(ws, rows=4, columns=4, top=1, left=1):
    for r in range(top, top + rows):
        for c in range(left, left + columns):
            ws.cell(r, c, f"h{c}" if r == top else r * c)


def buttons(rendered):
    return {
        (cell.row, cell.column): cell.button
        for row in rendered.rows
        for cell in row.cells
        if cell.button is not None
    }


def book(tmp_path, build):
    wb = openpyxl.Workbook()
    ws = wb.active
    grid(ws)
    build(ws)
    return save(wb, tmp_path / "filters.xlsx")


def test_every_header_column_draws_a_plain_button(tmp_path):
    path = book(tmp_path, lambda ws: setattr(ws.auto_filter, "ref", "A1:D4"))
    sheet = sheetpeek.read(path)
    assert sheet.autofilters == (AutoFilter(CellRange.parse("A1:D4")),)
    for mode in ("excel", "characters", "none"):
        found = buttons(sheet.layout(mode))
        assert found == {(1, c): "plain" for c in range(1, 5)}


def test_criteria_hidden_buttons_and_sorts(tmp_path):
    def build(ws):
        ws.auto_filter.ref = "A1:D4"
        columns = ws.auto_filter.filterColumn
        columns.append(filters.FilterColumn(colId=0, filters=filters.Filters(filter=["2"])))
        columns.append(filters.FilterColumn(colId=1, hiddenButton=True))
        columns.append(filters.FilterColumn(colId=2, showButton=False))
        columns.append(filters.FilterColumn(colId=9, filters=filters.Filters(filter=["x"])))
        ws.auto_filter.sortState = filters.SortState(
            ref="A2:D4",
            sortCondition=[
                filters.SortCondition(ref="A2:A4", descending=True),
                filters.SortCondition(ref="D2:D4"),
            ],
        )

    sheet = sheetpeek.read(book(tmp_path, build))
    (autofilter,) = sheet.autofilters
    # colId 9 lies past the range's four columns: left out.
    assert autofilter.columns == (
        FilterColumn(1, filtered=True),
        FilterColumn(2, button=False),
        FilterColumn(3, button=False),
    )
    assert autofilter.sort_conditions == (SortCondition(1, True), SortCondition(4))
    # hiddenButton and showButton="0" draw alike: no button (captured).
    assert buttons(sheet.layout("excel")) == {(1, 1): "filtered_descending", (1, 4): "ascending"}


@pytest.mark.parametrize(
    ("criteria", "expected"),
    [
        ({"filters": filters.Filters(blank=True)}, "filtered"),
        (
            {
                "customFilters": filters.CustomFilters(
                    customFilter=[filters.CustomFilter(operator="greaterThan", val="3")]
                )
            },
            "filtered",
        ),
        ({"top10": filters.Top10(val=1)}, "filtered"),
        ({"dynamicFilter": filters.DynamicFilter(type="aboveAverage")}, "filtered"),
        ({"colorFilter": filters.ColorFilter(dxfId=0)}, "filtered"),
        ({"iconFilter": filters.IconFilter(iconSet="3Arrows", iconId=0)}, "filtered"),
        ({}, "plain"),
    ],
)
def test_every_kind_of_criteria_draws_the_funnel(tmp_path, criteria, expected):
    def build(ws):
        ws.auto_filter.ref = "A1:D4"
        ws.auto_filter.filterColumn.append(filters.FilterColumn(colId=1, **criteria))

    rendered = sheetpeek.read(book(tmp_path, build)).layout("excel")
    assert rendered.cell(1, 2).button == expected


def test_a_sort_beside_the_filter_and_a_left_to_right_sort(tmp_path):
    def build(ws):
        ws.auto_filter.ref = "A1:D4"

    path = book(tmp_path, build)
    beside = '<sortState ref="A2:D4"><sortCondition descending="1" ref="B2:B4"/></sortState>'
    rewrite(
        path,
        SHEET,
        replace_once('<autoFilter ref="A1:D4" />', f'<autoFilter ref="A1:D4" />{beside}'),
    )
    assert sheetpeek.read(path).layout("excel").cell(1, 2).button == "descending"  # unverified
    rewrite(
        path,
        SHEET,
        replace_once('<sortState ref="A2:D4">', '<sortState ref="A2:D4" columnSort="1">'),
    )
    assert sheetpeek.read(path).layout("excel").cell(1, 2).button == "plain"  # unverified


def test_no_button_on_a_hidden_column_or_header_row(tmp_path):
    def build(ws):
        ws.auto_filter.ref = "A1:D4"
        ws.column_dimensions["B"].hidden = True

    path = book(tmp_path, build)
    sheet = sheetpeek.read(path)
    assert set(buttons(sheet.layout("excel"))) == {(1, 1), (1, 3), (1, 4)}
    assert (1, 2) not in buttons(sheet.layout("characters", show_hidden=True))

    def hide_header(ws):
        ws.auto_filter.ref = "A1:D4"
        ws.row_dimensions[1].hidden = True

    assert buttons(sheetpeek.read(book(tmp_path, hide_header)).layout("excel")) == {}


def test_a_range_read_keeps_the_header_cells_inside_it(tmp_path):
    path = book(tmp_path, lambda ws: setattr(ws.auto_filter, "ref", "A1:D4"))
    assert set(buttons(sheetpeek.read(path, ref="B1:C3").layout("excel"))) == {(1, 2), (1, 3)}
    # Without the header row: no buttons; a range that misses the filter: no filter at all.
    assert buttons(sheetpeek.read(path, ref="A2:D4").layout("excel")) == {}
    assert sheetpeek.read(path, ref="F6:G7").autofilters == ()


def test_merged_and_empty_header_cells_draw_a_button_each(tmp_path):
    def build(ws):
        ws.auto_filter.ref = "A1:D4"
        ws["C1"] = None
        ws.merge_cells("A1:B1")

    rendered = sheetpeek.read(book(tmp_path, build)).layout("excel")
    assert set(buttons(rendered)) == {(1, 1), (1, 2), (1, 3), (1, 4)}


def test_a_tables_filter_draws_buttons(tmp_path):
    def build(ws):
        grid(ws, rows=3, columns=2, top=1, left=6)
        table = tables.Table(displayName="Sales", ref="F1:G3")
        ws.add_table(table)

    path = book(tmp_path, build)
    sheet = sheetpeek.read(path)
    assert sheet.autofilters == (AutoFilter(CellRange.parse("F1:G3"), "Sales"),)
    assert buttons(sheet.layout("excel")) == {(1, 6): "plain", (1, 7): "plain"}


def test_a_table_without_a_filter_or_header_row_draws_none(tmp_path):
    def build(ws):
        grid(ws, rows=3, columns=2, top=1, left=6)
        ws.add_table(tables.Table(displayName="Sales", ref="F1:G3"))

    path = book(tmp_path, build)
    rewrite(
        path, "xl/tables/table1.xml", lambda text: text.replace('<autoFilter ref="F1:G3" />', "")
    )
    assert sheetpeek.read(path).autofilters == ()
    path = book(tmp_path, build)
    rewrite(path, "xl/tables/table1.xml", replace_once('headerRowCount="1"', 'headerRowCount="0"'))
    assert sheetpeek.read(path).autofilters == ()


def test_the_census_counts_the_buttons(tmp_path):
    path = book(tmp_path, lambda ws: setattr(ws.auto_filter, "ref", "A1:D4"))
    assert sheetpeek.read(path).layout("excel").features["filter.buttons"].count == 4
