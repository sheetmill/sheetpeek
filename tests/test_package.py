import zipfile

import openpyxl
import pytest

from sheetpeek._package import relationships, worksheet_parts
from tests.workbooks import replace_once, rewrite, save


@pytest.fixture
def two_sheets(tmp_path):
    wb = openpyxl.Workbook()
    wb.active.title = "First"
    wb.create_sheet("Second sheet")
    return save(wb, tmp_path / "two.xlsx")


def test_worksheet_parts(two_sheets):
    with zipfile.ZipFile(two_sheets) as archive:
        assert worksheet_parts(archive) == {
            "First": "xl/worksheets/sheet1.xml",
            "Second sheet": "xl/worksheets/sheet2.xml",
        }


def test_relative_targets(two_sheets):
    # openpyxl writes absolute targets; Excel writes them relative to the workbook part.
    rels = "xl/_rels/workbook.xml.rels"
    rewrite(two_sheets, rels, replace_once("/xl/worksheets/sheet1.xml", "worksheets/sheet1.xml"))
    rewrite(
        two_sheets, rels, replace_once("/xl/worksheets/sheet2.xml", "x/../worksheets/sheet2.xml")
    )
    with zipfile.ZipFile(two_sheets) as archive:
        assert worksheet_parts(archive) == {
            "First": "xl/worksheets/sheet1.xml",
            "Second sheet": "xl/worksheets/sheet2.xml",
        }


def test_part_without_relationships(two_sheets):
    with zipfile.ZipFile(two_sheets) as archive:
        assert relationships(archive, "xl/worksheets/sheet1.xml") == {}
