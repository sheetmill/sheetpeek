"""The screen suite ``font_cut`` (where Excel stops honoring a workbook's fonts), without Excel.

Its workbooks are built and their font tables read back from the files; its ink measure runs on
a synthetic capture with known marks.
"""

import re
import zipfile

import openpyxl
import pytest
from capture import font_cut as suite

from tests.screen_canvas import Canvas


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    return suite.build(tmp_path_factory.mktemp("font_cut"))


def _parts(path):
    with zipfile.ZipFile(path) as z:
        return z.read("xl/styles.xml").decode(), z.read("xl/worksheets/sheet1.xml").decode()


def test_each_list_probe_is_the_font_table_entry_its_case_names(built):
    styles, sheet = _parts(built["list"])
    fonts = re.search(r"<fonts[^>]*>(.*?)</fonts>", styles, re.S).group(1)
    assert len(re.findall(r"<font[ >]", fonts)) == suite.LIST_FIRST + suite.LIST_PROBES
    xfs = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", styles, re.S).group(1)
    font_of = [int(f) for f in re.findall(r'<xf [^>]*fontId="(\d+)"', xfs)]
    for cell, entry in suite.probes("list"):
        style = re.search(rf'<c r="{cell}"[^>]* s="(\d+)"', sheet)
        assert font_of[int(style.group(1))] == entry, cell


def test_the_rich_runs_follow_the_list_in_cell_order(built):
    styles, sheet = _parts(built["rich"])
    fonts = re.search(r"<fonts[^>]*>(.*?)</fonts>", styles, re.S).group(1)
    assert len(re.findall(r"<font[ >]", fonts)) == suite.RICH_FIRST
    runs = re.findall(r'<c r="(C\d+)" t="inlineStr"><is><r><rPr>', sheet)
    assert runs == [cell for cell, _ in suite.probes("rich")]
    assert [entry for _, entry in suite.probes("rich")] == list(range(995, 1005))


def test_the_ink_measure_tells_the_probes_red_from_the_default_fonts_black():
    ws = openpyxl.Workbook().active
    suite.fill(ws, "list")
    canvas = Canvas(ws, ["C4", "C5"])
    x0, y0, _, _ = canvas.box("C4")
    canvas.paint(x0 + 6, y0 + 8, x0 + 20, y0 + 30, (230, 0, 0))
    found = suite._ink(canvas.capture(), "C4")
    assert (found["red"], found["dark"]) == (str(14 * 22), "0")
    assert (found["ink_top"], found["ink_bottom"]) == ("4", "15")
    x0, y0, _, _ = canvas.box("C5")
    canvas.paint(x0 + 6, y0 + 10, x0 + 12, y0 + 24, (0, 0, 0))
    found = suite._ink(canvas.capture(), "C5")
    assert (found["red"], found["dark"]) == ("0", str(6 * 14))


def test_the_suite_is_a_screen_suite():
    from capture.screen.__main__ import SUITES

    assert "font_cut" in SUITES
