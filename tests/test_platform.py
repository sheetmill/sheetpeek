"""``platform=`` and ``regional_settings=``: resolved once when reading, carried by the sheet.

The Windows values asserted here are xlformat's (its published Windows tables at 96 dpi): a
default column of 64 px, built-in format 14 as ``m/d/yyyy``, text measured in whole pixels.
"""

import datetime
import io

import pytest
import xlformat

import sheetpeek
from sheetpeek import _platform

openpyxl = pytest.importorskip("openpyxl")


@pytest.fixture
def workbook(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = datetime.date(2026, 11, 23)
    ws["A1"].number_format = "mm-dd-yy"  # built-in format 14, written without its code
    ws["B1"] = "Hello world"
    path = tmp_path / "platform.xlsx"
    wb.save(path)
    return path


def test_a_name_or_an_object_resolves_to_xlformats_platform():
    assert _platform.resolve(None) is xlformat.Platform.WINDOWS  # the default
    assert _platform.resolve("mac") is xlformat.Platform.MAC
    assert _platform.resolve("windows") is xlformat.Platform.WINDOWS
    own = xlformat.Platform("windows", dpi=96)
    assert _platform.resolve(own) is own
    with pytest.raises(ValueError, match="'mac' or 'windows'"):
        _platform.resolve("linux")
    with pytest.raises(TypeError, match=r"xlformat\.Platform"):
        _platform.resolve(96)  # type: ignore[arg-type]


def test_regional_settings_follow_the_platform_unless_given():
    windows = xlformat.Platform.WINDOWS
    assert _platform.regional(None, windows) is xlformat.RegionalSettings.EN_US_WINDOWS
    assert _platform.regional(None, xlformat.Platform.MAC) is xlformat.RegionalSettings.EN_US_MAC
    given = xlformat.RegionalSettings.EN_US_MAC
    assert _platform.regional(given, windows) is given
    with pytest.raises(TypeError, match="RegionalSettings"):
        _platform.regional("en-US", windows)  # type: ignore[arg-type]


def test_platform_mac_reads_as_excel_for_mac(workbook):
    sheet = sheetpeek.read(workbook, platform="mac")
    assert sheet.platform is xlformat.Platform.MAC
    assert sheet.regional_settings is xlformat.RegionalSettings.EN_US_MAC
    assert sheet.cell(1, 1).number_format == "m/d/yy"
    assert [c.width_pt for c in sheet.layout("excel").columns[:2]] == [53.0, 53.0]


def test_a_sheet_read_for_windows_is_laid_out_as_windows_shows_it(workbook):
    sheet = sheetpeek.read(workbook, platform="windows")
    assert sheet.platform is xlformat.Platform.WINDOWS
    assert sheet.regional_settings is xlformat.RegionalSettings.EN_US_WINDOWS
    assert sheet.cell(1, 1).number_format == "m/d/yyyy"  # the regional short date
    rendered = sheet.layout("excel")
    assert [c.width_pt for c in rendered.columns[:2]] == [48.0, 48.0]  # 64 px at 96 dpi
    mac = sheetpeek.read(workbook, platform="mac").layout("excel")
    width = {r.text: r.width for r in rendered.rows[0].runs}["Hello world"]
    assert width != {r.text: r.width for r in mac.rows[0].runs}["Hello world"]
    assert width * 4 / 3 == round(width * 4 / 3)  # whole pixels
    assert _platform.current() is _platform.DEFAULT_PLATFORM  # the layout's platform ends


def test_open_takes_the_platform_for_every_sheet(workbook):
    with sheetpeek.open(workbook, platform=xlformat.Platform.WINDOWS) as book:
        assert book.sheet().platform is xlformat.Platform.WINDOWS


def test_render_and_notices_read_with_the_platform(workbook):
    mac = sheetpeek.render(workbook, "text", platform="mac", on_notice="ignore")
    windows = sheetpeek.render(workbook, "text", on_notice="ignore")  # the default
    assert mac != windows
    assert windows == sheetpeek.render(workbook, "text", platform="windows", on_notice="ignore")
    assert sheetpeek.notices(workbook, "text", platform="windows") is not None


def test_a_sheet_keeps_the_platform_it_was_read_with(workbook):
    sheet = sheetpeek.read(workbook, platform="windows")
    assert sheetpeek.render(sheet, "text", platform="windows", on_notice="ignore")
    for source in (sheet, sheet.layout("characters")):
        with pytest.raises(ValueError, match="read with"):
            sheetpeek.render(source, "text", platform="mac")
        with pytest.raises(ValueError, match="read with"):
            sheetpeek.notices(source, "text", regional_settings=xlformat.RegionalSettings())


def test_an_open_book_keeps_the_platform_it_was_opened_with(workbook):
    with sheetpeek.open(workbook, platform="mac") as book:
        with pytest.raises(ValueError, match="opened with"):
            sheetpeek.render_workbook(book, platform="windows")
        assert sheetpeek.render_workbook(book, platform="mac", on_notice="ignore")
    page = sheetpeek.render_workbook(workbook, platform="windows", on_notice="ignore")
    assert "m/d/yyyy" not in page  # the format is applied, not shown


def test_the_cli_takes_a_platform(workbook):
    from sheetpeek.cli import main

    out = io.StringIO()
    import contextlib

    with contextlib.redirect_stdout(out):
        assert main(["view", str(workbook), "--platform", "windows", "--notices", "none"]) == 0
    assert "#" in out.getvalue().splitlines()[0]
