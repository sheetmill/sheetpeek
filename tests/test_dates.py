"""ISO 8601 text as date serials (``t="d"`` cells)."""

from __future__ import annotations

import pytest

import sheetpeek
from sheetpeek._dates import iso_serial
from tests.workbooks import raw_sheet


@pytest.mark.parametrize(
    ("text", "serial"),
    [
        ("1900-01-01", 1),
        ("1900-02-28", 59),
        ("1900-02-29", 60),  # the 1900 system's leap day, which the calendar lacks
        ("1900-03-01", 61),
        ("2024-01-02", 45293),
        ("1899-12-30", 0),
        ("2024-01-02T12:00:00Z", 45293.5),
        ("2024-01-02T06:00", 45293.25),
        ("T18:00", 0.75),
        ("12:00:00.500", (43200.5) / 86400),
        ("PT36H", 1.5),
        ("PT1H30M", 0.0625),
    ],
)
def test_the_1900_system(text, serial):
    assert iso_serial(text) == pytest.approx(serial, abs=1e-12)


@pytest.mark.parametrize(
    ("text", "serial"),
    [("1904-01-01", 0), ("1904-01-02", 1), ("2024-01-02", 45293 - 1462), ("T06:00", 0.25)],
)
def test_the_1904_system(text, serial):
    assert iso_serial(text, date1904=True) == pytest.approx(serial, abs=1e-12)


@pytest.mark.parametrize(
    "text",
    [
        "",
        "T",
        "x",
        "2024-02-30",
        "2023-02-29",
        "2024-13-01",
        "24:00",
        "2024-01-02T10:61",
        "٢٠٢٤-01-02",  # Arabic-Indic digits: Excel's form is ASCII
    ],
)
def test_text_that_is_no_date_is_refused(text):
    with pytest.raises(ValueError, match="ISO 8601"):
        iso_serial(text)


def test_1900_02_29_does_not_exist_in_the_1904_system():
    with pytest.raises(ValueError, match="ISO 8601"):
        iso_serial("1900-02-29", date1904=True)


@pytest.mark.parametrize(
    "text",
    [
        "1900-01-01",
        "1900-02-28",
        "1900-03-01",
        "1899-12-29T12:00",
        "1850-06-15",
        "2024-01-02T23:59:59.999Z",
        "9999-12-31",
        "T00:00:01",
        "PT2H3M4.5S",
    ],
)
@pytest.mark.parametrize("date1904", [False, True])
def test_the_serials_agree_with_openpyxl(text, date1904):
    # openpyxl's conversion, written independently, gives the same serials.
    dates = pytest.importorskip("openpyxl.utils.datetime")
    epoch = dates.CALENDAR_MAC_1904 if date1904 else dates.CALENDAR_WINDOWS_1900
    expected = dates.to_excel(dates.from_ISO8601(text), epoch)
    assert iso_serial(text, date1904=date1904) == pytest.approx(expected, abs=1e-12)


def test_a_date_cell_reads_as_its_serial(tmp_path):
    path = raw_sheet(
        tmp_path / "dates.xlsx",
        "<sheetData><row r='1'>"
        "<c r='A1' t='d'><v>2024-01-02</v></c>"
        "<c r='B1' t='d'><v>1900-02-29</v></c>"
        "</row></sheetData>",
    )
    sheet = sheetpeek.read(path)
    assert [sheet.cell(1, c).value for c in (1, 2)] == [45293, 60]
    assert sheet.cell(1, 1).kind == "number"
