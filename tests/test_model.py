import pytest

from sheetpeek import CellRange


@pytest.mark.parametrize(
    ("ref", "expected"),
    [
        ("A1", CellRange(1, 1, 1, 1)),
        ("B2:D5", CellRange(2, 2, 5, 4)),
        ("D5:B2", CellRange(2, 2, 5, 4)),
        ("$A$1:$C$3", CellRange(1, 1, 3, 3)),
        ("aa10", CellRange(10, 27, 10, 27)),
        ("XFD1048576", CellRange(1048576, 16384, 1048576, 16384)),
    ],
)
def test_parse(ref, expected):
    assert CellRange.parse(ref) == expected


@pytest.mark.parametrize("ref", ["", "A", "1", "A0", "A1:", "XFE1", "A1048577", "A1:B2:C3", "R1C1"])
def test_parse_rejects(ref):
    with pytest.raises(ValueError, match=r"not a cell reference|outside the sheet"):
        CellRange.parse(ref)


def test_str_round_trips():
    assert str(CellRange.parse("B2:AA10")) == "B2:AA10"
    assert str(CellRange.parse("C3")) == "C3"


def test_shape_membership_and_intersection():
    area = CellRange.parse("B2:D5")
    assert (area.rows, area.columns) == (4, 3)
    assert (2, 2) in area
    assert (5, 4) in area
    assert (1, 2) not in area
    assert (2, 5) not in area
    assert "B2" not in area
    assert area.intersects(CellRange.parse("D5:F9"))
    assert not area.intersects(CellRange.parse("E1:F9"))


def test_invalid_range():
    with pytest.raises(ValueError, match="not a cell range"):
        CellRange(2, 1, 1, 1)
