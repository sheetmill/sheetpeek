"""The machine's recorded window geometry (``capture.screen.geometry``), without Excel."""

import pytest
from capture.screen import geometry
from capture.screen import run as screen_run

FOUND = {
    "screen_pt": [2560, 1440],
    "display_scale": 2,
    "window_pt": [0, 25, 2560, 1320],
    "usable_pt": [2560.0, 1089.0],
    "window_state": "window state maximized",
}


@pytest.fixture
def recorded(tmp_path, monkeypatch):
    path = tmp_path / "geometry.json"
    monkeypatch.setenv("SHEETPEEK_CAPTURE_GEOMETRY", str(path))
    return path


def test_the_file_is_named_by_the_environment_or_lies_in_the_home_folder(recorded, monkeypatch):
    assert geometry.expected_path() == recorded
    monkeypatch.delenv("SHEETPEEK_CAPTURE_GEOMETRY")
    assert geometry.expected_path().name == "geometry.json"
    assert geometry.expected_path().parent.name == ".sheetpeek-capture"


def test_a_capture_refuses_without_a_recorded_geometry(recorded):
    with pytest.raises(geometry.GeometryError, match="geometry record"):
        geometry.check(FOUND)


def test_a_recorded_geometry_is_read_back_and_passes(recorded):
    geometry.record(FOUND)
    geometry.check(FOUND)
    assert geometry.expected() == geometry.normalized(FOUND)


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("screen_pt", [1280, 720]),
        ("display_scale", 1),
        ("window_pt", [0, 25, 2560, 1319]),
        ("usable_pt", [2560.0, 1089.5]),
        ("window_state", "window state normal"),
    ],
)
def test_a_difference_in_any_field_refuses_and_names_it(recorded, field, value):
    geometry.record(FOUND)
    with pytest.raises(geometry.GeometryError, match=field):
        geometry.check({**FOUND, field: value})


def test_numbers_compare_to_hundredths_and_the_header_line_is_compact():
    assert geometry.differences({**FOUND, "display_scale": 2.001}, FOUND) == []
    line = geometry.text(FOUND)
    assert "\n" not in line
    assert " " not in line.replace("window state maximized", "")
    with pytest.raises(geometry.GeometryError, match="lacks"):
        geometry.normalized({"screen_pt": [1, 1]})


class _Session:
    """Stands in for ExcelSession: geometry only; any other request fails the test."""

    def __init__(self, found):
        self.found = found
        self.open_after = ""

    def __call__(self, path, evidence):
        return self

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def geometry(self):
        return self.found

    def view(self, *args):
        raise AssertionError("a view before the geometry was checked")


def test_a_capture_checks_the_geometry_before_its_first_view(recorded, tmp_path, monkeypatch):
    pytest.importorskip("openpyxl")  # the capture builds its workbook with it
    geometry.record(FOUND)
    monkeypatch.setattr(screen_run, "CONTAINER", tmp_path / "container")
    monkeypatch.setenv("EXCEL_EDITION", "Microsoft 365")
    monkeypatch.setattr(screen_run, "ExcelSession", _Session({**FOUND, "display_scale": 1}))
    with pytest.raises(geometry.GeometryError, match="display_scale"):
        screen_run.capture(lambda ws: None, cells=["C4"], out=tmp_path / "out")
    # the recorded geometry passes the check and reaches the view (which this stand-in refuses)
    monkeypatch.setattr(screen_run, "ExcelSession", _Session(FOUND))
    with pytest.raises(AssertionError, match="a view before"):
        screen_run.capture(lambda ws: None, cells=["C4"], out=tmp_path / "out")
