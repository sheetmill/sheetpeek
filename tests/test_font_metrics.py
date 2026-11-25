"""A face xlformat measures for some functions and not others (ARCHITECTURE.md §6.2).

xlformat measures the widths of more families than its row, line and inset tables cover. Each
function that raises ``LookupError`` for a face is answered as Calibri at the same size, with one
``font.no_metrics`` notice naming the face and the function; nothing raises and no file is
refused. Each test makes one xlformat function raise for Arial, a face every function measures,
and compares the layout with the same workbook in Calibri: take the fallback away and the
patched call raises, or the values differ.
"""

from __future__ import annotations

import functools
import io
from typing import Any

import pytest
import xlformat

import sheetpeek
from sheetpeek import _excel, _platform, _read, _rows, _text_styles
from tests.watchdog import in_child

openpyxl = pytest.importorskip("openpyxl")
from openpyxl.styles import Alignment, Font  # noqa: E402

CACHED = (
    _platform.covers,
    _rows._face,
    _excel._platform_row_model,
    _text_styles._platform_rules,
)


@pytest.fixture
def clear_caches():
    for cached in CACHED:
        cached.cache_clear()
    yield
    for cached in CACHED:
        cached.cache_clear()


def refuse(monkeypatch: pytest.MonkeyPatch, name: str, face: str = "Arial") -> None:
    """Make ``xlformat.<name>`` raise ``LookupError`` for ``face``, as for a face it lacks."""
    original = getattr(xlformat, name)

    @functools.wraps(original)
    def refusing(*args: Any, **kwargs: Any) -> Any:
        fonts = [a for a in (*args, *kwargs.values()) if isinstance(a, xlformat.Font)]
        if any(f.name == face for f in fonts):
            raise LookupError(f"no {name} for font {face!r}")
        return original(*args, **kwargs)

    monkeypatch.setattr(xlformat, name, refusing)
    for cached in CACHED:
        cached.cache_clear()


def workbook(face: str, size: float = 14) -> io.BytesIO:
    """Text in ``face``: plain, wrapped, underlined, struck out and accounting-underlined."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.column_dimensions["A"].width = 30
    cells = (
        ("A1", Font(name=face, size=size), None),
        ("A2", Font(name=face, size=size), Alignment(wrap_text=True)),
        ("A3", Font(name=face, size=size, underline="single", strike=True), None),
        ("A4", Font(name=face, size=size, underline="doubleAccounting"), None),
        ("A5", Font(name=face, size=size), Alignment(horizontal="right")),
    )
    for ref, font, alignment in cells:
        ws[ref] = "Hxg\none\ntwo" if ref == "A2" else "Hxg"
        ws[ref].font = font
        if alignment is not None:
            ws[ref].alignment = alignment
    for row in range(1, 6):
        ws.row_dimensions[row].height = None
    data = io.BytesIO()
    wb.save(data)
    data.seek(0)
    return data


def lay_out(face: str, platform: str) -> sheetpeek.RenderedSheet:
    sheet = sheetpeek.read(workbook(face), platform=platform)
    rendered = sheet.layout("excel")
    sheetpeek.render(rendered, "html")
    return rendered


def runs(rendered: sheetpeek.RenderedSheet) -> list[tuple[float, float, float, float]]:
    return [(r.height_pt, t.x, t.y - r.y_pt, t.width) for r in rendered.rows for t in r.runs]


def rules(rendered: sheetpeek.RenderedSheet) -> list[tuple[float, float]]:
    """Each line's offset from its run's baseline, and its thickness."""
    return [(u.y - t.y, u.thickness) for r in rendered.rows for t in r.runs for u in t.rules]


def no_metrics(rendered: sheetpeek.RenderedSheet) -> list[dict[str, Any]]:
    return [dict(n.context) for n in rendered.notices if n.code == "font.no_metrics"]


@pytest.mark.usefixtures("clear_caches")
@pytest.mark.parametrize(
    ("function", "platform"),
    [("row_metrics", "mac"), ("vertical_metrics", "windows")],
)
def test_a_face_without_rows_is_fitted_as_calibri(monkeypatch, function, platform):
    calibri = lay_out("Calibri", platform)
    refuse(monkeypatch, "row_metrics")  # xlformat covers the same faces in both
    refuse(monkeypatch, "vertical_metrics")
    rendered = lay_out("Arial", platform)
    assert no_metrics(rendered) == [{"font": "Arial", "function": function}]
    # rows and baselines as Calibri's at 14 points; widths and insets stay Arial's
    assert [r.height_pt for r in rendered.rows] == [r.height_pt for r in calibri.rows]
    assert [y for _, _, y, _ in runs(rendered)] == [y for _, _, y, _ in runs(calibri)]
    assert [w for *_, w in runs(rendered)] != [w for *_, w in runs(calibri)]


@pytest.mark.usefixtures("clear_caches")
@pytest.mark.parametrize("platform", ["mac", "windows"])
def test_a_face_without_lines_draws_calibris(monkeypatch, platform):
    calibri = lay_out("Calibri", platform)
    refuse(monkeypatch, "rule_metrics")
    rendered = lay_out("Arial", platform)
    assert no_metrics(rendered) == [{"font": "Arial", "function": "rule_metrics"}]
    assert rules(rendered) == rules(calibri)
    assert rules(rendered)


@pytest.mark.usefixtures("clear_caches")
@pytest.mark.parametrize("platform", ["mac", "windows"])
def test_a_face_without_insets_is_placed_between_calibris(monkeypatch, platform):
    calibri = lay_out("Calibri", platform)
    arial = lay_out("Arial", platform)
    refuse(monkeypatch, "text_insets")
    rendered = lay_out("Arial", platform)
    assert no_metrics(rendered) == [{"font": "Arial", "function": "text_insets"}]
    assert runs(rendered)[0][1] == runs(calibri)[0][1]
    # the right-aligned run ends where Calibri's does, its own width in from there
    (*_, x, _, width), (*_, cx, _, cwidth) = runs(rendered)[-1], runs(calibri)[-1]
    assert x + width == cx + cwidth
    assert width == runs(arial)[-1][3]


@pytest.mark.usefixtures("clear_caches")
def test_a_face_measured_only_for_widths_on_the_mac():
    # xlformat 0.9 measures Verdana's glyphs and insets, not its rows or lines
    rendered = lay_out("Verdana", "mac")
    assert no_metrics(rendered) == [
        {"font": "Verdana", "function": "row_metrics"},
        {"font": "Verdana", "function": "rule_metrics"},
    ]
    calibri = lay_out("Calibri", "mac")
    assert [r.height_pt for r in rendered.rows] == [r.height_pt for r in calibri.rows]
    assert rules(rendered) == rules(calibri)


@pytest.mark.usefixtures("clear_caches")
def test_a_face_the_mac_lacks_is_measured_in_its_substitute():
    # On the Mac, xlformat measures MS Sans Serif as Times New Roman, as Excel does.
    rendered = lay_out("MS Sans Serif", "mac")
    assert no_metrics(rendered) == []
    assert runs(rendered) == runs(lay_out("Times New Roman", "mac"))
    assert rules(rendered) == rules(lay_out("Times New Roman", "mac"))


@pytest.mark.usefixtures("clear_caches")
def test_a_normal_font_without_column_metrics_keeps_xlios_digit(monkeypatch):
    refuse(monkeypatch, "file_column_metrics")
    wb = openpyxl.Workbook()
    wb._fonts[0] = Font(name="Arial", sz=10)
    data = io.BytesIO()
    wb.save(data)
    assert _read._normal_digit(sheetpeek.Font("Arial", 10.0), xlformat.Platform.MAC) == 6.0
    sheet = sheetpeek.read(io.BytesIO(data.getvalue()))
    assert sheet.normal_font.name == "Arial"


SIZES = (1, 1.5, 5, 5.5, 9.5, 13, 72, 72.5, 73, 100, 192, 193, 286, 409, 409.5)
FACES = ("Calibri", "Verdana", "Segoe UI", "MS Sans Serif", "Courier New")


def every_size(platform: str) -> tuple[int, list[str]]:
    """Each face at each size, in every style that reaches a metrics function: how many sheets
    were laid out, and the failures.
    """
    failures, done = [], 0
    for face in FACES:
        for size in SIZES:
            wb = openpyxl.Workbook()
            wb._fonts[0] = Font(name=face, sz=size)
            ws = wb.active
            styles = (
                Font(name=face, size=size, bold=True, italic=True),
                Font(name=face, size=size, underline="singleAccounting", strike=True),
                Font(name=face, size=size, underline="double", vertAlign="superscript"),
                Font(name=face, size=size, vertAlign="subscript"),
            )
            alignments = (
                Alignment(wrap_text=True),
                Alignment(shrink_to_fit=True),
                Alignment(textRotation=90),
                Alignment(horizontal="right"),
            )
            for row, font in enumerate(styles, 1):
                for column, alignment in enumerate(alignments, 1):
                    cell = ws.cell(row, column, "Hxg 12.5")
                    cell.font, cell.alignment = font, alignment
            data = io.BytesIO()
            wb.save(data)
            try:
                sheet = sheetpeek.read(io.BytesIO(data.getvalue()), platform=platform)
                for mode in ("excel", "characters", "none"):
                    sheetpeek.render(sheet.layout(mode), "html")
                sheetpeek.render(sheet, "text")
                done += 1
            except Exception as error:
                failures.append(f"{face} {size}: {type(error).__name__}: {error}")
    return done, failures


@pytest.mark.parametrize("platform", ["mac", "windows"])
def test_every_face_and_size_lays_out(platform):
    assert in_child(every_size, platform) == (len(FACES) * len(SIZES), [])
