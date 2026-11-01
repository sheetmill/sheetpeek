"""Golden output of every backend for one workbook that exercises the layout.

The expected files under ``tests/golden/`` are regenerated on purpose with
``SHEETPEEK_UPDATE_GOLDEN=1 pytest tests/test_backends.py``, then read before committing.
"""

import datetime as dt
import io
import os
from pathlib import Path

import openpyxl
import pytest
from openpyxl.styles import Alignment, Font, PatternFill

import sheetpeek
from sheetpeek.backends import (
    DelimitedBackend,
    HTMLBackend,
    MarkdownBackend,
    RichBackend,
    TextBackend,
    backend_names,
    get_backend,
    register_backend,
)
from tests.workbooks import save

GOLDEN = Path(__file__).with_name("golden")
UPDATE = os.environ.get("SHEETPEEK_UPDATE_GOLDEN") == "1"


def build_showcase(path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"
    header = PatternFill("solid", fgColor="FFDDEBF7")
    ws["A1"] = "Quarterly report, all regions"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:D1")
    for col, title in zip("ABCDE", ["Region", "Sales", "Change", "Updated", "Note"], strict=True):
        ws[f"{col}2"] = title
        ws[f"{col}2"].font = Font(bold=True)
        ws[f"{col}2"].fill = header
    rows = [
        ("North", 1234567.891, 0.125, dt.datetime(2026, 10, 1, 9, 30), "on target"),
        ("South", 98765.4, -0.0425, dt.datetime(2026, 10, 2, 17, 5), None),
        ("East", -4321.5, 0, dt.datetime(2026, 10, 3), "a note long enough to spill"),
        ("West", 123456789012.0, None, None, None),
    ]
    for r, (region, sales, change, updated, note) in enumerate(rows, 3):
        ws[f"A{r}"] = region
        ws[f"B{r}"] = sales
        ws[f"B{r}"].number_format = "#,##0.00;[Red]-#,##0.00"
        if change is not None:
            ws[f"C{r}"] = change
            ws[f"C{r}"].number_format = "0.0%"
        if updated is not None:
            ws[f"D{r}"] = updated
            ws[f"D{r}"].number_format = "yyyy-mm-dd hh:mm"
        if note:
            ws[f"E{r}"] = note
    ws["A7"] = "Total"
    ws["A7"].font = Font(bold=True, italic=True)
    ws["B7"] = 123457928993.791
    ws["B7"].number_format = "#,##0"
    ws["C7"] = True
    ws["D7"] = "| pipes, commas"
    ws["A8"] = "Wrapped note in a narrow column"
    ws["A8"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["B8"] = "right aligned spill"
    ws["B8"].alignment = Alignment(horizontal="right")
    ws["C8"].fill = PatternFill("solid", fgColor="FFFFFF00")
    ws["E8"] = "hidden row below"
    ws["A9"] = "this row is hidden"
    ws.row_dimensions[9].hidden = True
    ws["A10"] = "Underlined"
    ws["A10"].font = Font(underline="single", color="FF0070C0")
    ws["B10"] = "centered across"
    for col in "BCD":
        ws[f"{col}10"].alignment = Alignment(horizontal="centerContinuous")
    ws["F2"] = "secret"
    ws.column_dimensions["F"].hidden = True
    ws.column_dimensions["A"].width = 12.7109375
    ws.column_dimensions["E"].width = 14.7109375
    return save(wb, path)


@pytest.fixture(scope="module")
def showcase(tmp_path_factory):
    return sheetpeek.read(build_showcase(tmp_path_factory.mktemp("golden") / "showcase.xlsx"))


def check(name: str, text: str) -> None:
    path = GOLDEN / name
    if UPDATE:
        GOLDEN.mkdir(exist_ok=True)
        path.write_text(text, encoding="utf-8", newline="")
    assert path.exists(), f"missing golden file {name}; run with SHEETPEEK_UPDATE_GOLDEN=1"
    assert text == path.read_text(encoding="utf-8")


@pytest.mark.parametrize(
    ("name", "backend"),
    [
        ("showcase.txt", TextBackend()),
        ("showcase_headings.txt", TextBackend("│", column_headings=True, row_headings=True)),
        ("showcase.md", MarkdownBackend()),
        ("showcase_letters.md", MarkdownBackend("letters", emphasis=True)),
        ("showcase.csv", DelimitedBackend(",")),
        ("showcase.tsv", DelimitedBackend("\t")),
        ("showcase.html", HTMLBackend(column_headings=True, row_headings=True)),
    ],
    ids=lambda value: value if isinstance(value, str) else "",
)
def test_golden(showcase, name, backend):
    check(name, sheetpeek.render(showcase, backend))


def test_golden_rich(showcase):
    pytest.importorskip("rich")
    backend = RichBackend(
        column_headings=True, row_headings=True, color_system="truecolor", force_terminal=True
    )
    check("showcase.ansi", sheetpeek.render(showcase, backend))


def test_rich_without_a_terminal_is_plain(showcase):
    pytest.importorskip("rich")
    plain = sheetpeek.render(showcase, RichBackend(force_terminal=False))
    text = sheetpeek.render(showcase, TextBackend())
    assert "\x1b" not in plain
    assert [line.rstrip() for line in plain.splitlines()] == text.splitlines()


def test_rich_print_hook(showcase):
    rich_console = pytest.importorskip("rich.console")
    console = rich_console.Console(file=io.StringIO(), width=200, color_system=None)
    console.print(showcase)
    printed = [line.rstrip() for line in console.file.getvalue().splitlines()]
    assert printed == sheetpeek.render(showcase, "text").splitlines()


def test_render_sources(tmp_path, showcase):
    path = build_showcase(tmp_path / "s.xlsx")
    from_path = sheetpeek.render(path, "csv", sheet="Report", ref="A2:B3")
    assert from_path == 'Region,Sales\nNorth,"1,234,567.89"\n'
    rendered = showcase.layout("characters")
    assert sheetpeek.render(rendered, "csv") == sheetpeek.render(showcase, "csv")  # relaid out
    out = io.StringIO()
    assert sheetpeek.render(showcase, "tsv", out=out) is None
    assert out.getvalue().startswith("Quarterly report, all regions\t")


def test_registry(showcase):
    assert backend_names() == ["csv", "html", "markdown", "rich", "text", "tsv"]
    with pytest.raises(ValueError, match="no backend named 'pdf'"):
        get_backend("pdf")

    class Shout:
        name = "shout"
        mode = "none"

        def render(self, sheet, out, /):
            out.write(sheet.rows[0].cells[0].lines[0].upper())

    assert isinstance(Shout(), sheetpeek.Backend)
    assert sheetpeek.render(showcase, Shout()) == "QUARTERLY REPORT, ALL REGIONS"
    register_backend("shout", Shout)
    try:
        assert sheetpeek.render(showcase, "shout") == "QUARTERLY REPORT, ALL REGIONS"
    finally:
        from sheetpeek import backends

        del backends._REGISTRY["shout"]


def test_empty_sheet(tmp_path):
    sheet = sheetpeek.read(save(openpyxl.Workbook(), tmp_path / "empty.xlsx"))
    for name in ("text", "markdown", "csv"):
        assert sheetpeek.render(sheet, name) == ""


def test_separator_must_be_one_character():
    with pytest.raises(ValueError, match="one character"):
        TextBackend(" | ")
    with pytest.raises(ValueError, match="'first' or 'letters'"):
        MarkdownBackend("none")
