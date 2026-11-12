import io
import re
import subprocess
import sys

import openpyxl
import pytest

import sheetpeek
from sheetpeek import cli
from tests.workbooks import replace_part, rewrite, save


@pytest.fixture
def book(tmp_path):
    wb = openpyxl.Workbook()
    wb.active.title = "Data"
    wb.active["A1"] = "a long text value"
    wb.active["A2"] = 1234.5
    wb.active["A3"] = "=1+1"
    wb.create_sheet("Other")["B2"] = "other"
    return save(wb, tmp_path / "book.xlsx")


def run(*args, terminal=False):
    out, err = io.StringIO(), io.StringIO()
    out.isatty = lambda: terminal
    parser = cli._parser()
    status = cli._view(parser.parse_args(["view", *map(str, args)]), out, err)
    return status, out.getvalue(), err.getvalue()


def test_view_is_text_when_piped(book):
    status, out, err = run(book)
    assert status == 0
    # The used range is column A alone; the text spills into an added column.
    assert out == "a long text value\n  1234.5\n\n"
    assert "1 formula cell(s) have no cached value" in err


def test_view_is_rich_on_a_terminal(book):
    pytest.importorskip("rich")
    status, out, _ = run(book, "--ref", "A1:A2", terminal=True)
    assert status == 0
    assert out.splitlines()[1].rstrip() == "  1234.5"


def test_view_options(book):
    assert run(book, "--sheet", "Other", "--column-headings", "--row-headings")[1] == (
        "     B\n2 other\n"
    )
    assert run(book, "--sheet", "Other", "--column-headings")[1] == "   B\nother\n"
    assert run(book, "--sheet", "Other", "--row-headings")[1] == "2 other\n"
    assert run(book, "-t", "csv", "-r", "A1:A2")[1] == "a long text value\n1234.5\n"
    page = run(book, "-t", "html", "--column-headings")[1]
    assert page.startswith("<!doctype html>")
    assert '<th class="sp-colhead" scope="col">A</th>' in page
    assert run(book, "--list")[1] == "Data\nOther\n"


def test_view_errors(book, tmp_path):
    status, _, err = run(book, "--sheet", "Missing")
    assert (status, err) == (
        1,
        "sheetpeek: no sheet named 'Missing'; the sheets are ['Data', 'Other']\n",
    )
    status, _, err = run(tmp_path / "absent.xlsx")
    assert status == 1
    assert err.startswith("sheetpeek: no such file:")
    (tmp_path / "not.xlsx").write_text("not a workbook")
    status, _, err = run(tmp_path / "not.xlsx")
    assert status == 1
    assert err.startswith("sheetpeek: cannot read")
    status, _, err = run(book, "--ref", "A1:")
    assert (status, err) == (1, "sheetpeek: not a cell reference: 'A1:'\n")


def test_module_entry_point(book):
    result = subprocess.run(
        [sys.executable, "-m", "sheetpeek", "view", str(book), "-t", "tsv", "-r", "A2"],
        capture_output=True,
        text=True,
        check=True,
    )
    assert result.stdout == "1234.5\n"


def test_view_limits(book):
    status, _, err = run(book, "--max-parts", "1")
    assert status == 1
    assert "parts, more than the limit of 1" in err
    status, _, err = run(book, "--max-cells", "1")
    assert status == 1
    assert "more than the limit of 1" in err
    assert str(book.parent) not in err.replace(str(book), "")


def test_limit_options_must_be_counts(book):
    with pytest.raises(SystemExit):
        run(book, "--max-parts", "-1")
    with pytest.raises(SystemExit):
        run(book, "--max-parts", "many")


def test_view_refused_layout(book, tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "1" * 300
    long = tmp_path / "long.xlsx"
    wb.save(long)
    status, _, err = run(long, "--max-cells", "2", "--ref", "A1:A2")
    assert status == 1
    assert err.startswith("sheetpeek: cannot show")


def test_serve_help_says_what_it_is_not(capsys):
    with pytest.raises(SystemExit):
        cli.main(["serve", "--help"])
    text = " ".join(capsys.readouterr().out.split())
    assert "no authentication" in text
    assert "Not for production" in text


# --- what reaches a terminal: ESC is 0x1b, and U+009B is the one-character form of ESC [

CONTROL = re.compile("[\\x00-\\x09\\x0b-\\x1f\\x7f-\\x9f]")


@pytest.fixture
def c1_name(tmp_path):
    # XML can hold U+009B, by reference; it cannot hold ESC at all.
    wb = openpyxl.Workbook()
    wb.active["A1"] = "cell"
    path = save(wb, tmp_path / "c1.xlsx")
    rewrite(path, "xl/workbook.xml", lambda t: t.replace('name="Sheet"', 'name="a&#155;[31mb"'))
    return path


def test_list_prints_no_control_character(c1_name):
    status, out, _ = run(c1_name, "--list")
    assert status == 0
    assert out == "a[31mb\n"


def test_list_with_esc_in_a_name(book, monkeypatch):
    # Not from XML, which cannot hold it; the rule does not depend on that.
    class Stub:
        def __init__(self):
            self.sheets = (
                sheetpeek.SheetInfo("a\x1b[31mb\x9b[2Jc\nforged", "worksheet", "visible"),
                sheetpeek.SheetInfo("d\x1bx\ny", "chartsheet", "very_hidden"),
            )

        def __enter__(self):
            return self

        def __exit__(self, *args):
            pass

    monkeypatch.setattr(cli, "open_book", lambda *a, **k: Stub())
    assert run(book, "--list")[1] == "a[31mb[2Jcforged\ndxy  (very hidden, chartsheet, not shown)\n"


def test_the_model_and_sheet_keep_the_original_name(c1_name):
    with sheetpeek.open(c1_name) as opened:
        assert opened.sheet_names == ["a\x9b[31mb"]
    status, out, _ = run(c1_name, "--sheet", "a\x9b[31mb")
    assert status == 0
    assert "cell" in out
    # What --list shows is not the name, and the error says what the name is, in repr's escapes.
    status, _, err = run(c1_name, "--sheet", "a[31mb")
    assert status == 1
    assert "\\x9b" in err
    assert not CONTROL.search(err)


def test_html_on_stdout_has_no_control_character(c1_name):
    assert "<title>a[31mb</title>" in run(c1_name, "-t", "html")[1]
    assert not CONTROL.search(run(c1_name, "-t", "html")[1])


def test_a_font_name_in_html_has_no_control_character(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "x"
    wb.active["A1"].font = openpyxl.styles.Font(name="Zed")
    path = save(wb, tmp_path / "font.xlsx")
    rewrite(path, "xl/styles.xml", lambda t: t.replace('val="Zed"', 'val="Z&#155;ed"'))
    status, out, err = run(path, "-t", "html")
    assert status == 0
    assert "Zed" in out
    assert not CONTROL.search(out)
    assert not CONTROL.search(err)


def test_a_warning_has_no_control_character(tmp_path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.active["A1"].number_format = "0.00E"
    path = save(wb, tmp_path / "format.xlsx")
    rewrite(path, "xl/styles.xml", lambda t: t.replace("0.00E", "0.00E&#155;"))
    status, _, err = run(path)
    assert status == 0
    assert "format.refused" in err
    assert not CONTROL.search(err)


def test_a_part_name_in_a_refusal_has_no_control_character(book):
    replace_part(book, "xl/\x1b[31m\x9bevil.xml", b"<!DOCTYPE x><x/>")
    status, _, err = run(book)
    assert status == 1
    assert "[31mevil.xml: declares a DTD" in err
    assert not CONTROL.search(err)


def test_a_file_name_has_no_control_character(tmp_path):
    (tmp_path / "a\x1b[31m.xlsx").write_text("not a workbook")
    status, _, err = run(tmp_path / "a\x1b[31m.xlsx")
    assert status == 1
    assert not CONTROL.search(err)
    status, _, err = run(tmp_path / "b\x1b[31m.xlsx")  # no such file
    assert status == 1
    assert not CONTROL.search(err)


def test_serve_says_nothing_with_a_control_character(tmp_path):
    out, err = io.StringIO(), io.StringIO()
    args = cli._parser().parse_args(["serve", str(tmp_path / "no\x1b[31m")])
    assert cli._serve(args, out, err) == 1
    assert "no[31m" in err.getvalue()
    assert not CONTROL.search(err.getvalue())
