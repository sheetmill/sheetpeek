import io
import subprocess
import sys

import openpyxl
import pytest

from sheetpeek import cli
from tests.workbooks import save


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


def test_serve_help_says_there_is_no_authentication(capsys):
    with pytest.raises(SystemExit):
        cli.main(["serve", "--help"])
    assert "no authentication" in " ".join(capsys.readouterr().out.split())
