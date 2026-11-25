"""The capture suite ``wrap_rows`` (rows Excel fits to wrapped, justified and distributed text,
ARCHITECTURE.md §4.7), without Excel.

The workbook is built and laid out with sheetpeek, whose heights are compared with Excel's in
``tests/captures/wrap_rows.tsv`` once it exists. The reading of Excel's answer runs on a
synthetic answer.
"""

import pytest
from capture import wrap_rows as suite

import sheetpeek
from tests.capture_report import REPORT

pytest.importorskip("openpyxl")

KNOWN: dict[str, str] = {}
"""Cases whose row sheetpeek fits differently from Excel for Mac, with the reason."""


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    return suite.build(tmp_path_factory.mktemp("wrap_rows"))


def test_each_case_is_built_as_described(built):
    sheet = sheetpeek.read(built)
    assert len(suite.CASES) == 96
    for row, case in enumerate(suite.CASES, 1):
        cell = sheet.cell(row, 1)
        assert (cell.font.name, cell.font.size) == (case.face, case.size), case.name
        assert cell.value == case.text, case.name
        assert cell.style.h_align == {"wrap": "general", "justify": "justify"}.get(
            case.alignment, "distributed"
        ), case.name
        assert cell.style.wrap == (case.alignment != "distributed unwrapped"), case.name
    assert len({c.name for c in suite.CASES}) == len(suite.CASES)


def test_a_words_case_takes_a_line_a_word(built):
    # One word fits a line of the default column, two do not: sheetpeek wraps each case's
    # words into as many lines as it has words, on both platforms.
    for platform in ("mac", "windows"):
        rendered = sheetpeek.read(built, platform=platform).layout("excel")
        for row, case in enumerate(suite.CASES, 1):
            if case.form == "words" and case.alignment == "wrap":
                (laid,) = [r for r in rendered.rows if r.index == row]
                assert len(laid.runs) == case.lines, (platform, case.name)


def test_an_answer_reads_into_records(tmp_path):
    answer = tmp_path / "answer.tsv"
    lines = ["# excel: 16.115 (test), en_US", "# captured: 2026-11-24"]
    lines += [
        f"{suite.WORKBOOK}\t{k}\t{c.face}\t{c.size}\t15.0" for k, c in enumerate(suite.CASES, 1)
    ]
    answer.write_text("\n".join(lines) + "\n", encoding="utf-8")
    out = tmp_path / "wrap_rows.tsv"
    suite.write(answer, out)
    rows = [r for r in out.read_text(encoding="utf-8").splitlines() if not r.startswith("#")]
    assert rows[0].split("\t") == list(suite.COLUMNS)
    assert len(rows) == len(suite.CASES) + 1
    with pytest.raises(ValueError, match="expected"):
        suite.records("\n".join(lines[:-1]))


@pytest.mark.excel
def test_sheetpeek_fits_the_rows_excel_fits(built):
    if not suite.OUTPUT.exists():
        pytest.skip("not captured yet: run python -m capture.wrap_rows run")
    rows = [
        r.split("\t")
        for r in suite.OUTPUT.read_text(encoding="utf-8").splitlines()
        if r and not r.startswith("#")
    ]
    captured = {r[0]: float(r[7]) for r in rows[1:]}
    heights = {r.index: r.height_pt for r in sheetpeek.read(built).layout("excel").rows}
    ours = {case.name: heights[row] for row, case in enumerate(suite.CASES, 1)}
    differ = sorted(name for name in captured if ours[name] != captured[name])
    REPORT["wrap_rows"] = (
        f"wrap_rows: {len(captured) - len(differ)} of {len(captured)} rows as Excel fits them, "
        f"{len(differ)} known deviations"
    )
    assert set(differ) == set(KNOWN), [(n, ours[n], captured[n]) for n in differ]
