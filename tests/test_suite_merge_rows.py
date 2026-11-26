"""The capture suite ``merge_rows`` (rows Excel fits around vertical merges, ARCHITECTURE.md §4.7),
without Excel.

The workbook is built and laid out with sheetpeek, whose heights must be Excel's in
``tests/captures/merge_rows.tsv``. ``PREDICTIONS`` stays as written before the run: sheetpeek's
then held for 18 of the 22 rows, counting a merge's font in its first row, which Excel does not.
The reading of Excel's answer runs on a synthetic answer.
"""

import pytest
from capture import merge_rows as suite

import sheetpeek

pytest.importorskip("openpyxl")


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    return suite.build(tmp_path_factory.mktemp("merge_rows"))


def _heights(path):
    return {
        row.index: row.height_pt
        for row in sheetpeek.read(path, platform="mac").layout("excel").rows
    }


def test_each_case_is_built_and_merged_as_described(built):
    sheet = sheetpeek.read(built, platform="mac")
    merges = {str(m) for m in sheet.merges}
    for block in suite.BLOCKS:
        first = next(c.row for c in suite.CASES if c.block is block)
        cell = sheet.cell(first, 1)
        assert cell.font.size == block.size, block.label
        assert cell.value is not None, block.label
        assert cell.style.wrap == bool(block.lines), block.label
        if block.rows > 1 or block.wide > 1:
            end = "ABC"[block.wide - 1]
            assert f"A{first}:{end}{first + block.rows - 1}" in merges, block.label
    assert len(suite.CASES) == sum(b.rows for b in suite.BLOCKS)


def test_the_candidates_predict_every_row_and_differ():
    keys = {(c.block.label, c.offset) for c in suite.CASES}
    tables = list(suite.PREDICTIONS.values())
    assert all(set(table) == keys for table in tables)
    assert len({tuple(sorted(table.items())) for table in tables}) == len(tables)


def test_an_answer_reads_into_records(tmp_path):
    answer = tmp_path / "answer.tsv"
    lines = ["# excel: 16.115 (test), en_US", "# captured: 2026-11-23"]
    lines += [f"{suite.WORKBOOK}\t{c.row}\tCalibri\t{c.block.size}\t15.0" for c in suite.CASES]
    answer.write_text("\n".join(lines) + "\n", encoding="utf-8")
    out = tmp_path / "merge_rows.tsv"
    suite.write(answer, out)
    rows = [r for r in out.read_text(encoding="utf-8").splitlines() if not r.startswith("#")]
    assert rows[0].split("\t") == list(suite.COLUMNS)
    assert len(rows) == len(suite.CASES) + 1
    with pytest.raises(ValueError, match="expected"):
        suite.records("\n".join(lines[:-1]))


def _captured():
    rows = [
        r.split("\t")
        for r in suite.OUTPUT.read_text(encoding="utf-8").splitlines()
        if r and not r.startswith("#")
    ]
    header, body = rows[0], rows[1:]
    return {
        (r["label"], int(r["offset"])): float(r["height"])
        for r in (dict(zip(header, row, strict=True)) for row in body)
    }


@pytest.mark.excel
def test_sheetpeek_fits_the_rows_excel_fits(built):
    heights = _heights(built)
    assert {(c.block.label, c.offset): heights[c.row] for c in suite.CASES} == _captured()


@pytest.mark.excel
def test_the_predictions_written_before_the_run_score_as_recorded():
    captured = _captured()
    score = {
        name: sum(table[key] == height for key, height in captured.items())
        for name, table in suite.PREDICTIONS.items()
    }
    assert score == {"sheetpeek": 18, "last_row": 11, "first_row": 12, "font_every_row": 13}
