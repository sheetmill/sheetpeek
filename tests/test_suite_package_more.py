"""The capture suite ``package_more`` (sizes no column, row or font can have), without Excel.

Its workbooks are built, each holding the one change its case names, and read back with sheetpeek
in a child under the watchdog (the cases are extreme numbers, and a regressed guard could turn
one into a huge allocation); what sheetpeek takes for each is the suite's ``EXPECTED``. Its
measuring (Excel's AppleScript answers turned into records) runs on synthetic answers. The
comparison with ``tests/captures/package_more.tsv`` is skipped until that capture exists.
"""

import io
import zipfile

import pytest
from capture import package_more as suite
from capture import package_probe

import sheetpeek
from tests.capture_report import REPORT
from tests.watchdog import in_child

CAPTURE = suite.OUTPUT


def sheetpeek_takes(names=None):
    """Case → None where sheetpeek refuses it, else (width A, width B, height 1, height 2, size).

    Widths and heights in points from the excel layout with hidden columns shown, as the
    ``package`` comparison in ``tests/test_captures.py`` lays them out. Run in a child.
    """
    data = package_probe.base()
    found = {}
    for name, case in suite.CASES:
        if names is not None and name not in names:
            continue
        try:
            sheet = sheetpeek.read(io.BytesIO(case(data)), ref="A1:B2", platform="mac")
        except sheetpeek.PackageError:
            found[name] = None
            continue
        rendered = sheet.layout("excel", show_hidden=True)
        sheetpeek.render(rendered, "html")
        sheet.layout("characters")
        found[name] = (
            rendered.columns[0].width_pt,
            rendered.columns[1].width_pt,
            rendered.rows[0].height_pt,
            rendered.rows[1].height_pt,
            sheet.cell(1, 1).style.font.size,
        )
    return found


def normal_fonts():
    """Case → (name, size) of the Normal font sheetpeek reads, for every case it reads."""
    data = package_probe.base()
    found = {}
    for name, case in suite.CASES:
        try:
            font = sheetpeek.read(io.BytesIO(case(data)), ref="A1:B2", platform="mac").normal_font
        except sheetpeek.PackageError:
            continue
        found[name] = (font.name, font.size)
    return found


@pytest.fixture(scope="module")
def taken():
    return in_child(sheetpeek_takes)


# --- the workbooks


def test_the_cases_cover_the_rules():
    names = [name for name, _ in suite.CASES]
    assert names[0] == "control"
    assert names[-1] == "control_again"
    assert len(set(names)) == len(names) == len(suite.EXPECTED)
    for prefix, count in (
        ("width_1e", 3),
        ("default_width_", 4),
        ("base_width_", 4),
        ("font_size_", 5),
        ("arial_10", 4),
        ("calibri_14", 4),
    ):
        assert sum(name.startswith(prefix) for name in names) == count, prefix
    # a name shared with the ``package`` suite is the same change: the controls, and the refused
    # anchors kept so that each edge is found between two cases of one run
    shared = set(names) & {name for name, _ in package_probe.CASES}
    assert shared == {
        "control",
        "control_again",
        "default_width_300",
        "base_width_300",
        "font_size_410",
    }
    base = package_probe.base()
    package = dict(package_probe.CASES)
    for name in shared:
        with (
            zipfile.ZipFile(io.BytesIO(dict(suite.CASES)[name](base))) as ours,
            zipfile.ZipFile(io.BytesIO(package[name](base))) as theirs,
        ):
            assert ours.namelist() == theirs.namelist()
            for part in ours.namelist():
                assert ours.read(part) == theirs.read(part), (name, part)


def test_every_workbook_holds_its_change(tmp_path):
    paths = suite.build(tmp_path)
    assert [p.name for p in paths] == [f"package_more_{name}.xlsx" for name, _ in suite.CASES]
    changes = {
        "width_1e100": ("xl/worksheets/sheet1.xml", 'width="1e100" customWidth="1"'),
        "default_width_299": ("xl/worksheets/sheet1.xml", 'defaultColWidth="299" '),
        "base_width_256": ("xl/worksheets/sheet1.xml", 'baseColWidth="256"'),
        "font_size_409.5": ("xl/styles.xml", '<sz val="409.5"'),
        "font_size_minus_10": ("xl/styles.xml", '<sz val="-10"'),
        "arial_10": ("xl/styles.xml", '<name val="Arial" /><family val="2" /><color theme="1" />'
                     '<sz val="10" /></font>'),
        "calibri_14_width_1e6": ("xl/worksheets/sheet1.xml", 'width="1e6" customWidth="1"'),
    }  # fmt: skip
    for path, (name, _) in zip(paths, suite.CASES, strict=True):
        with zipfile.ZipFile(path) as archive:
            assert archive.testzip() is None
            assert b"<v>111</v>" in archive.read("xl/worksheets/sheet1.xml")
            if name in changes:
                part, text = changes[name]
                assert text in archive.read(part).decode(), name
    with zipfile.ZipFile(tmp_path / "package_more_calibri_14_width_1e6.xlsx") as archive:
        assert '<sz val="14" /></font>' in archive.read("xl/styles.xml").decode()
        assert "scheme" not in archive.read("xl/styles.xml").decode().split("</font>")[0]


def test_the_workbooks_read_back_as_built():
    fonts = in_child(normal_fonts)
    assert fonts["control"] == fonts["control_again"] == ("Calibri", 11.0)
    for name in ("arial_10", "arial_10_width_minus_1", "arial_10_width_1e30"):
        assert fonts[name] == ("Arial", 10.0)
    for name in ("calibri_14", "calibri_14_width_1e6"):
        assert fonts[name] == ("Calibri", 14.0)
    assert fonts["font_size_minus_10"] == ("Calibri", 10.0)


def test_expected_is_what_sheetpeek_takes(taken):
    assert taken == suite.EXPECTED


# --- the measuring, on synthetic answers


def _opened(sizes):
    width_a, width_b, height_1, height_2, size = sizes
    return (
        f"opened\t111.0\t1\t8.0\t{width_a}\t{height_1}\t8.0\t{width_b}\t{height_2}\t{size}\tCalibri"
    )


def _answer(override=None):
    lines = []
    for name, _ in suite.CASES:
        sizes = suite.EXPECTED[name]
        outcome = "refused\t-50" if sizes is None else _opened(sizes)
        lines.append(f"{name}\t{(override or {}).get(name, outcome)}")
    return "\n".join(lines)


def test_measure_returns_the_answers():
    records = suite.measure(_answer())
    assert [r["case"] for r in records] == [name for name, _ in suite.CASES]
    first = records[0]
    assert first == {
        "case": "control",
        "outcome": "opened",
        "A1": "111.0",
        "sheets": "1",
        "width A": "8.0",
        "width A pt": "53.0",
        "height 1": "15.0",
        "width B": "8.0",
        "width B pt": "53.0",
        "height 2": "15.0",
        "font size": "11.0",
        "font name": "Calibri",
    }


def test_measure_keeps_a_refusal_and_its_number():
    records = suite.measure(_answer({"base_width_256": "refused\t-50"}))
    record = next(r for r in records if r["case"] == "base_width_256")
    assert record["outcome"] == "refused"
    assert record["A1"] == "-50"
    assert record["width A pt"] == record["font name"] == ""


def test_measure_reads_each_field_in_its_place():
    answer = "opened\t111.0\t1\t1.0921E+4\t65526.0\t15.0\t8.0\t53.0\t14.0\t10.0\tArial"
    records = suite.measure(_answer({"arial_10_width_minus_1": answer}))
    record = next(r for r in records if r["case"] == "arial_10_width_minus_1")
    assert (record["width A"], record["width A pt"], record["width B pt"]) == (
        "1.0921E+4",
        "65526.0",
        "53.0",
    )
    assert (record["height 1"], record["height 2"]) == ("15.0", "14.0")
    assert (record["font size"], record["font name"]) == ("10.0", "Arial")


def test_measure_refuses_answers_that_do_not_fit():
    lines = _answer().split("\n")
    with pytest.raises(ValueError, match=f"expected {len(suite.CASES)} cases"):
        suite.measure("\n".join(lines[:-1]))
    with pytest.raises(ValueError, match="answered for width_1e200"):
        suite.measure("\n".join([lines[0], lines[2], lines[1], *lines[3:]]))
    with pytest.raises(ValueError, match="not an answer"):
        suite.measure("\n".join([lines[0], "width_1e100\topened\t111.0", *lines[2:]]))
    with pytest.raises(ValueError, match="control_again did not open"):
        suite.measure(_answer({"control_again": "refused\t-50"}))


def test_write_measures_saved_answers_again(tmp_path):
    saved = tmp_path / "excel.tsv"
    saved.write_text("# excel: 16.115\n# date: 2026-01-01\n" + _answer() + "\n", encoding="utf-8")
    output = tmp_path / "out.tsv"
    suite.write(saved, output)
    lines = output.read_text(encoding="utf-8").splitlines()
    assert lines[:3] == ["# excel: 16.115", "# date: 2026-01-01", "# suite: package_more"]
    rows = [line.split("\t") for line in lines if not line.startswith("# ")]
    assert rows[0] == list(suite.COLUMNS)
    assert len(rows) == 1 + len(suite.CASES)
    assert rows[1][:2] == ["control", "opened"]
    assert rows[-1][:2] == ["control_again", "opened"]
    refused = [row for row in rows[1:] if row[1] == "refused"]
    assert refused
    assert all(row[2:] == ["-50"] for row in refused)


# --- the comparison with Excel
#
# Tolerance none. Excel answers whole or half points for widths (``width`` of a range) and
# heights, and the ``package`` suite's 68 cases match sheetpeek's widths, heights and font sizes
# exactly; a width a point off here is the rule failing (the 6 pt digit, or the edge where 0
# starts), not noise. The comparison uses the
# width in points, which needs no font measurements, and records
# the width in characters as Excel reports it beside it.

KNOWN_DEVIATIONS: dict[str, str] = {}

PINNED = {
    # Excel reports a size of 409.5 as 409 where sheetpeek keeps 409.5; both lay it out at 1,769 pt,
    # wider than 409's 1,761.
    "font_size_409.5": "Excel reports 409.5 pt as 409",
}
"""Cases where sheetpeek differs from Excel for a reason recorded here, most of them in a sibling
package. Each comparison asserts exactly these, so a new difference fails, and so does a fixed
one (unpin it)."""


def _pinned(captured, prefixes):
    return {n for n in PINNED if n in captured and n.startswith(prefixes)}


@pytest.fixture(scope="module")
def captured():
    if not CAPTURE.exists():
        pytest.skip("not captured yet: run python -m capture.package_more run")
    lines = [ln for ln in CAPTURE.read_text(encoding="utf-8").splitlines() if ln[:2] != "# "]
    names = lines[0].split("\t")
    records = []
    for line in lines[1:]:
        fields = line.split("\t")
        records.append(dict(zip(names, fields + [""] * (len(names) - len(fields)), strict=True)))
    return {record["case"]: record for record in records}


def _excel(record):
    """Excel's answer in ``EXPECTED``'s form."""
    if record["outcome"] == "refused":
        return None
    return tuple(
        float(record[key])
        for key in ("width A pt", "width B pt", "height 1", "height 2", "font size")
    )


def _compare(captured, taken, prefixes):
    deviations = {}
    for name, record in captured.items():
        if not name.startswith(prefixes):
            continue
        excel, ours = _excel(record), taken[name]
        if excel is None or ours is None:
            if excel != ours:
                deviations[name] = (ours, excel)
            continue
        if name in KNOWN_DEVIATIONS:
            excel, ours = (excel[0], excel[1], excel[4]), (ours[0], ours[1], ours[4])
        if ours != excel:
            deviations[name] = (ours, excel)
    return deviations


def _report(key, label, count, deviations):
    REPORT[key] = (
        f"{label}: {count} cases compared with Excel, {len(deviations)} deviations"
        + "".join(f"; {name}" for name in sorted(deviations))
    )


@pytest.mark.excel
def test_the_capture_holds_every_case(captured):
    assert list(captured) == [name for name, _ in suite.CASES]
    for name in ("control", "control_again"):
        record = captured[name]
        assert (record["outcome"], record["A1"], record["sheets"]) == ("opened", "111.0", "1")
    assert all(r["A1"] == "111.0" for r in captured.values() if r["outcome"] == "opened")


@pytest.mark.excel
def test_where_a_huge_width_becomes_0(captured, taken):
    # Rule package_limits (§5.5, ``<col width>``): from 1e9 to 1e30 a width is one point, and
    # 1e308 is 0; 1e100, 1e200 and 1e300 tell where 0 starts.
    deviations = _compare(captured, taken, ("width_", "control"))
    _report("package_more widths", "package_more, huge widths and the controls", 5, deviations)
    assert set(deviations) == _pinned(captured, ("width_", "control")), deviations


@pytest.mark.excel
def test_default_and_base_widths_at_their_edges(captured, taken):
    # Rule package_limits (§5.5, ``defaultColWidth``, ``baseColWidth``): 300 is refused for both;
    # 255, 256 and 299 tell where the refusal starts and what a width there is.
    deviations = _compare(captured, taken, ("default_width_", "base_width_"))
    _report("package_more sheet widths", "package_more, sheet default widths", 8, deviations)
    assert set(deviations) == _pinned(captured, ("default_width_", "base_width_")), deviations


@pytest.mark.excel
def test_font_sizes_at_their_edges(captured, taken):
    # Rule package_limits (§5.5, font size): its magnitude from 1 to 409, 410 refused; 409.5 is
    # the edge, and -2 and -10 tell whether any size below -1 is its magnitude.
    deviations = _compare(captured, taken, ("font_size_",))
    _report("package_more font sizes", "package_more, font sizes", 5, deviations)
    assert set(deviations) == _pinned(captured, ("font_size_",)), deviations


@pytest.mark.excel
def test_the_digit_of_other_normal_fonts(captured, taken):
    # Rule package_limits (§5.5, ``<col width>``): kept in points modulo 65,536 at the Normal
    # font's digit, as Calibri 14's 7 pt shows. Excel's
    # font name confirms the Normal font took.
    for name, record in captured.items():
        if name.startswith("calibri_14"):
            assert record["font name"] == "Calibri", name
    deviations = _compare(captured, taken, ("calibri_14",))
    _report("package_more Calibri 14", "package_more, a Calibri 14 Normal font", 4, deviations)
    assert set(deviations) == _pinned(captured, ("calibri_14",)), deviations


@pytest.mark.excel
def test_an_arial_10_normal_font_as_captured(captured, taken):
    # Arial 10's digit is 6 pt, as Calibri 11's, so its huge widths wrap as Calibri 11's do.
    for name, record in captured.items():
        if name.startswith("arial_10"):
            assert record["font name"] == "Arial", name
    deviations = _compare(captured, taken, ("arial_10",))
    _report("package_more Arial 10", "package_more, an Arial 10 Normal font", 4, deviations)
    assert deviations == {}
