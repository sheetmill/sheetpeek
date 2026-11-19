"""The capture suite ``odd_values`` (odd ``textRotation`` and undefined ``numFmtId``), no Excel.

Its workbooks are built and read back with sheetpeek as built, and what sheetpeek reads from
each is the suite's ``EXPECTED``. The parsing of the script's answers and the writing of the
TSV run on synthetic answers. The comparison with ``tests/captures/odd_values.tsv`` is skipped
until that capture exists.
"""

import zipfile

import pytest
from capture import odd_values

# --- the workbooks


def test_every_case_builds_with_its_values_and_reads_back_as_expected(tmp_path):
    paths = odd_values.build(tmp_path)
    assert len(paths) == len(odd_values.CASES)
    for case, path in zip(odd_values.CASES, paths, strict=True):
        with zipfile.ZipFile(path) as package:
            styles = package.read(odd_values.STYLES).decode()
        assert f'textRotation="{case.rotation}"' in styles, case
        assert f'numFmtId="{case.format_id}"' in styles, case
        assert "<numFmt " not in styles, case  # no id is defined: 4 is built in
        # EXPECTED is what sheetpeek reads today, not a guess at Excel
        assert odd_values.sheetpeek_reads(path.read_bytes()) == odd_values.EXPECTED[case.name]


def test_the_cases_are_named_once_between_two_controls():
    names = [case.name for case in odd_values.CASES]
    assert len(set(names)) == len(names)
    assert set(names) == set(odd_values.EXPECTED)
    first, last = odd_values.CASES[0], odd_values.CASES[-1]
    for control in (first, last):
        assert control.name.startswith("control")
        assert (control.rotation, control.format_id) == (
            odd_values.CONTROL_ROTATION,
            odd_values.CONTROL_FORMAT_ID,
        )
    # each odd case changes one value of the control
    for case in odd_values.CASES[1:-1]:
        changed = (case.rotation != first.rotation) + (case.format_id != first.format_id)
        assert changed == 1, case


def test_the_odd_values_are_the_triaged_ones():
    rotations = {c.rotation for c in odd_values.CASES} - {odd_values.CONTROL_ROTATION}
    formats = {c.format_id for c in odd_values.CASES} - {odd_values.CONTROL_FORMAT_ID}
    assert rotations == {181, 254, 256, 1000}
    assert all(not (0 <= r <= 180 or r == 255) for r in rotations)
    assert formats == {170, 300}
    assert all(f >= 164 for f in formats)  # the custom range: no built-in code


def test_a_piece_of_the_styles_must_be_there_once():
    with pytest.raises(ValueError, match="exactly once"):
        odd_values._patch(odd_values.base(), [('textRotation="7"', 'textRotation="8"')])


# --- reading Excel's answers


@pytest.mark.parametrize(
    ("orientation", "expected"),
    [
        ("orientation horizontal", (0, False)),
        ("orientation upward", (90, False)),
        ("orientation downward", (-90, False)),
        ("orientation vertical", (0, True)),
        ("45", (45, False)),
        ("-45", (-45, False)),
        ("45.0", (45, False)),
    ],
)
def test_excel_orientations_read_as_sheetpeek_rotations(orientation, expected):
    assert odd_values.degrees(orientation) == expected


def test_an_orientation_excel_does_not_give_is_refused():
    with pytest.raises(ValueError, match="not an orientation"):
        odd_values.degrees("ERR no such property")


def test_the_script_answers_become_rows():
    opened = odd_values.parse("control", "opened\tOdd value\t45\t#,##0.00\t1,234.50")
    assert opened == {
        "case": "control",
        "outcome": "opened",
        "A1": "Odd value",
        "orientation": "45",
        "number_format": "#,##0.00",
        "text": "1,234.50",
    }
    repaired = odd_values.parse("format_170", "repaired\tOdd value\t45\tGeneral\t1234.5")
    assert repaired["outcome"] == "repaired"
    refused = odd_values.parse("rotation_1000", "refused\t-50")
    assert refused == dict(
        zip(odd_values.COLUMNS, ["rotation_1000", "refused", "-50", "", "", ""], strict=True)
    )
    with pytest.raises(ValueError, match="not an answer"):
        odd_values.parse("control", "timed out after 60 s")


def test_written_rows_read_back(tmp_path):
    path = tmp_path / "odd_values.tsv"
    rows = _agreeing()
    odd_values.write(rows, "Excel 16.0 for Mac", path)
    lines = path.read_text(encoding="utf-8").splitlines()
    assert lines[0].startswith("# suite: odd_values (capture.odd_values); Excel 16.0 for Mac")
    assert lines[1].startswith("# captured: ")
    assert lines[2] == "\t".join(odd_values.COLUMNS)
    assert [dict(zip(odd_values.COLUMNS, line.split("\t"), strict=True)) for line in lines[3:]] == (
        rows
    )


# --- the capture (suite odd_values)


def _captured():
    path = odd_values.OUTPUT
    if not path.exists():
        pytest.skip("not captured yet: run python -m capture.odd_values run")
    lines = path.read_text(encoding="utf-8").splitlines()
    rows = [line.split("\t") for line in lines if not line.startswith("#")]
    header, body = rows[0], rows[1:]
    assert header == list(odd_values.COLUMNS)
    return [dict(zip(header, row + [""] * (len(header) - len(row)), strict=True)) for row in body]


# Known deviations, each with its reason; none yet.
DEVIATIONS: dict[str, str] = {}

# Tolerances: none. Every compared value is a whole angle, a format code or the text a cell
# shows: Excel answers them exactly, and sheetpeek either reads them as Excel does or not.


def _compare(rows):
    """Each odd case where sheetpeek reads it otherwise than Excel, with what each reads."""
    found = {}
    for row in rows[1:-1]:
        name = row["case"]
        rotation, stacked, number_format, text = odd_values.EXPECTED[name]
        if row["outcome"] == "refused":
            # Excel refusing the file is the answer; sheetpeek reads every case.
            found[name] = f"Excel refuses it ({row['A1']}); sheetpeek reads it"
            continue
        excel = (
            row["A1"],
            odd_values.degrees(row["orientation"]),
            row["number_format"],
            row["text"],
        )
        ours = (odd_values.TEXT, (rotation, stacked), number_format, text)
        if excel != ours:
            found[name] = {"excel": excel, "sheetpeek": ours}
    return found


@pytest.mark.excel
def test_every_case_is_captured_in_order():
    assert [r["case"] for r in _captured()] == [c.name for c in odd_values.CASES]


@pytest.mark.excel
def test_the_controls_open_and_read_as_built():
    rows = _captured()
    for control in (rows[0], rows[-1]):
        assert control["outcome"] == "opened", control
        assert control["A1"] == odd_values.TEXT
        assert odd_values.degrees(control["orientation"]) == (odd_values.CONTROL_ROTATION, False)
        assert control["number_format"] == odd_values.CONTROL_FORMAT
        assert control["text"] == odd_values.EXPECTED[control["case"]][3]


@pytest.mark.excel
def test_each_odd_value_is_read_as_excel_reads_it():
    found = _compare(_captured())
    assert {name: why for name, why in found.items() if name not in DEVIATIONS} == {}


def _agreeing():
    rows = []
    for case in odd_values.CASES:
        rotation, stacked, number_format, text = odd_values.EXPECTED[case.name]
        orientation = "orientation vertical" if stacked else str(rotation)
        if (rotation, stacked) == (0, False):
            orientation = "orientation horizontal"
        values = [case.name, "opened", odd_values.TEXT, orientation, number_format, text]
        rows.append(dict(zip(odd_values.COLUMNS, values, strict=True)))
    return rows


def test_the_comparison_accepts_a_capture_that_agrees_with_sheetpeek(tmp_path, monkeypatch):
    path = tmp_path / "odd_values.tsv"
    monkeypatch.setattr(odd_values, "OUTPUT", path)
    checks = (
        test_every_case_is_captured_in_order,
        test_the_controls_open_and_read_as_built,
        test_each_odd_value_is_read_as_excel_reads_it,
    )
    rows = _agreeing()
    odd_values.write(rows, "Excel", path)
    for check in checks:
        check()
    # and refuses one that does not: an odd angle kept, a refusal, an undefined id shown in a
    # format of its own, a control that did not open
    refused = dict(
        zip(odd_values.COLUMNS, ["rotation_256", "refused", "-50", "", "", ""], strict=True)
    )
    mutations = [
        ("rotation_181", {"orientation": "-1"}, test_each_odd_value_is_read_as_excel_reads_it),
        ("rotation_256", refused, test_each_odd_value_is_read_as_excel_reads_it),
        (
            "format_300",
            {"number_format": "0.00", "text": "1234.50"},
            test_each_odd_value_is_read_as_excel_reads_it,
        ),
        ("control_again", {"outcome": "repaired"}, test_the_controls_open_and_read_as_built),
        ("control", {"case": "first"}, test_every_case_is_captured_in_order),
    ]
    for name, change, check in mutations:
        changed = [{**r, **change} if r["case"] == name else r for r in rows]
        assert changed != rows, name
        odd_values.write(changed, "Excel", path)
        with pytest.raises(AssertionError):
            check()
