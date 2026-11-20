"""The capture suite ``rotation_placement`` (angled and stacked text at more angles, rows and
sizes, ARCHITECTURE.md §4.17), without Excel where it can be.

The workbook is built and read back with sheetpeek. The comparisons with
``tests/captures/rotation_placement.tsv`` model each case's ink from sheetpeek's own layout, as
``tests/test_suite_rotation_more.py`` does, and skip until the capture exists. The cases that
differ are named in ``PINNED`` with their reasons, so a new difference fails and so does a fixed
one.
"""

import math

import pytest
from capture import rotation_more as rm
from capture import rotation_placement as suite

import sheetpeek
from tests.test_rotation import _TOPS

INK_BOUND = 1.5
"""Points, as in ``tests/test_suite_rotation_more.py``."""


@pytest.fixture(scope="module")
def laid_out(tmp_path_factory):
    """Each case's runs (x, baseline, width and angle from its cell's left and top) and width."""
    path = tmp_path_factory.mktemp("rotation_placement") / "rotation_placement.xlsx"
    rm.build(path, suite.CASES)
    book = sheetpeek.open(path)
    found = {}
    for name, cell, case in rm.placements(suite.CASES):
        rendered = book.sheet(name).layout("excel")
        r = int(cell[1:])
        row = next(x for x in rendered.rows if x.index == r)
        column = next(c for c in rendered.columns if c.index == 5)
        runs = [run for x in rendered.rows for run in x.runs if run.source == (r, 5)]
        found[case.name] = {
            "runs": [(run.x - column.x_pt, run.y - row.y_pt, run.width, run.angle) for run in runs],
            "width": column.width_pt,
            "height": row.height_pt,
            "case": case,
        }
    return found


def test_the_workbook_reads_back_as_built(tmp_path):
    path = tmp_path / "rotation_placement.xlsx"
    placed = rm.build(path, suite.CASES)
    assert [case for _, _, case in placed] == suite.CASES
    assert len({c.name for c in suite.CASES}) == len(suite.CASES)
    book = sheetpeek.open(path)
    for name, cell, case in placed:
        style = book.sheet(name).cell(int(cell[1:]), 5).style
        stored = case.cell.rotation
        assert style.stacked is (stored == 255), case.name
        if stored != 255:
            assert style.rotation == (stored if stored <= 90 else 90 - stored), case.name
        assert style.h_align == (case.cell.horizontal or "general"), case.name
        assert style.v_align == (case.cell.vertical or "bottom"), case.name
        assert bool(style.borders) is (case.group == "top" and "lrtb" in case.name), case.name


def test_the_cases_cover_the_rules():
    spread = {c.name for c in suite.CASES if c.group == "spread"}
    assert {"-90 fill", "-90 justify", "-90 distributed", "-90 left"} <= spread
    stacked = [c for c in suite.CASES if c.group == "stacked"]
    assert {c.cell.size for c in stacked} == {9, 11, 12, 14, 16, 20, 24, 28, 36}
    assert {c.cell.width for c in stacked} == {rm.STACKED_WIDTH}


def _ink(entry):
    """The modeled ink of a one-run case: left, right past the far edge, top and bottom."""
    ((x, y, w, angle),) = entry["runs"]
    case = entry["case"]
    a = math.radians(angle)
    cos, sin = math.cos(a), math.sin(a)
    size, text = case.cell.size, case.cell.text
    points = [
        (x + s * cos - t * sin, y - s * sin - t * cos)
        for s, letter in ((0.8, text[0]), (w - 0.8, text[-1]))
        for t in (0.0, _TOPS.get(letter, 0.68) * size)
    ]
    xs, ys = [p[0] for p in points], [p[1] for p in points]
    return min(xs), max(xs) - entry["width"], min(ys), max(ys)


def _captured():
    path = suite.OUTPUT
    if not path.exists():
        pytest.skip("not captured yet: run python -m capture.rotation_placement run")
    lines = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
    rows = [line for line in lines if line[0] and not line[0].startswith("#")]
    header, body = rows[0], rows[1:]
    return {r["name"]: r for r in (dict(zip(header, b, strict=False)) for b in body)}


def _names(text):
    return frozenset(text.split("|"))


PINNED: dict[str, frozenset[str]] = {
    # Fill stops at the cell's edge in Excel, which the ink model does not cut (30 and -15: 20 and
    # 31 pt); right placement at -15 sits 1.9 pt from sheetpeek's and fill and justify at -90 1.55
    # pt (its insets there); in 30 pt rows a box at 20 degrees is taller than the row and clipped.
    "across": _names("-15 distributed|-15 fill|-15 justify|-90 fill|-90 justify|20 top 30|30 fill"),
    # Bottom-aligned angled text sits 1.2 to 4.4 pt higher in Excel at 30 to 85 and -75 to -90
    # degrees, not yet modeled; fill's cut edge also cuts its top; boxes at 15 and 20 degrees in
    # 30 pt rows are clipped.
    "height": _names(
        "-15 fill|-75 fill|-75 left|-90 distributed|-90 fill|-90 justify|-90 left|15 top 30|"
        "20 top 30|30 fill|30 left|60 distributed|60 fill|60 justify|60 left|85 distributed|"
        "85 fill|85 justify|85 left"
    ),
    # Top alignment between slanted sides at 15 to 20 degrees and at 50 lies 1.5 to 2.1 pt across
    # from sheetpeek's (8 to 13 pt where the box is taller than its 30 pt row).
    "between": _names(
        "15 lrtb top 30|15 lrtb top 45|20 lrtb top 30|20 lrtb top 50|20 lrtb top 70|"
        "20 lrtb top 90|50 lrtb top 90"
    ),
}
"""Cases where Excel 16.115 differs from sheetpeek, by comparison, with their reasons."""


def _check(key, wrong):
    assert set(wrong) == PINNED.get(key, frozenset()), (key, sorted(wrong))


def _angled(group, slanted):
    return [
        c
        for c in suite.CASES
        if c.group == group and c.cell.rotation != 255 and ("lrtb" in c.name) == slanted
    ]


@pytest.mark.excel
def test_angled_text_sits_across_where_sheetpeek_puts_it(laid_out):
    """§4.17: spread alignments by the angle's sign, and top alignment without sides."""
    found = _captured()
    wrong = {}
    for case in _angled("spread", False) + _angled("top", False):
        record = found[case.name]
        left, right, _, _ = _ink(laid_out[case.name])
        off = (float(record["ink_left"]) - left, float(record["ink_right"]) - right)
        if max(map(abs, off)) > INK_BOUND:
            wrong[case.name] = tuple(round(o, 2) for o in off)
    _check("across", wrong)


@pytest.mark.excel
def test_angled_text_sits_in_height_where_sheetpeek_puts_it(laid_out):
    """§4.17: how high bottom- and top-aligned angled text sits, without sides."""
    found = _captured()
    wrong = {}
    for case in _angled("spread", False) + _angled("top", False):
        record = found[case.name]
        _, _, top, bottom = _ink(laid_out[case.name])
        off = (float(record["ink_top"]) - top, float(record["ink_bottom"]) - bottom)
        if max(map(abs, off)) > INK_BOUND:
            wrong[case.name] = tuple(round(o, 2) for o in off)
    _check("height", wrong)


@pytest.mark.excel
def test_top_aligned_text_between_slanted_sides_as_captured(laid_out):
    """§4.17: the highest corner 3 pt below the top, across the line between the sides."""
    found = _captured()
    wrong = {}
    for case in _angled("top", True):
        record, entry = found[case.name], laid_out[case.name]
        left, right, _, _ = _ink(entry)
        across = abs(math.sin(math.radians(entry["runs"][0][3])))
        off = (
            (float(record["ink_left"]) - left) * across,
            (float(record["ink_right"]) - right) * across,
        )
        if max(map(abs, off)) > INK_BOUND:
            wrong[case.name] = tuple(round(o, 2) for o in off)
    _check("between", wrong)


@pytest.mark.excel
def test_stacked_characters_sit_where_sheetpeek_puts_them(laid_out):
    """§4.17: each line's ink center against its character's center, as rotation_more compares
    it; "WiW"'s "i" draws two bands, its dot and its stem, each centered on the "i"."""
    found = _captured()
    wrong = {}
    for case in (c for c in suite.CASES if c.group == "stacked"):
        record, entry = found[case.name], laid_out[case.name]
        lines = [tuple(map(float, line.split(":"))) for line in record["lines"].split(";")]
        runs = entry["runs"]
        if case.cell.text == "WiW":
            runs = [runs[0], runs[1], runs[1], runs[2]]
        offsets = [
            (ink_left + entry["width"] + ink_right) / 2 - (x + w / 2)
            for (ink_left, ink_right), (x, _, w, _) in zip(lines, runs, strict=True)
        ]
        if max(map(abs, offsets)) > INK_BOUND:
            wrong[case.name] = round(max(offsets, key=abs), 2)
    _check("stacked", wrong)
