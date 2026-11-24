"""Captures from Excel for Windows (``tests/captures/windows/``), against sheetpeek's layout with
``platform="windows"``.

Each file is the Mac suite's schema, read from a capture job's result by ``capture/job.py``; its
header names the job and holds the result's header. A comparison is skipped until its file exists.
"""

from pathlib import Path

import pytest

import sheetpeek
from tests.capture_report import REPORT

WINDOWS = Path(__file__).with_name("captures") / "windows"


def _rows(name):
    path = WINDOWS / name
    if not path.exists():
        pytest.skip(f"not captured yet: {path.name}")
    lines = path.read_text(encoding="utf-8").splitlines()
    head = [line for line in lines if line.startswith("# ")]
    assert "# platform: windows" in head
    return [line.split("\t") for line in lines if line and not line.startswith("# ")][1:]


@pytest.fixture(scope="module")
def cells(tmp_path_factory):
    pytest.importorskip("openpyxl")
    from capture.cell_text import build

    path = build(tmp_path_factory.mktemp("cell_text") / "cell_text.xlsx")
    return sheetpeek.open(path, platform="windows")


@pytest.mark.excel
def test_merge_width_matches_excel_for_windows(cells):
    rows: dict[tuple[str, bool], list[str]] = {}
    for width, merge, _value, _fmt, shown in _rows("merge_width.tsv"):
        rows.setdefault((width, merge == "A:C, B hidden"), []).append(shown)
    compared, deviations = 0, {}
    for (width, hidden), texts in rows.items():
        rendered = cells.sheet(("h" if hidden else "w") + width).layout("excel")
        for row, shown in enumerate(texts, 1):
            compared += 1
            if rendered.cell(row, 1).text != shown:
                deviations[(width, hidden, row)] = (rendered.cell(row, 1).text, shown)
    REPORT["merge width (Windows)"] = (
        f"merge width, Excel for Windows: {compared} cells compared, {len(deviations)} deviations"
    )
    assert compared == 150
    assert deviations == {}


@pytest.mark.excel
def test_error_width_matches_excel_for_windows(cells):
    from capture.cell_text import ERROR_WIDTHS, ERRORS

    rendered = cells.sheet("errors").layout("excel")
    records = _rows("error_width.tsv")
    deviations = {}
    for error, width, shown in records:
        cell = rendered.cell(ERRORS.index(error) + 1, ERROR_WIDTHS.index(float(width)) + 2)
        if cell.text != shown:
            deviations[(error, width)] = (cell.text, shown)
    REPORT["error width (Windows)"] = (
        f"error width, Excel for Windows: {len(records)} cells compared, {len(deviations)} "
        "deviations"
    )
    assert len(records) == 56
    assert deviations == {}


# --- borders and fills: the Mac's lines and tiles in device pixels


def _table(name):
    path = WINDOWS / name
    if not path.exists():
        pytest.skip(f"not captured yet: {path.name}")
    lines = path.read_text(encoding="utf-8").splitlines()
    head = [line for line in lines if line.startswith("# ")]
    rows = [line.split("\t") for line in lines if line and not line.startswith("# ")]
    return head, [dict(zip(rows[0], r, strict=True)) for r in rows[1:]]


def _px(length, zoom):
    """A length in sheet points at ``zoom`` as Windows device pixels (96 dpi)."""
    px = float(length) * 96 / 72 * int(zoom) / 100
    assert abs(px - round(px)) < 1e-6, (length, zoom)
    return round(px)


@pytest.mark.excel
def test_the_gridline_gray_is_the_one_recorded_for_windows():
    from capture.screen.window import GRIDLINE

    for name in ("borders.tsv", "fills.tsv"):
        head, _ = _table(name)
        (line,) = [x for x in head if x.startswith("# gridline: ")]
        gray = line.split()[2]
        assert gray == "".join(f"{c:02X}" for c in GRIDLINE["windows"]) == "E0E0E0"


@pytest.mark.excel
def test_windows_dashes_have_sheetpeeks_lengths_in_device_pixels():
    # As the Mac's test (tests/test_borders.py): every dashed side, dash and gap lengths exact in
    # the platform's unit (a device pixel here, where the Mac's is a screen point), the starting
    # point free; each device row against its own row of the tile (slantDashDot's two differ:
    # the row before the gridline is the tile's first, the row on it its third).
    from tests.test_borders import _fits_cycle, _tile_runs

    _, body = _table("borders.tsv")
    checked, wrong = 0, []
    for row in body:
        style, pattern = row["style"], row["pattern"]
        if style == "none" or pattern in ("solid", "empty", ""):
            continue
        tokens = pattern.removeprefix("irregular: ").split()
        measured = [
            (kind == "on", _px(length, row["zoom"]))
            for kind, length in zip(tokens[::2], tokens[1::2], strict=True)
        ]
        tile_row = 0 if float(row["offset"]) < 0 else 2
        checked += 1
        if not _fits_cycle(measured, _tile_runs(style, "windows", tile_row)):
            wrong.append((row["zoom"], style, row["edge"], row["offset"], pattern))
    REPORT["borders (Windows)"] = (
        f"borders, Excel for Windows: {checked} dashed sides compared, {len(wrong)} deviations"
    )
    assert checked == 104  # one device row a unit, where the Mac has two; slantDashDot's 16
    assert wrong == []


@pytest.mark.excel
def test_windows_lines_sit_where_the_macs_do_in_device_pixels():
    # Each style's device rows against its gridline, in Windows pixels, are the Mac's in screen
    # points: the Mac's two device rows a screen point read as one row each.
    mac = (WINDOWS.parent / "borders.tsv").read_text(encoding="utf-8").splitlines()
    rows = [x.split("\t") for x in mac if x and not x.startswith("#")]
    mac_body = [dict(zip(rows[0], r, strict=True)) for r in rows[1:]]
    _, body = _table("borders.tsv")

    def offsets(records, unit):
        found: dict[tuple[str, str, str], set[int]] = {}
        for r in records:
            n = float(r["offset"]) * int(r["zoom"]) / 100 / unit
            found.setdefault((r["zoom"], r["style"], r["edge"]), set()).add(int(n // 1))
        return found

    assert offsets(body, 0.75) == offsets(mac_body, 1.0)


@pytest.mark.excel
def test_windows_fill_tiles_are_sheetpeeks_in_device_pixels():
    from sheetpeek._fill_border_model import TILES

    def turns(tile):
        rows = tile.split("/")
        h, w = len(rows), len(rows[0])
        return {
            "/".join("".join(rows[(y + dy) % h][(x + dx) % w] for x in range(w)) for y in range(h))
            for dy in range(h)
            for dx in range(w)
        }

    _, body = _table("fills.tsv")
    compared, wrong = 0, []
    for row in body:
        if row["pattern"] not in TILES:
            continue
        compared += 1
        tile = TILES[row["pattern"]]
        # the measured tile repeats the pattern to a whole number of units on each axis
        h, w = len(row["tile_sheet"].split("/")), len(row["tile_sheet"].split("/")[0])
        ours = "/".join(
            "".join(tile[y % len(tile)][x % len(tile[0])] for x in range(w)) for y in range(h)
        )
        if row["tile_unit"] != "device_px" or row["tile_sheet"] not in turns(ours):
            wrong.append((row["zoom"], row["pattern"], row["cell"], row["tile_sheet"]))
    REPORT["fills (Windows)"] = (
        f"fills, Excel for Windows: {compared} pattern tiles compared, {len(wrong)} deviations"
    )
    assert compared == 74  # 17 patterns in two color sets and the 3 anchor cells, 2 zooms
    assert wrong == []


# --- text_steps: indents and shrink to fit in device pixels


@pytest.fixture(scope="module")
def steps_built(tmp_path_factory):
    pytest.importorskip("openpyxl")
    from capture import text_steps

    return text_steps.build(tmp_path_factory.mktemp("text_steps"))


def _steps_drawn(path, group):
    """sheetpeek's runs for each case with platform="windows": the first run's left and the last
    run's right from the cell's edges, in points, and the run's scale."""
    from capture import text_steps
    from openpyxl.utils import column_index_from_string
    from openpyxl.utils.cell import coordinate_from_string

    rendered = sheetpeek.read(path, platform="windows").layout("excel")
    columns = {c.index: c for c in rendered.columns}
    found = {}
    for case in text_steps.cases(group):
        letters, r = coordinate_from_string(case.cell)
        c = column_index_from_string(letters)
        row = next(x for x in rendered.rows if x.index == r)
        runs = [t for t in row.runs if t.source == (r, c)]
        left = columns[c].x_pt
        found[case.cell] = (
            min(t.x for t in runs) - left,
            max(t.x + t.width for t in runs) - left - columns[c].width_pt,
            runs[0].scale,
        )
    return found


@pytest.mark.excel
def test_windows_indents_move_text_as_sheetpeek_moves_it(steps_built):
    # As the Mac's test: the ink's edge at each level less its edge at level 0, against
    # sheetpeek's runs, within one device pixel (0.75 pt).
    from capture import text_steps

    _, body = _table("text_steps.tsv")
    captured = {(r["group"], r["cell"]): r for r in body}
    compared, wrong = 0, {}
    for group in text_steps.GROUPS:
        if not group.startswith("indent"):
            continue
        drawn = _steps_drawn(steps_built[group], group)
        found = text_steps.cases(group)
        for case in found:
            base = found[found.index(case) - int(case.name.split()[-1]) * 2]
            side, key = (0, "ink_left") if " left " in case.name else (1, "ink_right")
            moved = float(captured[(group, case.cell)][key]) - float(
                captured[(group, base.cell)][key]
            )
            expected = drawn[case.cell][side] - drawn[base.cell][side]
            compared += 1
            if abs(moved - expected) > 0.75:
                wrong[(group, case.name)] = (moved, expected)
    REPORT["indent (Windows)"] = (
        f"indent, Excel for Windows: {compared} cases compared, {len(wrong)} deviations"
    )
    assert compared == 108
    assert wrong == {}


KNOWN_SHRINK: dict[str, tuple[int, ...]] = {
    # Known deviations, the lengths by group, one reason for all: sheetpeek draws Windows' shrink
    # at the Mac's whole-point levels (every 4/3 px) where Excel for Windows draws whole pixels,
    # never 6, 2 the smallest.
    "shrink": (8, 9, 10, 11, 12, *range(16, 31)),
    "shrink_calibri16": (8, 9, 10, 12, 13, 14, 15, 16, 17, 18, *range(24, 52)),
    "shrink_aptos_narrow11": (8, 9, 10, 11, 12, *range(16, 52)),
}


@pytest.mark.excel
def test_windows_shrinks_each_length_to_the_pixel_level_sheetpeek_gives_it(steps_built):
    # The "H" advance Excel draws each length at, in device pixels: its ink width plus the
    # unshrunk text's side bearings (from the group's references), over its length; in Calibri
    # 11's group, the ink width over each length's unshrunk copy's times the unshrunk advance
    # (xlformat's, 9 px). A text clipped at the cell's right edge (ink to its last pixel) has no
    # measurable advance: sheetpeek must draw it at the smallest level, 2 px, as nothing fits.
    import xlformat
    from capture import text_steps

    _, body = _table("text_steps.tsv")
    captured = {(r["group"], r["cell"]): r for r in body}

    def ink(record):
        return float(record["cell_width"]) - float(record["ink_left"]) + float(record["ink_right"])

    compared, clipped, wrong = 0, 0, {}
    for group in ("shrink", "shrink_calibri16", "shrink_aptos_narrow11"):
        drawn = _steps_drawn(steps_built[group], group)
        if group == "shrink":
            font = xlformat.Font("Calibri", 11)
            unshrunk = round(xlformat.text_width("H", font=font, platform="windows") * 4 / 3)
        else:
            short, long = (ink(captured[(group, c)]) for c in text_steps.REFERENCES)
            unshrunk = round((long - short) / 10 * 4 / 3)
            bearings = (long - short) - short
        for case in text_steps.cases(group):
            if not case.name.startswith("shrunk"):
                continue
            record = captured[(group, case.cell)]
            ours = drawn[case.cell][2] * unshrunk
            compared += 1
            if float(record["ink_right"]) == 0:
                clipped += 1
                if round(ours, 6) != 2:
                    wrong[(group, case.name)] = ("clipped", ours)
                continue
            if group == "shrink":
                level = round(float(record["width_ratio"]) * unshrunk)
            else:
                n = int(case.name.split()[1])
                level = round((ink(record) + bearings) * 4 / 3 / n)
            if level != round(ours, 6):
                wrong[(group, case.name)] = (level, ours)
    REPORT["shrink (Windows)"] = (
        f"shrink, Excel for Windows: {compared} lengths compared ({clipped} clipped), "
        f"{len(wrong)} deviations, known (whole-point levels)"
    )
    assert compared == 27 + 48 + 48
    known = {(group, f"shrunk {n}") for group, lengths in KNOWN_SHRINK.items() for n in lengths}
    assert set(wrong) == known
