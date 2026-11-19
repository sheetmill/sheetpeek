"""The screen suite ``tab_strip``: its workbooks, its measuring, and the comparison.

ARCHITECTURE.md §4.15 (the tab strip, tab colors). The workbooks are built and read back with
sheetpeek; the measuring runs on a synthetic capture (``Canvas``) with a strip painted below
the cells; the comparison with ``tests/captures/tab_strip.tsv`` is skipped until the suite has
been captured.
"""

from __future__ import annotations

import itertools
import re
import zipfile
from pathlib import Path

import openpyxl
import pytest
from capture import tab_strip as suite
from capture.screen.run import write_tsv
from capture.screen.window import MARKER

import sheetpeek
from tests.capture_report import REPORT
from tests.screen_canvas import Canvas

CAPTURE = Path(__file__).with_name("captures") / "tab_strip.tsv"
FILL_TOLERANCE = 8
"""Per channel, for a tab's solid fill: the screen's color conversion moves saturated colors by
up to 6 (``00B050`` showed as ``06B050`` in the suite ``tables``)."""
GRADIENT_TOLERANCE = 12
"""Per channel, for the active tab's two ends: sampled 3 points inside the fill's top and 2 above
its bottom, not at the gradient's very ends, which sheetpeek's stops stand for."""
INK_TOLERANCE = 24
"""Per channel, for a tab's text color: the core of 11 to 12 pt strokes (about 2 device pixels
wide), where smoothing leaves more of the fill than in the 18 pt strokes of other suites."""
SAME_TOLERANCE = 4
"""Per channel, for two tabs Excel should draw alike (auto beside uncolored): the same pixels,
less the conversion's rounding."""
ORDER_TOLERANCE = 40
"""Per channel, for naming an ``order`` tab by its color: wide enough for a fill drawn a little
off its color, and the colors are twice as far apart, so a tab names one sheet only."""
GEOMETRY_TOLERANCE = 1.0
"""Points, for a tab's height, width and gap: each edge is anti-aliased over a device pixel."""
LABEL = 15.0
"""Points of ink at least as wide as a tab's name: the scroll arrows (6.5 to 7 pt) and the
new-sheet button (10) draw narrower ink, the narrowest name 20.5, and a gap between tabs none."""

# --- the workbooks


def test_every_group_builds_and_reads_back_as_built(tmp_path):
    placed = suite.build(tmp_path)
    assert len(placed) == sum(len(tabs) for tabs in suite.GROUPS.values())
    for group, tabs in suite.GROUPS.items():
        path = tmp_path / f"tab_strip_{group}.xlsx"
        with sheetpeek.open(path) as book:
            found = [(s.name, s.state) for s in book.sheets]
        states = {"veryHidden": "very_hidden"}
        assert found == [(t.name, states.get(t.state, t.state)) for t in tabs], group
        for appearance in suite.APPEARANCES:
            view = suite.sheetpeek_view(path, appearance)
            assert view == suite.EXPECTED[appearance][group], (appearance, group)


def test_the_order_group_sheet_ids_differ_from_the_file_order(tmp_path):
    suite.build(tmp_path)
    with zipfile.ZipFile(tmp_path / "tab_strip_order.xlsx") as package:
        workbook = package.read("xl/workbook.xml").decode()
    ids = [int(n) for n in re.findall(r'sheetId="(\d+)"', workbook)]
    assert ids == list(suite.SHEET_IDS)
    assert ids != sorted(ids)
    assert 'activeTab="0"' in workbook


def test_the_precedence_group_carries_two_kinds_in_each_tab_color(tmp_path):
    suite.build(tmp_path)
    with zipfile.ZipFile(tmp_path / "tab_strip_precedence.xlsx") as package:
        for k, tab in enumerate(suite.PRECEDENCE, 1):
            sheet = package.read(f"xl/worksheets/sheet{k}.xml").decode()
            found = re.findall(r"<tabColor [^>]*/>", sheet)
            assert found == ([tab.xml] if tab.xml else []), tab.name
    with sheetpeek.open(tmp_path / "tab_strip_precedence.xlsx") as book:
        colors = [s.tab_color for s in book.sheets]
    # theme before rgb, rgb before indexed, rgb over auto, as Excel draws them (§4.15, captured)
    assert colors == [None, "C00000", "C0504D", "0070C0", "C0504D"]


def test_the_first_sheet_is_active_visible_and_holds_the_markers():
    for group in suite.GROUPS:
        wb = openpyxl.Workbook()
        suite.fill(wb.active, group)
        assert wb.active is wb.worksheets[0]
        assert wb.active.sheet_state == "visible"
        assert wb.active.title == suite.GROUPS[group][0].name
        for marker in suite.MARKERS:  # the strip is found down column A from the first
            assert wb.active[marker].value is None
        assert all(wb.active.cell(r, 1).value is None for r in range(1, 50))


def test_no_tab_draws_a_color_near_the_marker():
    marker = [int(MARKER[i : i + 2], 16) for i in (0, 2, 4)]
    drawn = {
        color
        for groups in suite.EXPECTED.values()
        for tabs in groups.values()
        for tab in tabs
        for color in tab[1:]
    }
    for color in drawn:
        rgb = [int(color[i : i + 2], 16) for i in (0, 2, 4)]
        assert max(abs(a - b) for a, b in zip(rgb, marker, strict=True)) > 96, color


def test_every_order_tab_color_tells_its_sheet():
    colors = [t.color.rgb[2:] for t in suite.ORDER if t.color is not None]
    assert len(set(colors)) == len(suite.ORDER) - 1
    for a in colors:
        for b in colors:
            if a != b:
                assert not _near(a, b, 2 * ORDER_TOLERANCE), (a, b)


# --- the measuring, on a synthetic capture

TOP = 1100
BACKGROUND = (50, 50, 50)


def _hex(rgb):
    return "".join(f"{c:02X}" for c in rgb)


def _gradient(row):
    top, bottom = (230, 241, 249), (153, 198, 230)
    return tuple(round(a + (b - a) * row / 43) for a, b in zip(top, bottom, strict=True))


def _strip():
    """A canvas with a dark strip at ``TOP``: arrows, an active tab, a red tab, two uncolored
    tabs, the new-sheet button; device pixels at 2 a point."""
    ws = openpyxl.Workbook().active
    suite.fill(ws, "colors")
    canvas = Canvas(ws, [*suite.MARKERS, "Z40"])
    canvas.paint(0, TOP, canvas.width, TOP + 52, BACKGROUND)
    canvas.paint(0, TOP, canvas.width, TOP + 2, (99, 99, 99))  # the strip's top line
    canvas.paint(20, 1120, 28, 1130, (160, 160, 160))  # arrows
    canvas.paint(60, 1120, 68, 1130, (160, 160, 160))
    for y in range(TOP, TOP + 44):  # the active tab, 78 x 22 pt, a vertical gradient
        canvas.paint(120, y, 276, y + 1, _gradient(y - TOP))
    canvas.paint(170, 1116, 226, 1130, (39, 132, 82))
    canvas.paint(290, TOP + 2, 434, TOP + 42, (192, 0, 0))  # 72 x 20 pt, 7 pt after
    canvas.paint(330, 1116, 380, 1130, (255, 255, 255))
    canvas.paint(470, 1116, 520, 1130, (187, 187, 187))  # two uncolored tabs
    canvas.paint(640, 1116, 690, 1130, (187, 187, 187))
    canvas.paint(780, 1120, 792, 1132, (160, 160, 160))  # the new-sheet button
    return canvas


def test_the_strip_is_found_below_the_sheet():
    assert suite.strip_top(_strip().capture()) == TOP
    ws = openpyxl.Workbook().active
    suite.fill(ws, "colors")
    assert suite.strip_top(Canvas(ws, [*suite.MARKERS, "Z40"]).capture()) is None


def test_the_strip_reads_as_items_left_to_right():
    records = suite.measure(_strip().capture(), "colors", "dark")
    assert all(set(suite.COLUMNS) >= set(r) for r in records)
    assert [r["item"] for r in records] == [str(k) for k in range(1, len(records) + 1)]
    named = _tabs(records)
    assert [(r["kind"], r["left"], r["width"]) for r in named] == [
        ("fill", "60", "78"),
        ("fill", "145", "72"),
        ("ink", "217", "72.5"),  # from the red tab to the middle of the gap
        ("ink", "289.5", "77.5"),
    ]
    active, red, plain, bare = named
    assert (active["height"], red["height"]) == ("22", "20")
    # 3 points below the top and 2 above the bottom: rows 6 and 39 of the gradient's 44
    assert (active["top_color"], active["bottom_color"]) == (
        _hex(_gradient(6)),
        _hex(_gradient(39)),
    )
    assert active["ink_color"] == "278452"
    assert (red["top_color"], red["bottom_color"], red["ink_color"]) == ("C00000",) * 2 + (
        "FFFFFF",
    )
    assert (red["ink_w"], red["ink_h"]) == ("25", "7")
    assert plain["kind"] == bare["kind"] == "ink"
    assert plain["ink_color"] == bare["ink_color"] == "BBBBBB"
    assert plain["ink"] == bare["ink"]  # the same mark: the same bitmap
    assert float(plain["left"]) + float(plain["width"]) == float(bare["left"])
    small = [r for r in records if r.get("ink_w") and float(r["ink_w"]) < LABEL]
    assert [r["ink_w"] for r in small] == ["4", "4", "6"]  # the arrows and the button
    profile = active["profile"].split()
    assert len(profile) == suite.BAND
    assert profile[-1] == "323232"  # below the tab's 22 points, the strip


def test_a_strip_without_a_found_top_is_one_row_saying_so():
    ws = openpyxl.Workbook().active
    suite.fill(ws, "order")
    records = suite.measure(Canvas(ws, [*suite.MARKERS, "Z40"]).capture(), "order", "light")
    assert records == [{"appearance": "light", "group": "order", "note": "no strip found"}]


def test_write_takes_appearances_groups_and_refusals(tmp_path, monkeypatch):
    monkeypatch.setattr(suite, "BUILD_DIR", tmp_path)
    monkeypatch.setattr(suite, "OUTPUT", tmp_path / "out.tsv")
    suite.record_refusal("dark", "precedence", "osascript: execution error: refused (-50)")
    (tmp_path / "light" / "order").mkdir(parents=True)
    write_tsv(
        tmp_path / "light" / "order" / "formats.tsv",
        {"excel": "16.115", "appearance": "light", "suite": "x"},
        [],
    )
    suite.write()
    lines = (tmp_path / "out.tsv").read_text(encoding="utf-8").splitlines()
    assert lines[:2] == ["# excel: 16.115", "# suite: tab_strip"]
    assert lines[2].split("\t") == list(suite.COLUMNS)
    rows = [dict(zip(suite.COLUMNS, line.split("\t"), strict=True)) for line in lines[3:]]
    assert [(r["appearance"], r["group"], r["note"]) for r in rows] == [
        ("dark", "precedence", "refused: refused (-50)")
    ]


# --- the comparison with Excel


def _near(a, b, tolerance):
    return all(abs(int(a[i : i + 2], 16) - int(b[i : i + 2], 16)) <= tolerance for i in (0, 2, 4))


def _tabs(records):
    """The items that are tabs: ink as wide as a name."""
    return [r for r in records if r.get("ink_w") and float(r["ink_w"]) >= LABEL]


def _row_color(record):
    """A tab's color in the row the strip is split at (its fill, or the strip's background)."""
    return record["top_color"] or record["profile"].split()[suite.ROW]


@pytest.fixture(scope="module")
def captured():
    if not CAPTURE.exists():
        pytest.skip("not captured yet: tests/captures/tab_strip.tsv")
    lines = [ln for ln in CAPTURE.read_text(encoding="utf-8").splitlines() if ln[:2] != "# "]
    names = lines[0].split("\t")
    rows = [dict(zip(names, line.split("\t"), strict=False)) for line in lines[1:]]
    refused = {(r["appearance"], r["group"]): r["note"] for r in rows if r["note"]}
    if refused:
        REPORT["tab_strip refused"] = "tab_strip: " + "; ".join(
            f"{a} {g} {note}" for (a, g), note in refused.items()
        )
    found: dict[tuple[str, str], list[dict[str, str]]] = {}
    for row in rows:
        if (row["appearance"], row["group"]) not in refused:
            found.setdefault((row["appearance"], row["group"]), []).append(row)
    return {key: _tabs(records) for key, records in found.items()}


@pytest.mark.excel
def test_the_strip_shows_the_visible_sheets_in_file_order(captured):
    # Hidden and very hidden sheets have no tab; the tabs follow the file's order, not the
    # sheetId order. Every inactive tab names its sheet by its color.
    colors = {t.color.rgb[2:]: t.name for t in suite.ORDER if t.color is not None}
    runs = [(a, tabs) for (a, g), tabs in captured.items() if g == "order"]
    if not runs:
        pytest.skip("group order not captured")
    for appearance, tabs in runs:
        expected = [name for name, *_ in suite.EXPECTED[appearance]["order"]]
        assert len(tabs) == len(expected), appearance
        named = []
        for tab in tabs[1:]:
            color = _row_color(tab)
            hits = [name for rgb, name in colors.items() if _near(color, rgb, ORDER_TOLERANCE)]
            named.append(hits[0] if len(hits) == 1 else f"? {color}")
        assert named == expected[1:], appearance


@pytest.mark.excel
def test_tab_colors_as_sheetpeek_draws_them(captured):
    # An inactive colored tab is its color, solid, with white text; the active tab runs from
    # sheetpeek's first stop to its second, with Excel's green text; the precedence group's
    # tabs show the kind sheetpeek takes. An uncolored inactive tab is the page's scheme, which
    # is reported, not asserted: it is a page's choice, not Excel's color.
    uncolored = []
    compared = 0
    for (appearance, group), tabs in captured.items():
        if group == "order":
            continue
        expected = suite.EXPECTED[appearance][group]
        assert len(tabs) == len(expected), (appearance, group)
        for k, (tab, (name, top, bottom, ink)) in enumerate(zip(tabs, expected, strict=True)):
            where = (appearance, group, name)
            if k == 0:
                assert _near(tab["top_color"], top, GRADIENT_TOLERANCE), where
                assert _near(tab["bottom_color"], bottom, GRADIENT_TOLERANCE), where
                assert _near(tab["ink_color"], ink, INK_TOLERANCE), where
            elif ink in ("FFFFFF", "000000"):  # a colored tab, its name white or black
                assert tab["kind"] == "fill", where
                assert _near(tab["top_color"], top, FILL_TOLERANCE), where
                assert _near(tab["bottom_color"], bottom, FILL_TOLERANCE), where
                assert _near(tab["ink_color"], ink, INK_TOLERANCE), where
            else:
                uncolored.append(f"{appearance} {name} {_row_color(tab)}/{tab['ink_color']}")
                continue
            compared += 1
    REPORT["tab_strip colors"] = (
        f"tab_strip: {compared} tabs drawn in sheetpeek's colors; uncolored tabs on screen "
        f"(fill/text): {', '.join(uncolored)}"
    )


@pytest.mark.excel
def test_auto_draws_an_uncolored_tab(captured):
    # Auto sits between Indexed and Bare (uncolored), Plain is uncolored too: Auto is drawn as
    # they are.
    runs = [(a, tabs) for (a, g), tabs in captured.items() if g == "colors"]
    if not runs:
        pytest.skip("group colors not captured")
    names = [name for name, *_ in suite.EXPECTED["dark"]["colors"]]
    for appearance, tabs in runs:
        by = dict(zip(names, tabs, strict=True))
        auto = by["Auto"]
        for other in (by["Plain"], by["Bare"]):
            assert auto["kind"] == other["kind"], appearance
            assert _near(_row_color(auto), _row_color(other), SAME_TOLERANCE), appearance
            assert _near(auto["ink_color"], other["ink_color"], SAME_TOLERANCE), appearance


@pytest.mark.excel
def test_tab_size_and_spacing(captured):
    # sheetpeek's tab is 72 x 20 pt, 7 pt apart: measured on the filled tabs next to each other
    # (Rgb, Theme, Tint, Indexed), whose edges show against the strip.
    runs = [(a, tabs) for (a, g), tabs in captured.items() if g == "colors"]
    if not runs:
        pytest.skip("group colors not captured")
    width, height = suite.TAB_SIZE
    for appearance, tabs in runs:
        filled = [t for t in tabs[2:6] if t["kind"] == "fill"]
        assert len(filled) == 4, appearance
        for tab in filled:
            assert abs(float(tab["height"]) - height) <= GEOMETRY_TOLERANCE, appearance
            assert float(tab["width"]) >= width - GEOMETRY_TOLERANCE, appearance
        for a, b in itertools.pairwise(filled):
            gap = float(b["left"]) - float(a["left"]) - float(a["width"])
            assert abs(gap - suite.TAB_GAP) <= GEOMETRY_TOLERANCE, appearance
