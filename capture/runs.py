"""Capture how Excel for Mac draws rich-text runs, from the screen (route C).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.runs run text_and_runs
    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.runs run
    uv run --group capture python -m capture.runs read           # the saved captures again
    uv run --group capture python -m capture.runs build PATH     # every group's workbook

ARCHITECTURE.md §4.16. Each group is its own workbook and window, at zoom 100, in the dark
appearance as the other suites. A cell's runs are written as Excel writes them, in a shared
string (``<si>`` with ``<r>`` elements, each ``<rPr>`` holding only the properties the case
names); the workbook is saved by openpyxl and its strings moved into ``xl/sharedStrings.xml``
before Excel opens it. Most texts are words of "HH" one space apart, each run starting at a word,
so that each word is in one font and its ink tells the size (cap height), the weight (ink area)
and the color drawn. Groups:

- ``main`` (Normal font Calibri 11). Column C keeps the default width, column J is 30
  characters wide. Every case has a row of its own.

  - C4 "a long text that spills with a LARGE run inside", "LARGE" a bold 16 pt run, and C6 the
    same text in one font: where the ink ends past C's right edge and which vertical gridlines
    from C to I the text hides (the widths add per run; the spill reach).
  - C8, C10, C12 "aaa BBB ccc ddd" wrapped, "BBB" a 14, 20 and 28 pt run: the fitted row height
    and each line's ink band (no letter descends, so a band's bottom is the line's baseline):
    each line takes the height of its largest run, lines drawn their largest run's pitch apart.
  - C14 a 20 pt cell whose runs are all 9 pt ("aa" and a bold "bb"): the fitted row height.
  - J16 a 16 pt red cell: "HH" a run without properties, "HH" a run whose properties hold only
    bold: a run without properties takes the cell's font, one with properties takes nothing
    from it.
  - J18 "aa BIG aa" in an 11 pt cell, "BIG" a 20 pt run, and J20 the same text in a 20 pt cell:
    the row height and the words' baselines (each line takes the height and baseline of its
    largest run).
  - J22 "HH HH HH" with the middle word a superscript run, J24 with it a subscript run: the row
    height and each word's ink (the run raised or lowered on its own, at two thirds).
  - J26 under ``[Red]@``, a run without properties and a blue run; J28 under ``[Blue]General``,
    a red run and a green one: the number format's section color replaces each run's color.
  - J30 under ``@" HH"`` and J32 under ``"HH "@``, a bold green 16 pt run and a run without
    properties: a text section that adds characters draws the result in the cell's font.
  - J34 to J46: "HH" alone in each font the cases' words may be drawn in (Calibri 11, 11 bold,
    16, 16 bold, 20, 11 superscript, 11 subscript), each word's ink to compare with.

- ``normal``: the Normal font is Arial 14 (the styles' first font, no font scheme). C4, in a
  cell font of red Calibri 11: a run without properties, a run of only bold, a run of only 11 pt,
  a run of only the name Calibri: a run's missing name and size are the Normal font's. C6 to C14:
  "HH" alone in Calibri 11, Arial 14 bold, Arial 11, Calibri 14 and Arial 14.
- ``text_and_runs``: C4's shared string holds both a ``<t>`` ("HH ") and a bold 11 pt Calibri
  run ("HH"), which the schema allows and Excel does not write, in a 16 pt red cell. Excel may
  refuse or repair the file (-50 with alerts off): it runs on its own, first, and a
  refusal is kept (``refused.txt``) and written as the group's one row, its ``name``
  ``refused:`` and Excel's message. C6 and C8 "HH" alone in Calibri 16 and Calibri 11 bold.

Measured for each case, in points from the cell's top left edge: the row's height and the text
(Excel's reads); the ink box; the ink bands of rows (``line_tops``, ``start+length`` each); the
words, found as the ink's columns split at the widest gaps, as many words as Excel's text holds
(``words`` from the cell's left edge, ``word_tops`` and ``word_bottoms`` from its top), each
word's color (the mean of its darkest tenth of ink pixels, ``RRGGBB``) and its ink area (pixels
of ink over the scale squared, square points); for the spilling texts, per column boundary from
C|D to H|I, whether its gridline is hidden (``1``) or shows (``0``) between 1 and 3 points
below the row's top, where no glyph reaches. ``read`` writes ``tests/captures/runs.tsv``, one row
per group, case and zoom.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "runs"
OUTPUT = ROOT / "tests" / "captures" / "runs.tsv"
GROUPS = ("main", "normal", "text_and_runs")
REFUSALS = ("text_and_runs",)
"""Groups whose file Excel may refuse: run on their own, a refusal is recorded as the answer."""
MARKERS = {"main": ("B2", "L48"), "normal": ("B2", "E16"), "text_and_runs": ("B2", "E10")}
ZOOM = 100
NORMAL = ("Arial", 14)
"""The ``normal`` group's Normal font."""
COLUMNS = (
    "group", "name", "zoom", "cell", "row_height", "text", "font_name", "font_size", "ink_left",
    "ink_right", "ink_top", "ink_bottom", "lines", "line_tops", "words", "word_tops",
    "word_bottoms", "word_colors", "word_areas", "hidden_gridlines",
)  # fmt: skip
READS = {
    "row_height": "geometry.height",
    "text": "geometry.text",
    "font_name": "display.font_name",
    "font_size": "display.font_size",
}
LONG = "a long text that spills with a LARGE run inside"
WRAPPED = "aaa BBB ccc ddd"
SPILL_COLUMNS = "CDEFGHI"
"""The spilling texts' columns: C holds the text, D to I are empty."""
RED, BLUE, GREEN = "FF0000", "0000FF", "00B050"


@dataclass(frozen=True)
class Run:
    """One ``<r>``: its text and the properties its ``<rPr>`` holds (``None``: no ``<rPr>``)."""

    text: str
    name: str | None = None
    size: float | None = None
    bold: bool = False
    color: str | None = None
    vert: str | None = None
    bare: bool = False
    """No ``<rPr>`` at all."""

    def xml(self) -> str:
        text = f'<t xml:space="preserve">{self.text}</t>'
        if self.bare:
            return f"<r>{text}</r>"
        found = []  # in the schema's order
        if self.name is not None:
            found.append(f'<rFont val="{self.name}"/>')
        if self.bold:
            found.append("<b/>")
        if self.color is not None:
            found.append(f'<color rgb="FF{self.color}"/>')
        if self.size is not None:
            found.append(f'<sz val="{self.size!r}"/>')
        if self.vert is not None:
            found.append(f'<vertAlign val="{self.vert}"/>')
        return f"<r><rPr>{''.join(found)}</rPr>{text}</r>"


def bare(text: str) -> Run:
    """A run without properties."""
    return Run(text, bare=True)


def calibri(text: str, size: float = 11, **options: Any) -> Run:
    """A run with a complete font, as Excel writes one."""
    return Run(text, "Calibri", size, **options)


@dataclass(frozen=True)
class Case:
    """One measured cell."""

    group: str
    name: str
    cell: str
    text: str | None = None
    """Plain text, for a cell in one font; else ``runs`` (or ``si``) holds it."""
    runs: tuple[Run, ...] = ()
    si: str | None = None
    """The shared string's inner XML when it is not runs alone."""
    font: dict[str, Any] = field(default_factory=dict)
    number_format: str | None = None
    wrap: bool = False
    measure: tuple[str, ...] = ("words", "ink")
    """``words``, ``lines``, ``ink``, ``gridlines``."""
    last: str | None = None
    """The last cell of the measured box, for text spilling right."""

    @property
    def shared(self) -> str | None:
        """The ``<si>``'s inner XML, or ``None`` for a cell holding plain text."""
        if self.si is not None:
            return self.si
        return "".join(run.xml() for run in self.runs) if self.runs else None

    @property
    def plain(self) -> str:
        """The text the cell holds, its runs' texts joined."""
        return self.text if self.text is not None else "".join(r.text for r in self.runs)

    @property
    def cells(self) -> list[str]:
        """The cells whose geometry measuring this case needs."""
        if "gridlines" in self.measure:
            row = self.cell[1:]
            return [f"{c}{row}" for c in SPILL_COLUMNS]
        return [self.cell, *([self.last] if self.last else [])]


def _reference(group: str, cell: str, label: str, **font: Any) -> Case:
    return Case(group, f"reference {label}", cell, text="HH", font=font, measure=("words", "ink"))


def _cases() -> list[Case]:
    head, tail = LONG.split("LARGE")
    spill = ("ink", "gridlines")
    wrapped = ("lines", "ink")
    main = [
        Case(
            "main",
            "widths add, bold 16 pt run",
            "C4",
            runs=(calibri(head), calibri("LARGE", 16, bold=True), calibri(tail)),
            measure=spill,
            last="I4",
        ),
        Case("main", "widths add, one font", "C6", text=LONG, measure=spill, last="I6"),
        *(
            Case(
                "main",
                f"wrapped, {size} pt run",
                cell,
                runs=(calibri("aaa "), calibri("BBB", size), calibri(" ccc ddd")),
                wrap=True,
                measure=wrapped,
            )
            for cell, size in (("C8", 14), ("C10", 20), ("C12", 28))
        ),
        Case(
            "main",
            "20 pt cell, 9 pt runs",
            "C14",
            runs=(calibri("aa ", 9), calibri("bb", 9, bold=True)),
            font={"size": 20},
            measure=("lines", "words", "ink"),
        ),
        Case(
            "main",
            "16 pt red cell, bare and bold-only runs",
            "J16",
            runs=(bare("HH "), Run("HH", bold=True)),
            font={"size": 16, "color": RED},
        ),
        Case(
            "main",
            "20 pt run in an 11 pt cell",
            "J18",
            runs=(calibri("aa "), calibri("BIG", 20), calibri(" aa")),
            measure=("lines", "words", "ink"),
        ),
        Case(
            "main",
            "20 pt cell",
            "J20",
            text="aa BIG aa",
            font={"size": 20},
            measure=("lines", "words", "ink"),
        ),
        *(
            Case(
                "main",
                f"{vert} run",
                cell,
                runs=(calibri("HH "), calibri("HH", vert=vert), calibri(" HH")),
            )
            for cell, vert in (("J22", "superscript"), ("J24", "subscript"))
        ),
        Case(
            "main",
            "[Red]@, bare and blue runs",
            "J26",
            runs=(bare("HH "), calibri("HH", color=BLUE)),
            number_format="[Red]@",
        ),
        Case(
            "main",
            "[Blue]General, red and green runs",
            "J28",
            runs=(calibri("HH ", color=RED), calibri("HH", color=GREEN)),
            number_format="[Blue]General",
        ),
        *(
            Case(
                "main",
                f"text section {fmt}, bold green 16 pt and bare runs",
                cell,
                runs=(calibri("HH", 16, bold=True, color=GREEN), bare(" HH")),
                number_format=fmt,
            )
            for cell, fmt in (("J30", '@" HH"'), ("J32", '"HH "@'))
        ),
        _reference("main", "J34", "Calibri 11"),
        _reference("main", "J36", "Calibri 11 bold", bold=True),
        _reference("main", "J38", "Calibri 16", size=16),
        _reference("main", "J40", "Calibri 16 bold", size=16, bold=True),
        _reference("main", "J42", "Calibri 20", size=20),
        _reference("main", "J44", "Calibri 11 superscript", vertAlign="superscript"),
        _reference("main", "J46", "Calibri 11 subscript", vertAlign="subscript"),
    ]
    normal = [
        Case(
            "normal",
            "red Calibri 11 cell: bare, bold-only, size-only, name-only runs",
            "C4",
            runs=(bare("HH "), Run("HH ", bold=True), Run("HH ", size=11), Run("HH", "Calibri")),
            font={"name": "Calibri", "size": 11, "color": RED},
        ),
        _reference("normal", "C6", "Calibri 11", name="Calibri", size=11),
        _reference("normal", "C8", "Arial 14 bold", name="Arial", size=14, bold=True),
        _reference("normal", "C10", "Arial 11", name="Arial", size=11),
        _reference("normal", "C12", "Calibri 14", name="Calibri", size=14),
        _reference("normal", "C14", "Arial 14", name="Arial", size=14),
    ]
    both = [
        Case(
            "text_and_runs",
            "<t> and a bold run in a 16 pt red cell",
            "C4",
            text="HH HH",
            si='<t xml:space="preserve">HH </t>' + calibri("HH", bold=True).xml(),
            font={"size": 16, "color": RED},
        ),
        _reference("text_and_runs", "C6", "Calibri 16", size=16),
        _reference("text_and_runs", "C8", "Calibri 11 bold", bold=True),
    ]
    return [*main, *normal, *both]


CASES = _cases()

EXPECTED: dict[str, dict[str, dict[str, Any]]] = {
    "main": {
        "C4": {
            "row_height": 21.0,
            "text": "a long text that spills with a LARGE run inside",
            "baselines": (17.0,),
            "reach": 166.0,
            "spill": 4,
        },
        "C6": {
            "row_height": 15.0,
            "text": "a long text that spills with a LARGE run inside",
            "baselines": (12.0,),
            "reach": 151.0,
            "spill": 3,
        },
        "C8": {"row_height": 36.0, "text": "aaa BBB ccc ddd", "baselines": (17.0, 33.0)},
        "C10": {"row_height": 59.0, "text": "aaa BBB ccc ddd", "baselines": (15.0, 39.0, 56.0)},
        "C12": {"row_height": 70.0, "text": "aaa BBB ccc ddd", "baselines": (15.0, 47.0, 67.0)},
        "C14": {
            "row_height": 26.0,
            "text": "aa bb",
            "baselines": (21.0,),
            "words": (("Calibri 9", "000000", 0.0), ("Calibri 9 bold", "000000", 0.0)),
        },
        "J16": {
            "row_height": 21.0,
            "text": "HH HH",
            "baselines": (17.0,),
            "words": (("Calibri 16", "FF0000", 0.0), ("Calibri 11 bold", "000000", 0.0)),
        },
        "J18": {
            "row_height": 26.0,
            "text": "aa BIG aa",
            "baselines": (21.0,),
            "words": (
                ("Calibri 11", "000000", 0.0),
                ("Calibri 20", "000000", 0.0),
                ("Calibri 11", "000000", 0.0),
            ),
        },
        "J20": {
            "row_height": 26.0,
            "text": "aa BIG aa",
            "baselines": (21.0,),
            "words": (
                ("Calibri 20", "000000", 0.0),
                ("Calibri 20", "000000", 0.0),
                ("Calibri 20", "000000", 0.0),
            ),
        },
        "J22": {
            "row_height": 17.0,
            "text": "HH HH HH",
            "baselines": (14.0,),
            "words": (
                ("Calibri 11", "000000", 0.0),
                ("Calibri 11 superscript", "000000", -5.0),
                ("Calibri 11", "000000", 0.0),
            ),
        },
        "J24": {
            "row_height": 17.0,
            "text": "HH HH HH",
            "baselines": (14.0,),
            "words": (
                ("Calibri 11", "000000", 0.0),
                ("Calibri 11 subscript", "000000", 2.0),
                ("Calibri 11", "000000", 0.0),
            ),
        },
        "J26": {
            "row_height": 15.0,
            "text": "HH HH",
            "baselines": (12.0,),
            "words": (("Calibri 11", "FF0000", 0.0), ("Calibri 11", "FF0000", 0.0)),
        },
        "J28": {
            "row_height": 15.0,
            "text": "HH HH",
            "baselines": (12.0,),
            "words": (("Calibri 11", "0000FF", 0.0), ("Calibri 11", "0000FF", 0.0)),
        },
        "J30": {
            "row_height": 21.0,
            "text": "HH HH HH",
            "baselines": (17.0,),
            "words": (
                ("Calibri 11", "000000", 0.0),
                ("Calibri 11", "000000", 0.0),
                ("Calibri 11", "000000", 0.0),
            ),
        },
        "J32": {
            "row_height": 21.0,
            "text": "HH HH HH",
            "baselines": (17.0,),
            "words": (
                ("Calibri 11", "000000", 0.0),
                ("Calibri 11", "000000", 0.0),
                ("Calibri 11", "000000", 0.0),
            ),
        },
        "J34": {
            "row_height": 15.0,
            "text": "HH",
            "baselines": (12.0,),
            "words": (("Calibri 11", "000000", 0.0),),
        },
        "J36": {
            "row_height": 15.0,
            "text": "HH",
            "baselines": (12.0,),
            "words": (("Calibri 11 bold", "000000", 0.0),),
        },
        "J38": {
            "row_height": 21.0,
            "text": "HH",
            "baselines": (17.0,),
            "words": (("Calibri 16", "000000", 0.0),),
        },
        "J40": {
            "row_height": 21.0,
            "text": "HH",
            "baselines": (17.0,),
            "words": (("Calibri 16 bold", "000000", 0.0),),
        },
        "J42": {
            "row_height": 26.0,
            "text": "HH",
            "baselines": (21.0,),
            "words": (("Calibri 20", "000000", 0.0),),
        },
        "J44": {
            "row_height": 17.0,
            "text": "HH",
            "baselines": (9.0,),
            "words": (("Calibri 11 superscript", "000000", 0.0),),
        },
        "J46": {
            "row_height": 17.0,
            "text": "HH",
            "baselines": (16.0,),
            "words": (("Calibri 11 subscript", "000000", 0.0),),
        },
    },
    "normal": {
        "C4": {
            "row_height": 19.0,
            "text": "HH HH HH HH",
            "baselines": (15.0,),
            "words": (
                ("Calibri 11", "FF0000", 0.0),
                ("Arial 14 bold", "000000", 0.0),
                ("Arial 11", "000000", 0.0),
                ("Calibri 14", "000000", 0.0),
            ),
        },
        "C6": {
            "row_height": 18.0,
            "text": "HH",
            "baselines": (15.0,),
            "words": (("Calibri 11", "000000", 0.0),),
        },
        "C8": {
            "row_height": 18.0,
            "text": "HH",
            "baselines": (15.0,),
            "words": (("Arial 14 bold", "000000", 0.0),),
        },
        "C10": {
            "row_height": 18.0,
            "text": "HH",
            "baselines": (15.0,),
            "words": (("Arial 11", "000000", 0.0),),
        },
        "C12": {
            "row_height": 19.0,
            "text": "HH",
            "baselines": (15.0,),
            "words": (("Calibri 14", "000000", 0.0),),
        },
        "C14": {
            "row_height": 18.0,
            "text": "HH",
            "baselines": (15.0,),
            "words": (("Arial 14", "000000", 0.0),),
        },
    },
    "text_and_runs": {
        "C4": {
            "row_height": 21.0,
            "text": "HH HH",
            "baselines": (17.0,),
            "words": (("Calibri 16", "FF0000", 0.0), ("Calibri 11 bold", "000000", 0.0)),
        },
        "C6": {
            "row_height": 21.0,
            "text": "HH",
            "baselines": (17.0,),
            "words": (("Calibri 16", "000000", 0.0),),
        },
        "C8": {
            "row_height": 15.0,
            "text": "HH",
            "baselines": (12.0,),
            "words": (("Calibri 11 bold", "000000", 0.0),),
        },
    },
}

"""What sheetpeek draws for each case, by group and cell: the row's height, the text drawn, each
line's baseline below the row's top, each word's font (as the references name it), color and
baseline shift, and for the spilling texts the reach of the last advance past C's right edge and
the columns spilled into. All in points at zoom 100."""


def cases(group: str | None = None) -> list[Case]:
    """Every case of ``group``, or of every group, in order."""
    return [c for c in CASES if group is None or c.group == group]


def fill(ws: Any, group: str) -> list[Case]:
    """Lay one group out on a worksheet, its rich strings as placeholders; its cases.

    :func:`postprocess` then moves the cells' text into shared strings with their runs.
    """
    if group == "main":
        ws.column_dimensions["J"].width = 30
    else:
        ws.column_dimensions["C"].width = 30
    found = cases(group)
    for case in found:
        cell = ws[case.cell]
        cell.value = case.plain if case.shared is None else "placeholder"
        if case.font:
            options = dict(case.font)
            if "color" in options:
                options["color"] = "FF" + options["color"]
            options.setdefault("name", "Calibri")
            options.setdefault("size", 11)
            cell.font = Font(**options)
        if case.number_format is not None:
            cell.number_format = case.number_format
        if case.wrap:
            cell.alignment = Alignment(wrap_text=True)
    return found


def shared_strings(text: str, strings: dict[str, str]) -> str:
    """A worksheet's XML with the cells of ``strings`` (inline strings) turned into indices."""
    for index, coordinate in enumerate(strings):
        pattern = rf'<c r="{coordinate}"([^>]*?) t="inlineStr"([^>]*)><is>.*?</is></c>'
        text, count = re.subn(pattern, rf'<c r="{coordinate}"\1 t="s"\2><v>{index}</v></c>', text)
        if count != 1:
            raise ValueError(f"{coordinate} holds no inline string")
    return text


def normal_font(text: str) -> str:
    """The styles with the Normal font (the first ``<font>``) made :data:`NORMAL`, no scheme."""
    start = text.index("<font>")
    end = text.index("</font>", start) + len("</font>")
    name, size = NORMAL
    font = f'<font><sz val="{size}"/><name val="{name}"/><family val="2"/></font>'
    return text[:start] + font + text[end:]


def postprocess(group: str) -> Any:
    """Move a group's rich strings into ``xl/sharedStrings.xml``, and set the Normal font."""

    def edit(path: Path) -> None:
        from capture.screen.run import patch_part

        strings = {c.cell: c.shared for c in cases(group) if c.shared is not None}
        items = "".join(f"<si>{xml}</si>" for xml in strings.values())
        ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        sst = f'<sst {ns} count="{len(strings)}" uniqueCount="{len(strings)}">{items}</sst>'
        kind = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
        rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        edits: dict[str, Any] = {
            "xl/worksheets/sheet1.xml": lambda t: shared_strings(t, strings),  # type: ignore[arg-type]
            "xl/sharedStrings.xml": sst,
            "[Content_Types].xml": lambda t: t.replace(
                "</Types>",
                f'<Override PartName="/xl/sharedStrings.xml" ContentType="{kind}"/></Types>',
            ),
            "xl/_rels/workbook.xml.rels": lambda t: t.replace(
                "</Relationships>",
                f'<Relationship Id="rIdSst" Type="{rel}" Target="sharedStrings.xml"/>'
                "</Relationships>",
            ),
        }
        if group == "normal":
            edits["xl/styles.xml"] = normal_font
        patch_part(path, edits)

    return edit


def _cells(group: str) -> list[str]:
    return list(dict.fromkeys(c for case in cases(group) for c in case.cells))


def build(path: Path) -> list[Case]:
    """Write every group to a workbook of its own in the folder ``path``; every case."""
    path.mkdir(parents=True, exist_ok=True)
    found = []
    for group in GROUPS:
        wb = openpyxl.Workbook()
        found.extend(fill(wb.active, group))
        target = path / f"runs_{group}.xlsx"
        wb.save(target)
        postprocess(group)(target)
    return found


def run(groups: tuple[str, ...] | None = None) -> None:
    """Capture each group's window, then write the TSV (needs Excel).

    Without groups, every group but :data:`REFUSALS`, which run on their own first.
    """
    from capture.screen.run import capture
    from capture.screen.session import ExcelError

    if groups is None:
        groups = tuple(g for g in GROUPS if g not in REFUSALS)
    for group in groups:
        try:
            capture(
                lambda ws, group=group: fill(ws, group),
                cells=_cells(group),
                out=BUILD_DIR / group,
                markers=MARKERS[group],
                zooms=(ZOOM,),
                name=f"runs_{group}",
                postprocess=postprocess(group),
            )
        except ExcelError as error:
            if group not in REFUSALS:
                raise
            record_refusal(group, str(error))
            print(f"{group}: Excel refused the file", file=sys.stderr)
    write()


def record_refusal(group: str, message: str) -> None:
    """Keep Excel's refusal of a group's file in ``refused.txt``, without the script's path."""
    out = BUILD_DIR / group
    out.mkdir(parents=True, exist_ok=True)
    reason = message.rpartition("execution error: ")[2].strip() or message
    (out / "refused.txt").write_text(reason + "\n", encoding="utf-8")


# -- measuring -----------------------------------------------------------------------------------


def _span(cap: Any, start: float, length: float, origin: float) -> str:
    from capture.screen.ink import pt

    return f"{pt(cap, start - origin)}+{pt(cap, length)}"


def split_words(cap: Any, box: tuple[int, int, int, int], count: int) -> list[Any]:
    """The ink of ``count`` words: the ink's columns split at their ``count - 1`` widest gaps."""
    from capture.screen.ink import bands, ink

    columns = bands(cap, box, axis="x")
    if not columns:
        return []
    # the widest first (a gap is the next column's start less this column's end), ties in order
    gaps = sorted(
        range(len(columns) - 1),
        key=lambda k: -(columns[k + 1][0] - (columns[k][0] + columns[k][1])),
    )
    cuts = sorted(gaps[: max(0, count - 1)])
    found, start = [], 0
    for k in [*cuts, len(columns) - 1]:
        x0, x1 = columns[start][0], columns[k][0] + columns[k][1]
        found.append(ink(cap, (x0, box[1], x1, box[3])))
        start = k + 1
    return found


def word_color(cap: Any, box: tuple[int, int, int, int]) -> str:
    """The mean of the darkest tenth of a word's ink pixels, ``RRGGBB``."""
    from capture.screen.ink import is_dark
    from capture.screen.run import luma

    x0, y0, x1, y1 = box
    pixels = sorted(
        (p for y in range(y0, y1) for x in range(x0, x1) if is_dark(p := cap.image.pixel(x, y))),
        key=luma,
    )
    core = pixels[: max(1, len(pixels) // 10)]
    return "".join(f"{round(sum(p[i] for p in core) / len(core)):02X}" for i in range(3))


def word_area(cap: Any, box: tuple[int, int, int, int]) -> str:
    """A word's ink pixels over the scale squared: its ink in square points."""
    from capture.screen.ink import is_dark

    x0, y0, x1, y1 = box
    count = sum(is_dark(cap.image.pixel(x, y)) for y in range(y0, y1) for x in range(x0, x1))
    return f"{round(count / cap.scale**2, 2):g}"


def hidden_gridlines(cap: Any, case: Case) -> str:
    """Per boundary from the case's first column to its last, ``1`` where no gridline shows."""
    from capture.screen.ink import nearest

    gray = cap.gridline
    top = cap.edge(case.cell, "top")
    ys = range(round(top + 1 * cap.scale), round(top + 3 * cap.scale))
    flags = []
    for cell in case.cells[:-1]:
        x = round(cap.edge(cell, "right"))
        shows = gray is not None and any(
            nearest(cap.image.pixel(px, y), {"grid": gray}, 20) == "grid"
            for y in ys
            for px in range(x - 1, x + 2)
        )
        flags.append("0" if shows else "1")
    return "".join(flags)


def measure(cap: Any, case: Case) -> dict[str, str] | None:
    """One case's measurements at one zoom, or ``None`` when it is outside the window."""
    from capture.screen.ink import bands, cells_box, ink, pt, visible

    if not all(visible(cap, cell) for cell in case.cells):
        return None
    left, top = cap.edge(case.cell, "left"), cap.edge(case.cell, "top")
    right = cap.edge(case.last or case.cell, "right")
    box = cells_box(cap, case.cell, case.last, inset=1)
    found: dict[str, str] = {}
    glyph = ink(cap, box)
    if glyph is not None and "ink" in case.measure:
        found["ink_left"] = pt(cap, glyph[0] - left)
        found["ink_right"] = pt(cap, glyph[2] - right)
        found["ink_top"] = pt(cap, glyph[1] - top)
        found["ink_bottom"] = pt(cap, glyph[3] - top)
    if "lines" in case.measure:
        lines = bands(cap, box)
        found["lines"] = str(len(lines))
        found["line_tops"] = " ".join(_span(cap, y, n, top) for y, n in lines)
    if glyph is not None and "words" in case.measure:
        text = cap.value(case.cell, "geometry.text") or case.plain
        words = [w for w in split_words(cap, box, len(text.split())) if w is not None]
        found["words"] = " ".join(_span(cap, w[0], w[2] - w[0], left) for w in words)
        found["word_tops"] = " ".join(pt(cap, w[1] - top) for w in words)
        found["word_bottoms"] = " ".join(pt(cap, w[3] - top) for w in words)
        found["word_colors"] = " ".join(word_color(cap, w) for w in words)
        found["word_areas"] = " ".join(word_area(cap, w) for w in words)
    if "gridlines" in case.measure:
        found["hidden_gridlines"] = hidden_gridlines(cap, case)
    return found


def write() -> None:
    """Measure the saved captures again (no Excel) and write ``runs.tsv``."""
    from capture.screen.run import header, load

    records = []
    first: dict[str, str] = {}
    for group in GROUPS:
        out = BUILD_DIR / group
        refused = out / "refused.txt"
        if refused.exists():
            reason = refused.read_text(encoding="utf-8").strip()
            records.append({"group": group, "name": f"refused: {reason}"})
            continue
        if not (out / "formats.tsv").exists():
            print(f"{group}: not captured yet, left out", file=sys.stderr)
            continue
        first = first or header(out)
        for cap in load(out):
            for case in cases(group):
                found = measure(cap, case)
                if found is None:
                    continue
                reads = {key: cap.value(case.cell, read) or "" for key, read in READS.items()}
                records.append(
                    {
                        "group": group,
                        "name": case.name,
                        "zoom": str(cap.zoom),
                        "cell": case.cell,
                        **reads,
                        **found,
                    }
                )
    with open(OUTPUT, "w", encoding="utf-8") as out:
        for key, value in first.items():
            if key not in ("markers", "suite", "workbooks_open_after"):
                out.write(f"# {key}: {value}\n")
        out.write("# suite: runs\n" + "\t".join(COLUMNS) + "\n")
        for record in records:
            out.write("\t".join(record.get(c, "") for c in COLUMNS) + "\n")
    print(f"{len(records)} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        placed = build(Path(argv[1]))
        print(f"{len(placed)} cases in {argv[1]}")
        return 0
    if argv[:1] == ["run"] and argv[1:] and all(group in GROUPS for group in argv[1:]):
        run(tuple(argv[1:]))
        return 0
    if argv == ["run"]:
        run()
        return 0
    if argv == ["read"]:
        write()
        return 0
    print(__doc__.split("\n\n")[1], file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
