"""The reader's detectors: what a workbook uses that sheetpeek does not read yet.

A detector reports what a file **uses**, never what it merely contains: an element counts only
when it states something Excel draws differently from its default (a print area, not default
margins; a filter with a range, not an empty ``autoFilter``). Elements that draw nothing on the
sheet (margins, page setup, protection, validation, the selection) are never reported. Anything
the sheet holds that no rule names is reported as unrecognized, so a feature nobody listed is
not dropped silently either. A feature counts as used when the file sets it to something other
than its default; ``ARCHITECTURE.md`` lists the rules.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from collections.abc import Callable, Iterable, Iterator
from dataclasses import dataclass, field
from typing import Any

import xlio

from sheetpeek._model import CellRange
from sheetpeek._notices import Collector, Outcome, Reason
from sheetpeek._structure import PHONETIC, RICH, Structure
from sheetpeek._style import Code

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main"
_CHUNK = 1 << 20

# What each reader code reports: outcome, reason, and what is drawn instead (or why not).
CODES: dict[str, tuple[Outcome, Reason, str]] = {
    "format.palette": ("approximated", "not_implemented", ""),
    "text.phonetic": ("omitted", "not_implemented", ""),
    "style.row_column": (
        "omitted",
        "not_implemented",
        "fills and borders of whole rows and columns, on cells the file has no element for",
    ),
    "view.right_to_left": ("approximated", "not_implemented", "drawn left to right"),
    "view.formulas": (
        "approximated",
        "not_implemented",
        "values are shown; Excel shows the formulas",
    ),
    "view.page_layout": ("approximated", "not_implemented", "drawn as the Normal view"),
    "view.panes": ("omitted", "not_implemented", ""),
    "cf": ("omitted", "not_implemented", ""),
    "filter.buttons": ("omitted", "not_implemented", ""),
    "table": ("omitted", "not_implemented", "table styles and buttons not drawn"),
    "sparkline": ("omitted", "not_implemented", ""),
    "slicer": ("omitted", "not_implemented", ""),
    "timeline": ("omitted", "not_implemented", ""),
    "drawing.picture": ("omitted", "not_implemented", ""),
    "drawing.picture_linked": (
        "omitted",
        "unsafe",
        "Excel fetches the linked file; sheetpeek never fetches",
    ),
    "drawing.chart": ("omitted", "not_implemented", ""),
    "drawing.shape": ("omitted", "not_implemented", ""),
    "drawing.other": ("omitted", "not_implemented", ""),
    "drawing.ole": ("omitted", "not_implemented", ""),
    "control": ("omitted", "not_implemented", ""),
    "picture.background": ("omitted", "not_implemented", ""),
    "comment": ("omitted", "not_implemented", "the note indicators and shown notes"),
    "comment.threaded": ("omitted", "not_implemented", "the comment indicators"),
    "sheet.tab_color": (
        "omitted",
        "option",
        "drawn only on the HTML workbook page (sheetpeek.render_workbook)",
    ),
    "sheet.hidden": ("omitted", "option", "drawn only with hidden_sheets=True"),
    "sheet.chartsheet": ("omitted", "not_implemented", ""),
    "sheet.dialogsheet": (
        "omitted",
        "not_implemented",
        "a legacy form: its frame and controls are not drawn",
    ),
    "sheet.macrosheet": (
        "approximated",
        "not_implemented",
        "shown as a worksheet; its formulas are Excel 4.0 macros",
    ),
    "sheet.unknown_element": ("omitted", "not_implemented", ""),
    "sheet.unknown_extension": ("omitted", "not_implemented", ""),
}

# Children of a sheet that draw nothing on it: never reported.
_SILENT = frozenset(
    {
        "dimension",
        "sheetCalcPr",
        "sheetProtection",
        "protectedRanges",
        "scenarios",
        "sortState",
        "dataConsolidate",
        "customSheetViews",
        "phoneticPr",
        "dataValidations",
        "printOptions",
        "pageMargins",
        "pageSetup",
        "headerFooter",
        "customProperties",
        "cellWatches",
        "ignoredErrors",
        "smartTags",
        "webPublishItems",
        "legacyDrawingHF",
        "drawingHF",
    }
)
# sheetPr and sheetViews are also checked below.
_READ = frozenset(
    {
        "sheetPr", "sheetViews", "sheetFormatPr", "cols", "mergeCells", "hyperlinks", "sheetData",
        "rowBreaks", "colBreaks",
    }
)  # fmt: skip
"""Children the model reads. Page breaks are read and, like the print area, not reported: Excel's
Normal view draws neither, so a sheet with them is shown as Excel shows it (§4.18)."""

# Sheet extension URIs. Excel compares them exactly: conditional formatting's is captured
# (lower-case "407d"; the upper-case form printed in [MS-XLSX] is ignored by Excel). The others
# are matched ignoring case until captured (unverified), so that none is missed.
_CF_URI = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"
_URIS_ANY_CASE: dict[str, str | None] = {
    "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}": "sparkline",
    "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}": "slicer",
    "{3A4CF648-6AED-40F4-86FF-DC5316D8AED3}": "slicer",
    "{7E03D99C-DC04-49D9-9315-930204A7B6E9}": "timeline",
    # These draw nothing on the sheet (validation, protected ranges, ignored errors) or are
    # drawn through the drawing part (web extensions).
    "{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}": None,
    "{FC87AEE6-9EDD-4A0A-B7FB-166176984837}": None,
    "{01252117-D84E-4E92-8308-4BE1C098FCBB}": None,
    "{F7C9EE02-42E1-4005-9D12-6889AFFD525C}": None,
}

_SHEET_KINDS = (
    ("/worksheet", "worksheet"),
    ("/chartsheet", "chartsheet"),
    ("/dialogsheet", "dialogsheet"),
    ("macrosheet", "macrosheet"),  # Microsoft's xlMacrosheet and xlIntlMacrosheet
)
_OBJECT_TYPE = re.compile(rb'ObjectType\s*=\s*"([A-Za-z]+)"')


@dataclass(frozen=True, slots=True)
class WorkbookInfo:
    """What the workbook part says about its sheets."""

    kinds: dict[str, str] = field(default_factory=dict)
    """Sheet name → ``worksheet``, ``chartsheet``, ``dialogsheet`` or ``macrosheet``."""
    show_objects: str = "all"
    strings_part: str | None = None


def _related(parts: xlio.Parts, part: str) -> list[tuple[str, str, str, bool]]:
    """A part's relationships as (id, type, target, external); targets resolved to part names."""
    return [
        (r.id, r.type, r.target if r.external else (r.target_part or ""), r.external)
        for r in parts.relationships(part)
    ]


def _workbook_part(parts: xlio.Parts) -> str:
    """The workbook part's name, from the package's own relationship to it."""
    return next(
        (
            r.target_part
            for r in parts.relationships("")
            if r.type.endswith("/officeDocument") and r.target_part
        ),
        "xl/workbook.xml",
    )


def workbook_info(parts: xlio.Parts) -> WorkbookInfo:
    """Read sheet kinds, ``showObjects`` and the shared strings part."""
    part = _workbook_part(parts)
    rels = {rid: (kind, target) for rid, kind, target, _ in _related(parts, part)}
    root = ET.fromstring(parts.read(part))
    kinds: dict[str, str] = {}
    for sheet in root.iter(f"{{{_MAIN}}}sheet"):
        name = sheet.get("name", "")
        kind_url = rels.get(sheet.get(f"{{{_REL}}}id", ""), ("", ""))[0]
        kinds[name] = next(
            (k for suffix, k in _SHEET_KINDS if suffix in kind_url.lower()), "worksheet"
        )
    properties = root.find(f"{{{_MAIN}}}workbookPr")
    show_objects = properties.get("showObjects", "all") if properties is not None else "all"
    strings = next(
        (target for kind, target in rels.values() if kind.endswith("/sharedStrings")), None
    )
    return WorkbookInfo(kinds, show_objects, strings)


def book_notices(collector: Collector, info: WorkbookInfo) -> None:
    """Workbook-level notices: sheets that cannot be shown as worksheets."""
    for name, kind in info.kinds.items():
        if kind == "chartsheet":
            _report(collector, "sheet.chartsheet", context={"sheet": name}, detail=f"{name!r}")
        elif kind == "dialogsheet":
            _, _, how = CODES["sheet.dialogsheet"]
            detail = f"{name!r}, {how}"
            _report(collector, "sheet.dialogsheet", context={"sheet": name}, detail=detail)


def hidden_sheets(collector: Collector, names: list[str], very_hidden: int) -> None:
    """The workbook-level notice for hidden sheets a workbook page leaves out."""
    _, _, how = CODES["sheet.hidden"]
    detail = f"{very_hidden} of them very hidden; {how}" if very_hidden else how
    context: dict[str, str | int | float] = {"very_hidden": very_hidden}
    _report(collector, "sheet.hidden", count=len(names), detail=detail, context=context)


def string_marks(parts: xlio.Parts, part: str | None) -> bytes | None:
    """Per shared string, :data:`RICH` and :data:`PHONETIC` bits; ``None`` when none is set.

    A byte search skips the parse for the many workbooks whose strings have neither.
    """
    if part is None or part not in parts:
        return None
    found = False
    tail = b""
    with parts.open(part) as handle:
        while chunk := handle.read(_CHUNK):
            data = tail + chunk
            if b"<r>" in data or b":r>" in data or b"rPh" in data:
                found = True
                break
            tail = data[-8:]
    if not found:
        return None
    marks = bytearray()
    with parts.open(part) as handle:
        for _, element in ET.iterparse(handle):
            if element.tag.rpartition("}")[2] != "si":
                continue
            mark = 0
            for child in element:
                local = child.tag.rpartition("}")[2]
                if local == "r":
                    mark |= RICH
                elif local == "rPh":
                    mark |= PHONETIC
            marks.append(mark)
            element.clear()
    return bytes(marks) if any(marks) else None


def visible_style(styles: Any, index: int) -> bool:
    """Whether a row or column style draws on empty cells: a fill or a border side.

    ``styles`` is the workbook's ``xlio.Styles``.
    """
    formats = styles.cell_formats
    if not 0 <= index < len(formats):
        return False
    fill = formats[index].fill
    if fill is not None and getattr(fill, "pattern", "gradient") not in (None, "none"):
        return True
    border = formats[index].border
    sides = ("left", "right", "top", "bottom")
    return border is not None and any(
        getattr(getattr(border, s, None), "style", None) not in (None, "none") for s in sides
    )


@dataclass(slots=True)
class SheetContext:
    """What a sheet's rules need besides the element: the package and what was found so far."""

    collector: Collector
    parts: xlio.Parts
    rels: dict[str, tuple[str, str, bool]]
    objects: bool
    """Whether the workbook shows objects (``showObjects`` is not ``none``)."""


Rule = Callable[[SheetContext, ET.Element], None]
"""Reports what one child of the sheet uses."""

RULES: dict[str, Rule] = {}
"""Rules by a sheet child's local name (in the main namespace). A child with no rule, and not
silent or read, is reported as unrecognized. A feature that comes to be drawn replaces its rule."""


def rule(*names: str) -> Callable[[Rule], Rule]:
    """Register a rule for sheet children with these local names."""

    def register(function: Rule) -> Rule:
        for name in names:
            RULES[name] = function
        return function

    return register


def sheet_notices(
    collector: Collector,
    parts: xlio.Parts,
    part: str,
    structure: Structure,
    info: WorkbookInfo,
    name: str,
    ref: CellRange | None,
) -> None:
    """Report what one sheet uses at sheet level and from its other parts."""
    if info.kinds.get(name) == "macrosheet":
        _report(collector, "sheet.macrosheet")
    context = SheetContext(
        collector,
        parts,
        {rid: (kind, target, external) for rid, kind, target, external in _related(parts, part)},
        info.show_objects != "none",
    )
    for child in _children(structure.children):
        local, namespace = _local(child)
        if namespace != _MAIN:
            _report(
                collector,
                "sheet.unknown_element",
                detail=f"<{local}>",
                context={"element": child.tag},
            )
        elif local in RULES:
            RULES[local](context, child)
        elif local not in _SILENT and local not in _READ:
            _report(
                collector, "sheet.unknown_element", detail=f"<{local}>", context={"element": local}
            )
    _comments(collector, parts, context.rels, ref)


@rule("sheetPr")
def _sheet_properties(context: SheetContext, child: ET.Element) -> None:
    tab = child.find(f"{{{_MAIN}}}tabColor")
    if tab is not None and not _true(tab.get("auto")) and tab.attrib:
        _report(context.collector, "sheet.tab_color")


@rule("sheetViews")
def _sheet_views(context: SheetContext, child: ET.Element) -> None:
    view = child.find(f"{{{_MAIN}}}sheetView")
    if view is not None:
        _view(context.collector, view)


@rule("conditionalFormatting")
def _conditional(context: SheetContext, child: ET.Element) -> None:
    if rules := sum(1 for _ in child.iter(f"{{{_MAIN}}}cfRule")):
        _report(context.collector, "cf", count=rules)


@rule("autoFilter")
def _filter(context: SheetContext, child: ET.Element) -> None:
    if child.get("ref"):
        _report(context.collector, "filter.buttons")


@rule("tableParts")
def _tables(context: SheetContext, child: ET.Element) -> None:
    if tables := sum(1 for _ in child.iter(f"{{{_MAIN}}}tablePart")):
        _report(context.collector, "table", count=tables)


@rule("drawing", "legacyDrawing")
def _drawings(context: SheetContext, child: ET.Element) -> None:
    if not context.objects:
        return
    target = context.rels.get(child.get(f"{{{_REL}}}id", ""))
    if target is None or target[2]:
        return
    if _local(child)[0] == "drawing":
        _drawing(context.collector, context.parts, target[1])
    else:
        _vml(context.collector, context.parts, target[1])


@rule("picture")
def _background(context: SheetContext, child: ET.Element) -> None:
    if child.get(f"{{{_REL}}}id"):
        _report(context.collector, "picture.background")


@rule("oleObjects", "controls")
def _embedded(context: SheetContext, child: ET.Element) -> None:
    local = _local(child)[0]
    item, code = ("oleObject", "drawing.ole") if local == "oleObjects" else ("control", "control")
    count = sum(1 for e in child.iter() if _local(e)[0] == item)
    if count and context.objects:
        _report(context.collector, code, count=count)


@rule("extLst")
def _extension_list(context: SheetContext, child: ET.Element) -> None:
    _extensions(context.collector, child)


def _report(
    collector: Collector,
    code: str,
    *,
    count: int = 1,
    detail: str = "",
    context: dict[str, str | int | float] | None = None,
    row: int | None = None,
    column: int | None = None,
    cells: CellRange | None = None,
) -> None:
    outcome, reason, how = CODES[code]
    collector.add(
        code,
        outcome,
        reason,
        count=count,
        detail=detail or how,
        context=context,
        row=row,
        column=column,
        cells=cells,
    )


def report_cells(collector: Collector, code: str, cells: Iterable[CellRange]) -> None:
    """Report a cell-level code over ranges."""
    for cell_range in cells:
        _report(collector, code, cells=cell_range)


def report_style(collector: Collector, code: Code, row: int, column: int) -> None:
    """Report one style finding at a cell."""
    name, detail, context = code
    _report(collector, name, detail=detail, context=dict(context), row=row, column=column)


def _view(collector: Collector, view: ET.Element) -> None:
    if _true(view.get("rightToLeft")):
        _report(collector, "view.right_to_left")
    if _true(view.get("showFormulas")):
        _report(collector, "view.formulas")
    if view.get("view") in ("pageLayout", "pageBreakPreview"):
        _report(collector, "view.page_layout", context={"view": view.get("view", "")})
    pane = view.find(f"{{{_MAIN}}}pane")
    if pane is not None:
        state = pane.get("state", "split")
        # A frozen pane is drawn (HTML) or, like printing, has nothing to show (§4.18).
        if state not in ("frozen", "frozenSplit") and (
            _number(pane.get("xSplit")) or _number(pane.get("ySplit"))
        ):
            _report(collector, "view.panes", context={"state": state})


def _extensions(collector: Collector, ext_list: ET.Element) -> None:
    for ext in ext_list:
        if _local(ext)[0] != "ext":
            continue
        uri = ext.get("uri", "")
        if uri == _CF_URI:
            rules = sum(1 for e in ext.iter() if _local(e)[0] == "cfRule")
            if rules:
                _report(collector, "cf", count=rules)
        elif uri.upper() == _CF_URI.upper():
            continue  # Excel ignores the extension unless its URI matches exactly (captured)
        elif uri.upper() in _URIS_ANY_CASE:
            code = _URIS_ANY_CASE[uri.upper()]
            if code is not None:
                _report(collector, code)
        else:
            _report(collector, "sheet.unknown_extension", detail=uri, context={"uri": uri})


def _drawing(collector: Collector, parts: xlio.Parts, part: str) -> None:
    try:
        root = ET.fromstring(parts.read(part))
    except (KeyError, ET.ParseError):
        return
    counts: dict[str, int] = {}
    for anchor in root:
        for item in _children(list(anchor)):
            _drawn(item, counts)
    for code, count in counts.items():
        _report(collector, code, count=count)


def _drawn(item: ET.Element, counts: dict[str, int]) -> None:
    local = _local(item)[0]
    code = None
    if local == "pic":
        blip = item.find(f".//{{{_DRAWING}}}blip")
        linked = (
            blip is not None and blip.get(f"{{{_REL}}}link") and not blip.get(f"{{{_REL}}}embed")
        )
        code = "drawing.picture_linked" if linked else "drawing.picture"
    elif local == "graphicFrame":
        data = item.find(f".//{{{_DRAWING}}}graphicData")
        uri = data.get("uri", "") if data is not None else ""
        if uri.endswith("/chart") or "chartex" in uri:
            code = "drawing.chart"
        elif uri.endswith("/slicer") or uri.endswith("/timeslicer"):
            code = None  # reported through the sheet's extension list
        else:
            code = "drawing.other"
    elif local in ("sp", "cxnSp"):
        code = "drawing.shape"
    elif local == "grpSp":
        for child in _children(list(item)):
            _drawn(child, counts)
    elif local in ("contentPart",):
        code = "drawing.other"
    if code:
        counts[code] = counts.get(code, 0) + 1


def _vml(collector: Collector, parts: xlio.Parts, part: str) -> None:
    """Form controls in a legacy (VML) drawing; notes are reported from the comments part.

    VML is often not well-formed XML, so it is searched, not parsed.
    """
    try:
        data = parts.read(part)
    except KeyError:
        return
    controls = sum(1 for kind in _OBJECT_TYPE.findall(data) if kind != b"Note")
    if controls:
        _report(collector, "control", count=controls)


def _comments(
    collector: Collector,
    parts: xlio.Parts,
    rels: dict[str, tuple[str, str, bool]],
    ref: CellRange | None,
) -> None:
    threaded: set[str] = set()
    for rel_type, target, external in rels.values():
        if rel_type.endswith("/threadedComment") and not external:
            for element in _parse(parts, target):
                if _local(element)[0] == "threadedComment" and element.get("ref"):
                    threaded.add(element.get("ref", ""))
    for cell in sorted(threaded):
        _cell_report(collector, "comment.threaded", cell, ref)
    for rel_type, target, external in rels.values():
        if rel_type.endswith("/comments") and not external:
            for element in _parse(parts, target):
                cell = element.get("ref", "")
                if _local(element)[0] == "comment" and cell and cell not in threaded:
                    _cell_report(collector, "comment", cell, ref)


def _cell_report(collector: Collector, code: str, cell: str, ref: CellRange | None) -> None:
    try:
        cells = CellRange.parse(cell)
    except ValueError:
        return
    if ref is None or cells.intersects(ref):
        _report(collector, code, cells=cells)


def _parse(parts: xlio.Parts, part: str) -> Iterator[ET.Element]:
    try:
        root = ET.fromstring(parts.read(part))
    except (KeyError, ET.ParseError):
        return iter(())
    return root.iter()


def _children(elements: Iterable[ET.Element]) -> Iterator[ET.Element]:
    """Elements, with an ``mc:AlternateContent`` replaced by its first choice's children."""
    for element in elements:
        if element.tag == f"{{{_MC}}}AlternateContent":
            choice = element.find(f"{{{_MC}}}Choice")
            if choice is not None:
                yield from _children(list(choice))
        else:
            yield element


def _local(element: ET.Element) -> tuple[str, str]:
    """(local name, namespace) of an element."""
    tag = element.tag
    if tag.startswith("{"):
        namespace, _, local = tag[1:].partition("}")
        return local, namespace
    return tag, ""


def _true(value: str | None) -> bool:
    return value in ("1", "true")


def _number(value: str | None) -> bool:
    try:
        return bool(value) and float(value or 0) > 0
    except ValueError:
        return False
