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
import zipfile
from collections.abc import Callable, Iterable, Iterator
from dataclasses import dataclass, field
from typing import Any

from sheetpeek._model import CellRange
from sheetpeek._notices import Collector, Outcome, Reason
from sheetpeek._package import related, workbook_part
from sheetpeek._structure import PHONETIC, RICH, Structure
from sheetpeek._style import Code

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main"
_CHUNK = 1 << 20

# What each reader code reports: outcome, reason, and what is drawn instead (or why not).
CODES: dict[str, tuple[Outcome, Reason, str]] = {
    "border": ("omitted", "not_implemented", ""),
    "border.diagonal": ("omitted", "not_implemented", ""),
    "fill.pattern": ("approximated", "not_implemented", "drawn as the pattern color, solid"),
    "fill.gradient": ("omitted", "not_implemented", ""),
    "font.vert_align": ("approximated", "not_implemented", "drawn at full size on the baseline"),
    "font.scheme": ("approximated", "not_implemented", ""),
    "text.rotation": ("approximated", "not_implemented", "drawn horizontally"),
    "text.stacked": ("approximated", "not_implemented", "drawn horizontally"),
    "text.shrink": ("approximated", "not_implemented", "drawn at full size"),
    "text.reading_order": ("omitted", "not_implemented", ""),
    "text.runs": ("approximated", "not_implemented", "drawn in the cell's font"),
    "text.phonetic": ("omitted", "not_implemented", ""),
    "style.row_column": (
        "omitted",
        "not_implemented",
        "fills and borders of whole rows and columns, on cells the file has no element for",
    ),
    "view.hide_zeros": ("approximated", "not_implemented", "zeros are shown; Excel hides them"),
    "view.right_to_left": ("approximated", "not_implemented", "drawn left to right"),
    "view.formulas": (
        "approximated",
        "not_implemented",
        "values are shown; Excel shows the formulas",
    ),
    "view.page_layout": ("approximated", "not_implemented", "drawn as the Normal view"),
    "view.panes": ("omitted", "not_implemented", ""),
    "grid.color": ("approximated", "not_implemented", "drawn in the default gray"),
    "outline": ("omitted", "not_implemented", ""),
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
    "print.area": ("omitted", "not_implemented", ""),
    "print.titles": ("omitted", "not_implemented", ""),
    "print.breaks": ("omitted", "not_implemented", ""),
    "sheet.tab_color": ("omitted", "not_implemented", ""),
    "sheet.chartsheet": ("omitted", "not_implemented", ""),
    "sheet.dialogsheet": (
        "omitted",
        "not_implemented",
        "a legacy dialog: its frame and controls are not drawn",
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
# Children the model reads (sheetPr and sheetViews are also checked below).
_READ = frozenset(
    {"sheetPr", "sheetViews", "sheetFormatPr", "cols", "mergeCells", "hyperlinks", "sheetData"}
)

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
    print_areas: frozenset[str] = frozenset()
    print_titles: frozenset[str] = frozenset()
    show_objects: str = "all"
    strings_part: str | None = None


def workbook_info(archive: zipfile.ZipFile) -> WorkbookInfo:
    """Read sheet kinds, print names, ``showObjects`` and the shared strings part."""
    part = workbook_part(archive)
    rels = {rid: (kind, target) for rid, kind, target, _ in related(archive, part)}
    root = ET.fromstring(archive.read(part))
    kinds: dict[str, str] = {}
    names: list[str] = []
    for sheet in root.iter(f"{{{_MAIN}}}sheet"):
        name = sheet.get("name", "")
        names.append(name)
        kind_url = rels.get(sheet.get(f"{{{_REL}}}id", ""), ("", ""))[0]
        kinds[name] = next(
            (k for suffix, k in _SHEET_KINDS if suffix in kind_url.lower()), "worksheet"
        )
    areas: set[str] = set()
    titles: set[str] = set()
    for defined in root.iter(f"{{{_MAIN}}}definedName"):
        local = defined.get("localSheetId", "")
        if not local.isdigit() or int(local) >= len(names) or not (defined.text or "").strip():
            continue
        if defined.get("name") == "_xlnm.Print_Area":
            areas.add(names[int(local)])
        elif defined.get("name") == "_xlnm.Print_Titles":
            titles.add(names[int(local)])
    properties = root.find(f"{{{_MAIN}}}workbookPr")
    show_objects = properties.get("showObjects", "all") if properties is not None else "all"
    strings = next(
        (target for kind, target in rels.values() if kind.endswith("/sharedStrings")), None
    )
    return WorkbookInfo(kinds, frozenset(areas), frozenset(titles), show_objects, strings)


def book_notices(collector: Collector, info: WorkbookInfo) -> None:
    """Workbook-level notices: sheets that cannot be shown as worksheets."""
    for name, kind in info.kinds.items():
        if kind == "chartsheet":
            _report(collector, "sheet.chartsheet", context={"sheet": name}, detail=f"{name!r}")


def string_marks(archive: zipfile.ZipFile, part: str | None) -> bytes | None:
    """Per shared string, :data:`RICH` and :data:`PHONETIC` bits; ``None`` when none is set.

    A byte search skips the parse for the many workbooks whose strings have neither.
    """
    if part is None or part not in archive.namelist():
        return None
    found = False
    tail = b""
    with archive.open(part) as handle:
        while chunk := handle.read(_CHUNK):
            data = tail + chunk
            if b"<r>" in data or b":r>" in data or b"rPh" in data:
                found = True
                break
            tail = data[-8:]
    if not found:
        return None
    marks = bytearray()
    with archive.open(part) as handle:
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


def theme_fonts(theme: bytes | None) -> dict[str, str]:
    """The theme's Latin typefaces: ``{"major": ..., "minor": ...}``."""
    if not theme:
        return {}
    try:
        root = ET.fromstring(theme)
    except ET.ParseError:
        return {}
    result = {}
    for scheme, key in (("majorFont", "major"), ("minorFont", "minor")):
        latin = root.find(f".//{{{_DRAWING}}}{scheme}/{{{_DRAWING}}}latin")
        if latin is not None and latin.get("typeface"):
            result[key] = latin.get("typeface", "")
    return result


def visible_style(workbook: Any, index: int) -> bool:
    """Whether a row or column style draws on empty cells: a fill or a border side."""
    styles = getattr(workbook, "_cell_styles", None)
    if not styles or index >= len(styles):
        return False
    array = styles[index]
    fills = getattr(workbook, "_fills", ())
    borders = getattr(workbook, "_borders", ())
    fill = fills[array.fillId] if array.fillId < len(fills) else None
    if fill is not None and (
        getattr(fill, "tagname", None) == "gradientFill"
        or getattr(fill, "fill_type", None) not in (None, "none")
    ):
        return True
    border = borders[array.borderId] if array.borderId < len(borders) else None
    sides = ("left", "right", "top", "bottom")
    return border is not None and any(
        getattr(getattr(border, s, None), "style", None) for s in sides
    )


@dataclass(slots=True)
class SheetContext:
    """What a sheet's rules need besides the element: the package and what was found so far."""

    collector: Collector
    archive: zipfile.ZipFile
    rels: dict[str, tuple[str, str, bool]]
    objects: bool
    """Whether the workbook shows objects (``showObjects`` is not ``none``)."""
    symbols: bool | None = None
    """``showOutlineSymbols`` when the file states it."""


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
    archive: zipfile.ZipFile,
    part: str,
    structure: Structure,
    info: WorkbookInfo,
    name: str,
    ref: CellRange | None,
) -> None:
    """Report what one sheet uses at sheet level and from its other parts."""
    kind = info.kinds.get(name, "worksheet")
    if kind in ("dialogsheet", "macrosheet"):
        _report(collector, f"sheet.{kind}")
    context = SheetContext(
        collector,
        archive,
        {rid: (kind, target, external) for rid, kind, target, external in related(archive, part)},
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
    if (structure.outline_rows or structure.outline_columns) and context.symbols is not False:
        _report(
            collector,
            "outline",
            context={"rows": structure.outline_rows, "columns": structure.outline_columns},
        )
    if name in info.print_areas:
        _report(collector, "print.area")
    if name in info.print_titles:
        _report(collector, "print.titles")
    _comments(collector, archive, context.rels, ref)


@rule("sheetPr")
def _sheet_properties(context: SheetContext, child: ET.Element) -> None:
    tab = child.find(f"{{{_MAIN}}}tabColor")
    if tab is not None and not _true(tab.get("auto")) and tab.attrib:
        _report(context.collector, "sheet.tab_color")
    outline = child.find(f"{{{_MAIN}}}outlinePr")
    if outline is not None and context.symbols is None and outline.get("showOutlineSymbols"):
        context.symbols = _true(outline.get("showOutlineSymbols"))


@rule("sheetViews")
def _sheet_views(context: SheetContext, child: ET.Element) -> None:
    view = child.find(f"{{{_MAIN}}}sheetView")
    if view is not None:
        if view.get("showOutlineSymbols") is not None:
            context.symbols = _true(view.get("showOutlineSymbols"))
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
        _drawing(context.collector, context.archive, target[1])
    else:
        _vml(context.collector, context.archive, target[1])


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


@rule("rowBreaks", "colBreaks")
def _breaks(context: SheetContext, child: ET.Element) -> None:
    if manual := sum(1 for b in child.iter(f"{{{_MAIN}}}brk") if _true(b.get("man"))):
        _report(context.collector, "print.breaks", count=manual)


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
    if view.get("showZeros") is not None and not _true(view.get("showZeros")):
        _report(collector, "view.hide_zeros")
    if _true(view.get("rightToLeft")):
        _report(collector, "view.right_to_left")
    if _true(view.get("showFormulas")):
        _report(collector, "view.formulas")
    if view.get("view") in ("pageLayout", "pageBreakPreview"):
        _report(collector, "view.page_layout", context={"view": view.get("view", "")})
    color = view.get("colorId")
    default_color = view.get("defaultGridColor")
    if default_color is not None and not _true(default_color) and color not in (None, "64"):
        _report(collector, "grid.color", context={"color_id": color or ""})
    pane = view.find(f"{{{_MAIN}}}pane")
    if pane is not None:
        state = pane.get("state", "split")
        if (
            state in ("frozen", "frozenSplit")
            or _number(pane.get("xSplit"))
            or _number(pane.get("ySplit"))
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


def _drawing(collector: Collector, archive: zipfile.ZipFile, part: str) -> None:
    try:
        root = ET.fromstring(archive.read(part))
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


def _vml(collector: Collector, archive: zipfile.ZipFile, part: str) -> None:
    """Form controls in a legacy (VML) drawing; notes are reported from the comments part.

    VML is often not well-formed XML, so it is searched, not parsed.
    """
    try:
        data = archive.read(part)
    except KeyError:
        return
    controls = sum(1 for kind in _OBJECT_TYPE.findall(data) if kind != b"Note")
    if controls:
        _report(collector, "control", count=controls)


def _comments(
    collector: Collector,
    archive: zipfile.ZipFile,
    rels: dict[str, tuple[str, str, bool]],
    ref: CellRange | None,
) -> None:
    threaded: set[str] = set()
    for rel_type, target, external in rels.values():
        if rel_type.endswith("/threadedComment") and not external:
            for element in _parse(archive, target):
                if _local(element)[0] == "threadedComment" and element.get("ref"):
                    threaded.add(element.get("ref", ""))
    for cell in sorted(threaded):
        _cell_report(collector, "comment.threaded", cell, ref)
    for rel_type, target, external in rels.values():
        if rel_type.endswith("/comments") and not external:
            for element in _parse(archive, target):
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


def _parse(archive: zipfile.ZipFile, part: str) -> Iterator[ET.Element]:
    try:
        root = ET.fromstring(archive.read(part))
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
