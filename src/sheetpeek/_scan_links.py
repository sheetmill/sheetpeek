"""Hyperlinks in the sheet XML: ``<hyperlink>`` elements and ``HYPERLINK()`` formulas.

The structure scan's handlers keep what they find in ``Scan.extra`` under :data:`KEY`; the reader
then gives each cell its :class:`~sheetpeek.Hyperlink` and reports what cannot be linked
(``ARCHITECTURE.md`` §4.10).
"""

from __future__ import annotations

import contextlib
import re
import xml.etree.ElementTree as ET
from collections.abc import Sequence
from typing import TYPE_CHECKING

from sheetpeek._hyperlink import LOCAL_TARGETS, Hyperlink, check_target
from sheetpeek._model import CellRange
from sheetpeek._structure import Scan, Structure, handler

if TYPE_CHECKING:
    from sheetpeek._model import Cell
    from sheetpeek._notices import Collector

KEY = "links"
"""Where the handlers keep what they find, in ``Scan.extra`` and ``Structure.extra``."""

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_HYPERLINK = f"{_NS}hyperlink"
_FORMULA = f"{_NS}f"
_ROW = f"{_NS}row"
_STRING = re.compile(r'"(?:[^"]|"")*"')
_CALL = re.compile(r"(?<![\w.])HYPERLINK\s*\(", re.IGNORECASE)
_LITERAL = re.compile(r'\s*HYPERLINK\s*\(\s*"((?:[^"]|"")*)"\s*[,)]', re.IGNORECASE)
_LABEL = 24  # characters of a refused scheme kept in a message


class Found:
    """What the handlers found in one worksheet part."""

    __slots__ = ("computed", "elements", "formulas", "pending", "shared")

    def __init__(self) -> None:
        self.elements: list[tuple[CellRange, Hyperlink, bool]] = []
        """Each ``<hyperlink>``: its cells, its link, and whether its relationship is a part of
        the package rather than an external target."""
        self.formulas: dict[tuple[int, int], Hyperlink] = {}
        """Cells whose formula is ``HYPERLINK()`` with a literal target."""
        self.computed: list[tuple[CellRange, str]] = []
        """Cells with a ``HYPERLINK()`` that cannot be linked, and why: ``computed``,
        ``shared`` or ``array``."""
        self.shared: set[str] = set()
        """The ``si`` of shared formulas whose master calls ``HYPERLINK()``."""
        self.pending: dict[int, tuple[ET.Element, str, Hyperlink | None]] = {}
        """Formulas of the current row, by the element's ``id``, until the row places them."""


def _found(state: Scan) -> Found:
    found = state.extra.get(KEY)
    if found is None:
        found = state.extra[KEY] = Found()
    return found


@handler(_HYPERLINK)
def _element(state: Scan, element: ET.Element) -> None:
    """A ``<hyperlink>``: its range, target, location, tooltip and display text."""
    ref = element.get("ref")
    if not ref:
        return
    target: str | None = None
    part = False
    rel_id = element.get(_REL_ID)
    if rel_id and rel_id in state.rels:
        target, external = state.rels[rel_id]
        if not external:  # a part of the package: nothing a link can open
            target, part = None, True
    location = element.get("location") or None
    if target is None and location is None and not part:
        return
    link = Hyperlink(target, location, element.get("tooltip"), element.get("display"))
    _found(state).elements.append((CellRange.parse(ref), link, part))


@handler(_FORMULA)
def _formula(state: Scan, element: ET.Element) -> None:
    """A formula that calls ``HYPERLINK()``, placed when its row ends."""
    text = element.text or ""
    kind = element.get("t")
    if kind == "shared" and not text.strip():
        found = state.extra.get(KEY)
        if found is not None and element.get("si") in found.shared:
            found.pending[id(element)] = (element, "shared", None)
        return
    if "HYPERLINK" not in text.upper() or not _CALL.search(_STRING.sub('""', text)):
        return
    found = _found(state)
    if kind == "shared":
        found.shared.add(element.get("si") or "")
        found.pending[id(element)] = (element, "shared", None)
    elif kind == "array":
        found.pending[id(element)] = (element, "array", None)
    elif (target := literal_target(text)) is not None:
        found.pending[id(element)] = (element, "link", Hyperlink(target, source="formula"))
    else:
        found.pending[id(element)] = (element, "computed", None)


@handler(_ROW)
def _row(state: Scan, element: ET.Element) -> None:
    """Place the row's ``HYPERLINK()`` formulas, before the row is cleared."""
    found = state.extra.get(KEY)
    if found is None or not found.pending:
        return
    row = state.row_index  # set by the scan's own row handler, which runs first
    column = 0
    for cell in element:
        ref = cell.get("r")
        column = CellRange.parse(ref).min_col if ref else column + 1
        formula = cell.find(_FORMULA)
        entry = found.pending.get(id(formula)) if formula is not None else None
        if entry is None:
            continue
        formula, form, link = entry
        if link is not None:
            found.formulas[(row, column)] = link
            continue
        cells = CellRange(row, column, row, column)
        if form == "array" and (span := formula.get("ref")):
            with contextlib.suppress(ValueError):
                cells = CellRange.parse(span)
        found.computed.append((cells, form))
    found.pending.clear()


def literal_target(formula: str) -> str | None:
    """The target of a formula that is one ``HYPERLINK()`` call with a string literal first.

    ``HYPERLINK("https://example.com/", B2)`` gives the URL, its doubled quotes made single;
    ``HYPERLINK(A1)``, ``HYPERLINK("a" & B1)`` and ``IF(x, HYPERLINK("a"))`` give ``None``.
    """
    match = _LITERAL.match(formula)
    if match is None or not _whole_call(formula):
        return None
    return match[1].replace('""', '"')


def _whole_call(formula: str) -> bool:
    """Whether the first parenthesis closes at the formula's end, string literals aside."""
    bare = _STRING.sub('""', formula).strip()
    depth = 0
    for i, char in enumerate(bare):
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
            if depth == 0:
                return i == len(bare) - 1
    return False


def cell_links(structure: Structure, ref: CellRange) -> dict[tuple[int, int], Hyperlink]:
    """Each linked cell of ``ref`` and its link; an element's link wins over a formula's."""
    found = structure.extra.get(KEY)
    if not isinstance(found, Found):
        return {}
    links: dict[tuple[int, int], Hyperlink] = {}
    for area, link, _ in found.elements:
        if link.target is None and link.location is None:
            continue
        _cover(links, area, ref, link)
    for position, link in found.formulas.items():
        if position in ref:
            links.setdefault(position, link)
    return links


def _cover(
    links: dict[tuple[int, int], Hyperlink], area: CellRange, ref: CellRange, link: Hyperlink
) -> None:
    clipped = area.intersection(ref)
    if clipped is None:
        return
    for row in range(clipped.min_row, clipped.max_row + 1):
        for column in range(clipped.min_col, clipped.max_col + 1):
            links.setdefault((row, column), link)


def report(
    notes: Collector,
    structure: Structure,
    ref: CellRange,
    cells: Sequence[Sequence[Cell]],
) -> None:
    """Report the links of ``ref`` that no output can follow, and the conflicts."""
    found = structure.extra.get(KEY)
    if not isinstance(found, Found) or not cells:
        return
    effective: dict[tuple[int, int], tuple[Hyperlink, bool]] = {}
    for area, link, part in found.elements:
        clipped = area.intersection(ref)
        if clipped is None:
            continue
        for row in range(clipped.min_row, clipped.max_row + 1):
            for column in range(clipped.min_col, clipped.max_col + 1):
                effective.setdefault((row, column), (link, part))
    for position, link in found.formulas.items():
        if position in ref:
            if position in effective:
                _conflict(notes, position)
            else:
                effective[position] = (link, False)
    for area, form in found.computed:
        clipped = area.intersection(ref)
        if clipped is None:
            continue
        for row in range(clipped.min_row, clipped.max_row + 1):
            for column in range(clipped.min_col, clipped.max_col + 1):
                if (row, column) in effective:
                    _conflict(notes, (row, column))
                else:
                    notes.add(
                        "link.computed",
                        "omitted",
                        "not_implemented",
                        row=row,
                        column=column,
                        detail=_COMPUTED[form],
                        context={"form": form},
                    )
    empty = {p for p in effective if _shows_nothing(cells[p[0] - ref.min_row][p[1] - ref.min_col])}
    covered = _covered(empty, structure.merges, ref) if empty and structure.merges else set()
    for position, (link, part) in effective.items():
        if position in covered:
            continue
        row, column = position
        if position in empty:
            notes.add(
                "link.empty_cell",
                "omitted",
                "not_representable",
                row=row,
                column=column,
                detail="the cell shows no text, so there is nothing to click",
            )
        elif part and link.location is None:
            _refused(notes, position, "part", "a part of the package, not a URL")
        elif link.target is not None:
            url, scheme = check_target(link.target)
            if url is None and scheme not in LOCAL_TARGETS:  # local ones: each output reports
                label = scheme if len(scheme) <= _LABEL else scheme[:_LABEL] + "..."
                why = (
                    "a relative or file path"
                    if scheme in ("relative", "path")
                    else "no URL"
                    if scheme == "unparseable"
                    else f'scheme "{label}"'
                )
                _refused(notes, position, label, why)


_COMPUTED = {
    "computed": "HYPERLINK() with a computed target; needs a formula engine",
    "shared": "HYPERLINK() in a shared formula, whose target is not read",
    "array": "HYPERLINK() in an array formula, whose target is not read",
}


def _refused(notes: Collector, position: tuple[int, int], scheme: str, why: str) -> None:
    notes.add(
        "link.scheme",
        "omitted",
        "unsafe",
        row=position[0],
        column=position[1],
        detail=f"{why} is refused; the text is drawn without the link",
        context={"scheme": scheme},
    )


def _conflict(notes: Collector, position: tuple[int, int]) -> None:
    notes.add(
        "link.conflict",
        "approximated",
        "file",
        row=position[0],
        column=position[1],
        detail="a <hyperlink> element and a HYPERLINK() formula: the element's target is used",
    )


def _shows_nothing(cell: Cell) -> bool:
    return cell.value is None or cell.value == ""


def _covered(
    empty: set[tuple[int, int]], merges: Sequence[CellRange], ref: CellRange
) -> set[tuple[int, int]]:
    """The positions of ``empty`` that a merge covers (its anchor is not covered)."""
    covered: set[tuple[int, int]] = set()
    for merge in merges:
        clipped = merge.intersection(ref)
        if clipped is None:
            continue
        anchor = (merge.min_row, merge.min_col)
        if clipped.rows * clipped.columns <= len(empty):
            covered.update(
                (row, column)
                for row in range(clipped.min_row, clipped.max_row + 1)
                for column in range(clipped.min_col, clipped.max_col + 1)
                if (row, column) in empty
            )
        else:
            covered.update(p for p in empty if p in clipped)
        covered.discard(anchor)
    return covered
