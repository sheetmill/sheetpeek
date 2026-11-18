"""Hyperlinks: ``<hyperlink>`` elements and ``HYPERLINK()`` formulas.

xlio reads the elements (:func:`with_elements`); the reader finds the formulas in xlio's cell
stream (:func:`add_formula`). Both are kept in ``Structure.extra`` under :data:`KEY`; the reader
then gives each cell its
:class:`~sheetpeek.Hyperlink` and reports what cannot be linked (``ARCHITECTURE.md`` §4.10).
"""

from __future__ import annotations

import contextlib
import dataclasses
import re
from collections.abc import Sequence
from typing import TYPE_CHECKING

import xlio

from sheetpeek._hyperlink import LOCAL_TARGETS, Hyperlink, check_target
from sheetpeek._model import CellRange
from sheetpeek._structure import Structure

if TYPE_CHECKING:
    from sheetpeek._model import Cell
    from sheetpeek._notices import Collector

KEY = "links"
"""Where what is found is kept, in ``Structure.extra``."""

_STRING = re.compile(r'"(?:[^"]|"")*"')
_CALL = re.compile(r"(?<![\w.])HYPERLINK\s*\(", re.IGNORECASE)
_LITERAL = re.compile(r'\s*HYPERLINK\s*\(\s*"((?:[^"]|"")*)"\s*[,)]', re.IGNORECASE)
_LABEL = 24  # characters of a refused scheme kept in a message


class Found:
    """The links found in one worksheet part."""

    __slots__ = ("computed", "elements", "formulas", "shared")

    def __init__(self) -> None:
        self.elements: list[tuple[CellRange, Hyperlink, bool]] = []
        """Each ``<hyperlink>``: its cells, its link, and whether its relationship is a part of
        the package rather than an external target."""
        self.formulas: dict[tuple[int, int], Hyperlink] = {}
        """Cells whose formula is ``HYPERLINK()`` with a literal target."""
        self.computed: list[tuple[CellRange, str]] = []
        """Cells with a ``HYPERLINK()`` that cannot be linked, and why: ``computed``,
        ``shared`` or ``array``."""
        self.shared: set[int | str] = set()
        """The ``si`` of shared formulas whose master calls ``HYPERLINK()`` (``""`` for none)."""


def with_elements(structure: Structure, sheet: xlio.Sheet) -> Structure:
    """``structure`` with the sheet's ``<hyperlink>`` elements.

    Each with its range, target, location, tooltip and display text. A relationship to a part of
    the package is nothing a link can open.
    """
    elements: list[tuple[CellRange, Hyperlink, bool]] = []
    parts = {r.id for r in sheet.relationships if not r.external}
    for element in sheet.hyperlinks:
        if not element.ref:
            continue
        part = element.target is None and element.relationship_id in parts
        location = element.location or None
        if element.target is None and location is None and not part:
            continue
        link = Hyperlink(element.target, location, element.tooltip, element.display)
        elements.append((CellRange.parse(element.ref), link, part))
    if not elements:
        return structure
    found = structure.extra.get(KEY)
    if not isinstance(found, Found):
        found = Found()
    found.elements = elements
    return dataclasses.replace(structure, extra={**structure.extra, KEY: found})


def add_formula(found: Found, formula: xlio.Formula, row: int, column: int) -> None:
    """Note a cell's formula if it calls ``HYPERLINK()``, in the order of the cell stream.

    A shared formula's master comes before the cells that share it, so they are known by its
    ``si``.
    """
    text = formula.text or ""
    kind = formula.kind
    cell = CellRange(row, column, row, column)
    if kind == "shared" and not text.strip():
        if formula.si is not None and formula.si in found.shared:
            found.computed.append((cell, "shared"))
        return
    if "HYPERLINK" not in text.upper() or not _CALL.search(_STRING.sub('""', text)):
        return
    if kind == "shared":
        found.shared.add(formula.si if formula.si is not None else "")
        found.computed.append((cell, "shared"))
    elif kind == "array":
        if formula.ref:
            with contextlib.suppress(ValueError):
                cell = CellRange.parse(formula.ref)
        found.computed.append((cell, "array"))
    elif (target := literal_target(text)) is not None:
        found.formulas[(row, column)] = Hyperlink(target, source="formula")
    else:
        found.computed.append((cell, "computed"))


def with_formulas(structure: Structure, found: Found) -> Structure:
    """``structure`` with the ``HYPERLINK()`` formulas ``found`` in the cell stream."""
    if not found.formulas and not found.computed:
        return structure
    known = structure.extra.get(KEY)
    if isinstance(known, Found):
        found.elements = known.elements
    return dataclasses.replace(structure, extra={**structure.extra, KEY: found})


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
