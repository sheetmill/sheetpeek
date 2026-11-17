"""Notices: what a rendering leaves out or approximates, reported as data.

Three producers report through notices, so that nothing a workbook uses is dropped silently:

- the reader (``Sheet.notices``): what the file uses that sheetpeek does not read yet;
- the layout (``RenderedSheet.notices``): what it approximates in every output;
- each output (:func:`sheetpeek.notices`): what that output cannot show, from the layout's
  census of the features present and the backend's declared :class:`Support` for each.

Every code is registered in :data:`FEATURES`. ``ARCHITECTURE.md`` has the design.
"""

from __future__ import annotations

import dataclasses
import warnings
from collections.abc import Callable, Iterable, Mapping
from dataclasses import dataclass, field
from typing import Literal

from sheetpeek._model import CellRange

Outcome = Literal["omitted", "approximated"]
"""``omitted``: nothing of the feature is drawn; ``approximated``: something else is drawn."""

Reason = Literal[
    "not_implemented", "not_representable", "needs_dependency", "option", "unsafe", "file"
]
"""Why: sheetpeek does not do it yet; this output cannot show it; an optional package would;
the caller's option decided it; it is refused on purpose; or the file states something Excel
would not show either."""

Level = Literal["workbook", "sheet", "cells"]

_MAX_RANGES = 1_000
_MAX_RUNS = 100_000


class NotShownWarning(UserWarning):
    """What :func:`sheetpeek.render` issues once when something is not shown.

    It is not issued when the caller handles notices with ``on_notice``. Filter it like any
    warning.
    """


@dataclass(frozen=True, slots=True)
class Feature:
    """A registered feature: what a notice's code names."""

    code: str
    name: str
    """A plural noun phrase for messages, such as ``"borders"``."""
    level: Level
    source: Literal["reader", "layout", "census"]
    """Who reports it, the reader, the layout, or each output through the census."""


@dataclass(frozen=True, slots=True)
class Notice:
    """Something a rendering leaves out or approximates.

    Notices are aggregated: one per code, sheet, output and kind, never one per cell.
    """

    code: str
    """A registered code, such as ``"border.style"``; see :data:`sheetpeek.FEATURES`."""
    outcome: Outcome
    reason: Reason
    message: str
    """One line, complete by itself."""
    sheet: str | None = None
    """The sheet, or ``None`` for the workbook."""
    cells: tuple[CellRange, ...] = ()
    """Where, as ranges; empty for the whole sheet (or workbook). At most 1,000 ranges; then
    ``context["truncated"]`` is 1 and ``count`` stays exact."""
    count: int = 0
    """Cells affected, or the number of things (rules, objects, sheets) for a sheet-level notice."""
    output: str | None = None
    """The output that decided it (a backend name), or ``None`` when the model did."""
    context: Mapping[str, str | int | float] = field(default_factory=dict, hash=False)
    """The values in the message, as data."""
    dependency: str | None = None
    """The optional package, when ``reason`` is ``"needs_dependency"``."""

    @property
    def level(self) -> Level:
        """The level it applies at, ``workbook``, ``sheet`` or ``cells``."""
        if self.sheet is None:
            return "workbook"
        feature = FEATURES.get(self.code)
        return "cells" if feature is not None and feature.level == "cells" else "sheet"

    def matches(self, prefix: str) -> bool:
        """Whether the code is ``prefix`` or under it: ``"border"`` matches ``"border.style"``."""
        return self.code == prefix or self.code.startswith(prefix + ".")


@dataclass(frozen=True, slots=True)
class Support:
    """What an output does with one feature; a backend declares one per code in ``support``.

    Use :data:`SHOWN`, :meth:`approximated` or :meth:`omitted`.
    """

    outcome: Outcome | None
    reason: Reason | None = None
    how: str = ""
    """What is drawn instead (approximated), or why not (omitted)."""

    @classmethod
    def approximated(cls, how: str, reason: Reason = "not_representable") -> Support:
        """Something else is drawn: ``how`` says what, such as ``"drawn as plain text"``."""
        return cls("approximated", reason, how)

    @classmethod
    def omitted(cls, reason: Reason = "not_representable", why: str = "") -> Support:
        """Nothing of the feature is drawn."""
        return cls("omitted", reason, why)


SHOWN = Support(None)
"""The output shows the feature as Excel does."""


FEATURES: dict[str, Feature] = {}
"""Every notice code, with the feature it names."""


Source = Literal["reader", "layout", "census"]


def _register(code: str, name: str, level: Level, source: Source) -> None:
    FEATURES[code] = Feature(code, name, level, source)


# What the model carries: each output declares its support for these.
_level: Level
for _code, _name, _level in (
    ("font.name", "font names", "cells"),
    ("font.size", "font sizes", "cells"),
    ("font.bold", "bold text", "cells"),
    ("font.italic", "italic text", "cells"),
    ("text.runs", "rich-text runs", "cells"),
    ("font.underline", "underlines", "cells"),
    ("font.underline_accounting", "accounting underlines", "cells"),
    ("font.strike", "strikethrough", "cells"),
    ("font.vert_align", "superscript and subscript", "cells"),
    ("text.shrink", "shrink to fit", "cells"),
    ("text.reading_order", "reading order", "cells"),
    ("text.rotation", "rotated text", "cells"),
    ("text.stacked", "stacked text", "cells"),
    ("font.color", "font colors", "cells"),
    ("grid.color", "gridline colors", "sheet"),
    ("format.color", "number-format colors", "cells"),
    ("fill", "fills", "cells"),
    ("fill.pattern", "pattern fills", "cells"),
    ("fill.gradient", "gradient fills", "cells"),
    ("border", "borders", "cells"),
    ("border.diagonal", "diagonal borders", "cells"),
    ("border.slant_dash_dot", "slanted dash-dot borders", "cells"),
    ("border.slanted", "borders of angled text in a merge", "cells"),
    ("align.horizontal", "horizontal alignment", "cells"),
    ("align.vertical", "vertical alignment", "cells"),
    ("align.wrap", "wrapped text", "cells"),
    ("align.indent", "indents", "cells"),
    ("link", "hyperlinks", "cells"),
    ("link.internal", "links within the rendered cells", "cells"),
    ("link.elsewhere", "links to places outside the rendered cells", "cells"),
    ("link.local", "links to local files and relative paths", "cells"),
    ("merge", "merged cells", "sheet"),
    ("column.width", "column widths", "sheet"),
    ("row.height", "row heights", "sheet"),
    ("gridlines.hidden", "hidden gridlines", "sheet"),
    ("headings.hidden", "hidden row and column headings", "sheet"),
    ("outline", "outline symbols", "sheet"),
):
    _register(_code, _name, _level, "census")

# What the reader finds in the file and does not read yet.
for _code, _name, _level in (
    ("value.no_cached_value", "formulas without a cached value", "cells"),
    ("format.palette", "number-format colors from the workbook palette", "cells"),
    ("text.phonetic", "phonetic guides", "cells"),
    ("link.scheme", "link targets", "cells"),
    ("link.empty_cell", "links on empty cells", "cells"),
    ("link.computed", "HYPERLINK() formulas", "cells"),
    ("link.conflict", "cells with two links", "cells"),
    ("style.row_column", "row and column formats", "sheet"),
    ("view.hide_zeros", "hidden zero values", "sheet"),
    ("view.right_to_left", "right-to-left sheets", "sheet"),
    ("view.formulas", "formulas shown in cells", "sheet"),
    ("view.page_layout", "page layout and page break views", "sheet"),
    ("view.panes", "split panes", "sheet"),
    ("outline.level", "outline levels outside 0 to 7", "sheet"),
    ("cf", "conditional formatting", "sheet"),
    ("filter.buttons", "filter buttons", "sheet"),
    ("table", "tables", "sheet"),
    ("sparkline", "sparklines", "sheet"),
    ("slicer", "slicers", "sheet"),
    ("timeline", "timelines", "sheet"),
    ("drawing.picture", "pictures", "sheet"),
    ("drawing.picture_linked", "linked pictures", "sheet"),
    ("drawing.chart", "charts", "sheet"),
    ("drawing.shape", "shapes and text boxes", "sheet"),
    ("drawing.other", "other drawn objects", "sheet"),
    ("drawing.ole", "embedded objects", "sheet"),
    ("control", "form controls", "sheet"),
    ("picture.background", "background pictures", "sheet"),
    ("comment", "notes", "cells"),
    ("comment.threaded", "threaded comments", "cells"),
    ("sheet.tab_color", "tab colors", "sheet"),
    ("print.area_unusable", "parts of print areas", "sheet"),
    ("sheet.chartsheet", "chartsheets", "workbook"),
    ("sheet.dialogsheet", "dialog sheets", "workbook"),
    ("sheet.macrosheet", "macro sheets", "sheet"),
    ("sheet.hidden", "hidden sheets", "workbook"),
    ("sheet.unknown_element", "unrecognized sheet elements", "sheet"),
    ("sheet.unknown_extension", "unrecognized sheet extensions", "sheet"),
):
    _register(_code, _name, _level, "reader")

# What the layout approximates in every output.
for _code, _name, _level in (
    ("font.no_metrics", "fonts without glyph tables", "sheet"),
    ("format.refused", "number formats Excel refuses", "cells"),
    ("value.unshowable", "values no format can show", "cells"),
    ("rows.hidden_shown", "hidden rows", "sheet"),
    ("columns.hidden_shown", "hidden columns", "sheet"),
    ("outline.cut", "outline groups cut by the range", "sheet"),
    ("text.stacked_wrapped", "wrapped stacked text", "cells"),
):
    _register(_code, _name, _level, "layout")

LEGACY = frozenset(
    {
        "value.no_cached_value",
        "font.no_metrics",
        "format.refused",
        "value.unshowable",
    }
)
"""The codes of the warnings before notices; ``warnings`` keeps their messages."""


class CellSet:
    """Cells accumulated in row-major order, kept as runs and coalesced into ranges."""

    __slots__ = ("_rects", "_rows", "_runs", "count", "truncated")

    def __init__(self) -> None:
        self.count = 0
        self.truncated = False
        self._rows: dict[int, list[list[int]]] = {}
        self._rects: list[CellRange] = []
        self._runs = 0

    def add(self, row: int, column: int) -> None:
        """Add one cell."""
        self.count += 1
        if self.truncated:
            return
        runs = self._rows.get(row)
        if runs is None:
            self._rows[row] = [[column, column]]
        else:
            last = runs[-1]
            if column == last[1] + 1:
                last[1] = column
                return
            runs.append([column, column])
        self._runs += 1
        if self._runs > _MAX_RUNS:
            self.truncated = True

    def add_range(self, cells: CellRange) -> None:
        """Add a rectangle of cells."""
        self.count += (cells.max_row - cells.min_row + 1) * (cells.max_col - cells.min_col + 1)
        if not self.truncated:
            self._rects.append(cells)
            self._runs += 1
            self.truncated = self._runs > _MAX_RUNS

    def ranges(self) -> tuple[tuple[CellRange, ...], bool]:
        """The cells as ranges, runs in adjacent rows merged; and whether they were cut short."""
        result: list[CellRange] = list(self._rects)
        open_: dict[
            tuple[int, int], list[int]
        ] = {}  # (first col, last col) → [first row, last row]
        previous = None
        for row in sorted(self._rows):
            merged: list[list[int]] = []
            for first, last in sorted(self._rows[row]):
                if merged and first <= merged[-1][1] + 1:
                    merged[-1][1] = max(merged[-1][1], last)
                else:
                    merged.append([first, last])
            keys = {(a, b) for a, b in merged}
            for key in list(open_):
                if previous is None or row != previous + 1 or key not in keys:
                    top, bottom = open_.pop(key)
                    result.append(CellRange(top, key[0], bottom, key[1]))
            for key in keys:
                if key in open_:
                    open_[key][1] = row
                else:
                    open_[key] = [row, row]
            previous = row
        for key, (top, bottom) in open_.items():
            result.append(CellRange(top, key[0], bottom, key[1]))
        result.sort(key=lambda r: (r.min_row, r.min_col))
        truncated = self.truncated or len(result) > _MAX_RANGES
        return tuple(result[:_MAX_RANGES]), truncated


@dataclass(slots=True)
class _Entry:
    outcome: Outcome
    reason: Reason
    message: str | None
    detail: str
    context: dict[str, str | int | float]
    dependency: str | None
    cells: CellSet | None
    count: int = 0


class Collector:
    """Aggregates notices for one sheet (or the workbook) and one output."""

    def __init__(self, sheet: str | None, output: str | None = None) -> None:
        self.sheet = sheet
        self.output = output
        self._entries: dict[tuple[object, ...], _Entry | Notice] = {}

    def add(
        self,
        code: str,
        outcome: Outcome,
        reason: Reason,
        *,
        row: int | None = None,
        column: int | None = None,
        cells: CellRange | None = None,
        count: int = 1,
        detail: str = "",
        message: str | None = None,
        context: Mapping[str, str | int | float] | None = None,
        dependency: str | None = None,
    ) -> None:
        """Report a feature at a cell, a range, or (neither) the sheet.

        Notices with the same code, outcome, reason, detail and context are one notice. A
        ``message`` replaces the generated one; the first given is kept.
        """
        if code not in FEATURES:
            raise KeyError(f"unregistered notice code {code!r}")
        context = dict(context or {})
        key = (code, outcome, reason, detail, tuple(sorted(context.items())))
        entry = self._entries.get(key)
        if not isinstance(entry, _Entry):
            cell_level = FEATURES[code].level == "cells"
            entry = _Entry(
                outcome,
                reason,
                message,
                detail,
                context,
                dependency,
                CellSet() if cell_level else None,
            )
            self._entries[key] = entry
        if entry.cells is not None and row is not None and column is not None:
            entry.cells.add(row, column)
        elif entry.cells is not None and cells is not None:
            entry.cells.add_range(cells)
        else:
            entry.count += count

    def extend(self, notices: Iterable[Notice]) -> None:
        """Take finished notices as they are."""
        for notice in notices:
            self._entries[("=", len(self._entries))] = notice

    def notices(self) -> tuple[Notice, ...]:
        """The aggregated notices, in the order their features were first reported."""
        result = []
        for key, entry in self._entries.items():
            if isinstance(entry, Notice):
                result.append(entry)
                continue
            code = str(key[0])
            ranges: tuple[CellRange, ...] = ()
            context = dict(entry.context)
            count = entry.count
            if entry.cells is not None:
                ranges, truncated = entry.cells.ranges()
                count = entry.cells.count
                if truncated:
                    context["truncated"] = 1
            message = entry.message or describe(
                code,
                entry.outcome,
                entry.reason,
                count,
                entry.detail,
                self.output,
                entry.dependency,
            )
            result.append(
                Notice(
                    code,
                    entry.outcome,
                    entry.reason,
                    message,
                    self.sheet,
                    ranges,
                    count,
                    self.output,
                    context,
                    entry.dependency,
                )
            )
        return tuple(result)


_REASONS: dict[Reason, str] = {
    "not_implemented": "not implemented yet",
    "not_representable": "{output} cannot show it",
    "needs_dependency": "needs {dependency}",
    "option": "by the option given",
    "unsafe": "refused as unsafe",
    "file": "as the file states it",
}


def describe(
    code: str,
    outcome: Outcome,
    reason: Reason,
    count: int,
    detail: str = "",
    output: str | None = None,
    dependency: str | None = None,
) -> str:
    """The message for a notice: ``"borders not shown in csv (14 cells): csv cannot show it"``."""
    feature = FEATURES[code]
    verb = "not shown" if outcome == "omitted" else "approximated"
    where = f" in {output}" if output else ""
    amount = ""
    if feature.level == "cells":
        amount = f" ({count} cell{'s' if count != 1 else ''})"
    elif count > 1:
        amount = f" ({count})"
    why = detail or _REASONS[reason].format(output=output or "this output", dependency=dependency)
    return f"{feature.name} {verb}{where}{amount}: {why}"


def merge(notices: Iterable[Notice]) -> tuple[Notice, ...]:
    """Notices from several blocks of one sheet (its print areas) as one list, in first order.

    Equal notices, such as a sheet-level notice each block repeats, are kept once. Notices of
    one kind about cells become one: their ranges joined, their counts added and the count in
    the message changed to match.
    """
    groups: dict[tuple[object, ...], list[Notice]] = {}
    for notice in notices:
        key: tuple[object, ...] = (notice.code, notice.outcome, notice.reason, notice.sheet)
        key += (notice.output, None if notice.level == "cells" else notice.message)
        group = groups.setdefault(key, [])
        if notice not in group:
            group.append(notice)
    result = []
    for group in groups.values():
        first = group[0]
        if len(group) == 1:
            result.append(first)
            continue
        cells = tuple(dict.fromkeys(r for n in group for r in n.cells))
        count = sum(n.count for n in group)
        old = f" ({first.count} cell{'s' if first.count != 1 else ''})"
        new = f" ({count} cell{'s' if count != 1 else ''})"
        context = dict(first.context)
        if len(cells) > _MAX_RANGES or any(n.context.get("truncated") for n in group):
            cells, context["truncated"] = cells[:_MAX_RANGES], 1
        message = first.message.replace(old, new, 1)
        result.append(
            dataclasses.replace(first, message=message, cells=cells, count=count, context=context)
        )
    return tuple(result)


def visible(
    notices: Iterable[Notice], hidden_rows: set[int], hidden_cols: set[int]
) -> list[Notice]:
    """The notices restricted to visible cells; one left with no visible cell is dropped.

    A notice whose ranges were cut short is kept whole, since its other cells are unknown.
    """
    if not hidden_rows and not hidden_cols:
        return list(notices)
    result = []
    for notice in notices:
        if not notice.cells or notice.context.get("truncated"):
            result.append(notice)
            continue
        kept: list[CellRange] = []
        count = 0
        for cells in notice.cells:
            rows = _runs(cells.min_row, cells.max_row, hidden_rows)
            cols = _runs(cells.min_col, cells.max_col, hidden_cols)
            for top, bottom in rows:
                for left, right in cols:
                    kept.append(CellRange(top, left, bottom, right))
                    count += (bottom - top + 1) * (right - left + 1)
        if not kept:
            continue
        if count != notice.count and notice.code not in LEGACY:
            message = describe(
                notice.code,
                notice.outcome,
                notice.reason,
                count,
                _detail(notice),
                notice.output,
                notice.dependency,
            )
        else:
            message = notice.message
        result.append(
            Notice(
                notice.code,
                notice.outcome,
                notice.reason,
                message,
                notice.sheet,
                tuple(kept),
                count,
                notice.output,
                notice.context,
                notice.dependency,
            )
        )
    return result


def _detail(notice: Notice) -> str:
    """The detail part of a generated message (after the first ``": "`` past the name)."""
    _, _, detail = notice.message.partition("): ")
    return detail


def _runs(first: int, last: int, hidden: set[int]) -> list[tuple[int, int]]:
    """The runs of indices in ``first..last`` that are not hidden."""
    if not hidden:
        return [(first, last)]
    runs: list[tuple[int, int]] = []
    start = None
    for index in range(first, last + 1):
        if index in hidden:
            if start is not None:
                runs.append((start, index - 1))
                start = None
        elif start is None:
            start = index
    if start is not None:
        runs.append((start, last))
    return runs


def summary(notices: Iterable[Notice], sheet: str, output: str) -> str:
    """One line for :class:`NotShownWarning`: what kinds, and where to look."""
    items = list(notices)
    kinds = ", ".join(
        f"{n.code} ({n.count} cell{'s' if n.count != 1 else ''})" if n.level == "cells" else n.code
        for n in items[:6]
    )
    more = f", and {len(items) - 6} more" if len(items) > 6 else ""
    return (
        f"sheet {sheet!r} to {output}: {len(items)} kind(s) of content not shown or approximated "
        f"({kinds}{more}); see sheetpeek.notices() or pass on_notice"
    )


def deliver(
    notices: tuple[Notice, ...],
    on_notice: Callable[[Notice], object] | Literal["ignore"] | None,
    sheet: str,
    output: str,
    stacklevel: int,
) -> None:
    """Hand notices to the caller's handler, or issue one :class:`NotShownWarning`."""
    if on_notice == "ignore" or not notices:
        return
    if on_notice is None:
        warnings.warn(NotShownWarning(summary(notices, sheet, output)), stacklevel=stacklevel + 1)
        return
    if not callable(on_notice):
        raise TypeError("on_notice must be a callable, 'ignore' or None")
    for notice in notices:
        on_notice(notice)
