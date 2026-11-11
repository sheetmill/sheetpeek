"""The outline pass: groups from outline levels, and the symbols each laid-out row and column shows.

The rule (``ARCHITECTURE.md`` §4.11): for each level ``L``, a group is a maximal run of rows whose
level is at least ``L``; its summary row is the row after the run, or the row before it when
summaries are above. Its button is a minus while any detail row is visible, and a plus when all
are hidden or the summary row is marked collapsed. Each visible detail row gets a dot in the lane
after its own level, unless a button is there. Columns follow the same rule, with summaries right
or left.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import TYPE_CHECKING, Literal

from sheetpeek._model import column_letter
from sheetpeek._notices import Collector
from sheetpeek._outline_model import OutlineMark, OutlineMarkKind, RenderedOutline

if TYPE_CHECKING:
    from sheetpeek._layout import Draft
    from sheetpeek._rendered import RenderedSheet

_LAST = {"rows": 1_048_576, "columns": 16_384}


def outline_pass(draft: Draft) -> None:
    """Lay out the outline symbols, when the sheet has an outline and the file shows them."""
    sheet = draft.sheet
    outline = sheet.outline
    if outline is None or not outline.show_symbols or sheet.ref is None:
        return
    kept_rows = {r.index for r in draft.kept_rows}
    kept_columns = {c.index for c in draft.kept_columns}
    from sheetpeek._layout import _column_hidden, _row_hidden

    rows = marks(
        [(r.index, r.outline_level, _row_hidden(r, sheet), r.collapsed) for r in sheet.rows],
        kept_rows,
        outline.summary_below,
        outline.rows_beyond,
        "rows",
        draft.notes,
    )
    columns = marks(
        [(c.index, c.outline_level, _column_hidden(c), c.collapsed) for c in sheet.columns],
        kept_columns,
        outline.summary_right,
        outline.columns_beyond,
        "columns",
        draft.notes,
    )
    draft.extra["outline"] = RenderedOutline(
        rows,
        columns,
        outline.row_depth + 1 if outline.row_depth else 0,
        outline.column_depth + 1 if outline.column_depth else 0,
        outline.summary_below,
        outline.summary_right,
    )


def groups(levels: Sequence[int], summary_after: bool = True) -> list[tuple[int, int, int, int]]:
    """(level, first, last, summary) of each group, as positions in ``levels``.

    A group of level ``L`` is a maximal run of positions whose level is at least ``L``; its
    summary is the position after the run, or before it when ``summary_after`` is false. The
    summary may fall outside ``levels``.
    """
    found = []
    for level in range(1, max(levels, default=0) + 1):
        i, n = 0, len(levels)
        while i < n:
            if levels[i] >= level:
                j = i
                while j + 1 < n and levels[j + 1] >= level:
                    j += 1
                found.append((level, i, j, j + 1 if summary_after else i - 1))
                i = j + 1
            else:
                i += 1
    return found


def marks(
    items: Sequence[tuple[int, int, bool, bool]],
    kept: set[int],
    summary_after: bool,
    beyond: tuple[int, int],
    axis: Literal["rows", "columns"],
    notes: Collector,
) -> tuple[OutlineMark, ...]:
    """The marks of one axis.

    Args:
        items: (index, level, hidden in the file, collapsed) of each row or column in the range.
        kept: The indices laid out; only they get marks.
        summary_after: Summaries below (rows) or right (columns) of their details.
        beyond: The levels of the row or column just before and just after the range.
        axis: ``"rows"`` or ``"columns"``, for the notices.
        notes: Where a group cut by the range is reported.
    """
    if not items:
        return ()
    n = len(items)
    # Positions 0 and n + 1 stand for the neighbors outside the range: a run that reaches them
    # continues past the edge.
    levels = [beyond[0], *(level for _, level, _, _ in items), beyond[1]]
    found: dict[tuple[int, int], OutlineMarkKind] = {}
    buttons: set[tuple[int, int]] = set()

    def item(position: int) -> tuple[int, int, bool, bool]:
        return items[position - 1]

    for level, first, last, summary in groups(levels, summary_after):
        inside = range(max(first, 1), min(last, n) + 1)
        drawn = [p for p in inside if item(p)[0] in kept]
        if 1 <= summary <= n:
            index, _, _, collapsed = item(summary)
            all_hidden = bool(inside) and all(item(p)[2] for p in inside)
            if index in kept:
                found[(index, level)] = "closed" if collapsed or all_hidden else "open"
                buttons.add((index, level))
        elif inside:
            _report_cut(notes, axis, item(inside[0])[0], item(inside[-1])[0], summary, items)
        # The tick is at the end away from the summary, when the run ends inside the range.
        if summary_after:
            tick = drawn[0] if drawn and first >= 1 else None
        else:
            tick = drawn[-1] if drawn and last <= n else None
        for p in drawn:
            found[(item(p)[0], level)] = "start" if p == tick else "bar"
    for index, level, _, _ in items:
        if level and index in kept and (index, level + 1) not in buttons:
            found[(index, level + 1)] = "dot"
    return tuple(OutlineMark(i, lane, kind) for (i, lane), kind in sorted(found.items()))


def _report_cut(
    notes: Collector,
    axis: Literal["rows", "columns"],
    first: int,
    last: int,
    summary: int,
    items: Sequence[tuple[int, int, bool, bool]],
) -> None:
    """A group whose summary is outside the range: its bars are drawn, its button is not."""
    context: dict[str, str | int | float] = {"axis": axis, "first": first, "last": last}
    if summary in (0, len(items) + 1):  # the neighbor just outside; else further, unknown
        index = items[0][0] + summary - 1
        if not 1 <= index <= _LAST[axis]:
            return  # past the sheet's edge: there is no summary to cut off
        context["summary"] = index
    name = str if axis == "rows" else column_letter
    noun = "row" if axis == "rows" else "column"
    notes.add(
        "outline.cut",
        "omitted",
        "option",
        detail=f"the button of the group over {axis} {name(first)} to {name(last)}, whose "
        f"summary {noun} is outside the range",
        context=context,
    )


def presence(rendered: RenderedSheet) -> int:
    """How many outline symbols the laid-out sheet has, for the census."""
    outline = rendered.outline
    return len(outline.rows) + len(outline.columns) if outline is not None else 0
