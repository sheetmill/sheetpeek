"""Snapshot what sheetpeek's reader gives for a folder of workbooks, and diff two snapshots.

Used to change the reader without changing what it reads: snapshot a corpus with the version
before the change and with the one after, and every difference is listed for a reason to be
given. One JSON line per workbook and per sheet: the sheet list, the default sheet and the
workbook's notices, and for each sheet digests of its values, styles, runs, columns, rows,
merges, links and outline, the rows the file lists, its sheet format, view, print layout, its
notices with their cells, and its filters when it has any.

    find CORPUS -name '*.xlsx' > list.txt
    python -m tests.watchdog -- python tools/reader_snapshot.py sweep OLD/src list.txt a.jsonl
    python -m tests.watchdog -- python tools/reader_snapshot.py sweep src list.txt b.jsonl
    python tools/reader_snapshot.py diff a.jsonl b.jsonl

``sweep SRC LIST OUT [--workers N]`` runs ``snap`` in N child processes (default 4) with
``PYTHONPATH=SRC``, so each version is read from its own source tree, and joins their output.
Run it under ``tests/watchdog.py``: the watchdog bounds the whole process group, workers
included, and each read is bounded by ``max_cells`` (2,000,000) besides. ``snap LIST OUT
[START STOP]`` snapshots lines START to STOP of LIST; ``diff A B [EXAMPLES]`` prints how many
records differ in each field, with a few examples of each.

A snapshot depends on this tool's records as well as on the reader, so two snapshots compare
only when taken with the same version of the tool. A sweep of about 16,000 workbooks takes
about 10 minutes with 4 workers.

``sample LIST OUT SNAPSHOT... [--with FIELD ...] [--random N]`` writes a smaller list, a quick
tier, in LIST's order: every file whose records differ between any two consecutive SNAPSHOTs
given (oldest first), every file with a sheet whose FIELD (``runs``, ``outline``, ``links``,
...) is not that field's most common value in the first snapshot (the files carrying a
feature), and N more files drawn from the rest of LIST (default 300, with a fixed seed, so the
tier is the same each time):

    python tools/reader_snapshot.py sample list.txt quick.txt a.jsonl b.jsonl ... [--with FIELD]
    python -m tests.watchdog -- python tools/reader_snapshot.py sweep src quick.txt q.jsonl
"""

from __future__ import annotations

import collections
import hashlib
import itertools
import json
import os
import random
import subprocess
import sys
import warnings
from collections.abc import Iterable
from pathlib import Path
from typing import Any

MAX_CELLS = 2_000_000


def _digest(items: Iterable[Any]) -> str:
    found = hashlib.sha1(usedforsecurity=False)
    for item in items:
        found.update(repr(item).encode("utf-8", "backslashreplace"))
        found.update(b"\0")
    return found.hexdigest()[:12]


def _sheet(book: Any, name: str) -> dict[str, Any]:
    sheet = book.sheet(name)
    cells = [cell for row in sheet.cells for cell in row]
    return {
        "ref": str(sheet.ref),
        "used": str(sheet.used_range),
        "values": _digest((c.row, c.column, c.kind, repr(c.value)) for c in cells),
        "styles": _digest(c.style for c in cells),
        "runs": _digest((c.row, c.column, c.runs) for c in cells if c.runs),
        "cols": _digest((c.index, c.width, c.hidden) for c in sheet.columns),
        "rows": _digest((r.index, r.height, r.hidden) for r in sheet.rows),
        "listed": _digest(r.index for r in sheet.rows if r.listed),
        "format": repr(
            (
                sheet.state,
                sheet.date1904,
                sheet.default_column_width,
                sheet.base_column_width,
                sheet.default_row_height,
                sheet.has_sheet_format,
                sheet.custom_row_height,
                sheet.zero_row_height,
                sheet.thick_top,
                sheet.thick_bottom,
                sheet.show_headings,
            )
        ),
        "merges": _digest(map(str, sheet.merges)),
        "view": repr(sheet.view),
        "print": repr(
            (sheet.print_areas, sheet.print_area_problem, sheet.row_breaks, sheet.column_breaks)
        ),
        "links": _digest((c.row, c.column, c.hyperlink) for c in cells if c.hyperlink),
        "notices": sorted(f"{n.code}:{n.count}" for n in sheet.notices),
        "notice_cells": _digest(
            sorted((n.code, tuple(map(str, n.cells))) for n in sheet.notices if n.cells)
        ),
        "normal": repr(sheet.normal_font),
        "grid": repr((sheet.show_gridlines, sheet.grid_color)),
        "outline": _digest(
            [
                *(
                    (r.index, r.outline_level, r.collapsed)
                    for r in sheet.rows
                    if r.outline_level or r.collapsed
                ),
                *(
                    (c.index, c.outline_level, c.collapsed)
                    for c in sheet.columns
                    if c.outline_level or c.collapsed
                ),
                repr(sheet.outline),
            ]
        ),
    } | _filters(sheet)


def _filters(sheet: Any) -> dict[str, str]:
    """The sheet's filters, when it has any (a reader without them records nothing)."""
    found = getattr(sheet, "autofilters", ())
    return {"autofilters": repr(found)} if found else {}


def _default(book: Any) -> str | None:
    """The sheet ``book.sheet()`` reads, without reading it: a private name, held by both sides."""
    try:
        return str(book._default_sheet_name())
    except Exception as error:  # a workbook without one is a result to compare
        return f"{type(error).__name__}: {str(error)[:120]}"


def snap(listing: str, out: str, start: int = 0, stop: int | None = None) -> None:
    """Append the records of lines ``start`` to ``stop`` of ``listing`` to ``out``."""
    import sheetpeek

    limits = sheetpeek.ReadLimits(max_cells=MAX_CELLS)
    paths = [p for p in Path(listing).read_text().split("\n") if p][start:stop]
    with open(out, "a", encoding="utf-8") as handle:
        for path in paths:
            warnings.simplefilter("ignore")
            try:
                with sheetpeek.open(path, limits=limits) as book:
                    infos = [(i.name, i.kind, i.state, i.tab_color) for i in book.sheets]
                    head = {
                        "file": path,
                        "sheets": infos,
                        "default": _default(book),
                        "notices": sorted(f"{n.code}:{n.count}" for n in book.notices),
                    }
                    handle.write(json.dumps(head) + "\n")
                    for name, *_ in infos:
                        try:
                            record = _sheet(book, name)
                        except Exception as error:  # every refusal is a result to compare
                            record = {"error": f"{type(error).__name__}: {str(error)[:120]}"}
                        handle.write(json.dumps({"file": path, "sheet": name, **record}) + "\n")
            except Exception as error:
                failure = f"{type(error).__name__}: {str(error)[:120]}"
                handle.write(json.dumps({"file": path, "error": failure}) + "\n")
            handle.flush()


def sweep(source: str, listing: str, out: str, workers: int = 4) -> None:
    """Snapshot ``listing`` with the sheetpeek under ``source``, in ``workers`` processes."""
    count = sum(1 for line in Path(listing).read_text().split("\n") if line)
    step = (count + workers - 1) // workers
    env = {**os.environ, "PYTHONPATH": os.path.abspath(source)}
    parts = [f"{out}.{k}" for k in range(workers)]
    for part in parts:
        Path(part).unlink(missing_ok=True)
    children = [
        subprocess.Popen(  # this script, run with this interpreter
            [sys.executable, __file__, "snap", listing, part, str(k * step), str((k + 1) * step)],
            env=env,
        )
        for k, part in enumerate(parts)
    ]
    codes = [child.wait() for child in children]
    with open(out, "w", encoding="utf-8") as joined:
        for part in parts:
            joined.write(Path(part).read_text(encoding="utf-8"))
            Path(part).unlink()
    if any(codes):
        raise SystemExit(f"a worker failed: return codes {codes}")


def _load(path: str) -> dict[tuple[str, str | None, bool], dict[str, Any]]:
    found = {}
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        record = json.loads(line)
        found[(record["file"], record.get("sheet"), "sheets" in record)] = record
    return found


def diff(before: str, after: str, examples: int = 3) -> None:
    """Print how many records differ in each field, with ``examples`` of each."""
    a, b = _load(before), _load(after)
    only_a, only_b = a.keys() - b.keys(), b.keys() - a.keys()
    print(f"records: {len(a)} and {len(b)}; only before {len(only_a)}, only after {len(only_b)}")
    fields: collections.Counter[str] = collections.Counter()
    shown: dict[str, list[tuple[str, str | None, Any, Any]]] = collections.defaultdict(list)
    for key in sorted(a.keys() & b.keys(), key=str):
        for field in sorted(set(a[key]) | set(b[key])):
            if a[key].get(field) != b[key].get(field):
                fields[field] += 1
                if len(shown[field]) < examples:
                    name = key[0].rsplit("/", 1)[-1]
                    shown[field].append((name, key[1], a[key].get(field), b[key].get(field)))
    for field, n in fields.most_common():
        print(f"\n{field}: {n}")
        for name, sheet, x, y in shown[field]:
            print(f"  {name} [{sheet}]\n    before {str(x)[:160]}\n    after  {str(y)[:160]}")
    for key in sorted(only_a, key=str)[:examples] + sorted(only_b, key=str)[:examples]:
        print("only one side:", key[0].rsplit("/", 1)[-1], key[1])


def sample(
    listing: str, out: str, snapshots: list[str], fields: list[str], extra: int = 300
) -> None:
    """Write the quick tier's list: see the module's docstring."""
    paths = [p for p in Path(listing).read_text().split("\n") if p]
    loaded = [_load(path) for path in snapshots]
    chosen: set[str] = set()
    for a, b in itertools.pairwise(loaded):
        for key in a.keys() | b.keys():
            before, after = a.get(key, {}), b.get(key, {})
            shared = before.keys() & after.keys() if before and after else {"file"}
            # Fields one snapshot lacks (an older format) are not differences.
            if not (before and after) or any(before[f] != after[f] for f in shared):
                chosen.add(key[0])
    for field in fields:
        values = collections.Counter(r.get(field) for r in loaded[0].values() if "sheet" in r)
        common = values.most_common(1)[0][0] if values else None
        chosen.update(
            key[0] for key, r in loaded[0].items() if "sheet" in r and r.get(field) != common
        )
    rest = [p for p in paths if p not in chosen]
    chosen.update(random.Random(0).sample(rest, min(extra, len(rest))))
    kept = [p for p in paths if p in chosen]
    Path(out).write_text("".join(f"{p}\n" for p in kept), encoding="utf-8")
    print(f"{len(kept)} of {len(paths)} files")


def main(argv: list[str]) -> None:
    """The command line: ``snap``, ``sweep``, ``diff`` or ``sample``."""
    if len(argv) >= 3 and argv[0] == "snap":
        bounds = [int(v) for v in argv[3:5]]
        snap(argv[1], argv[2], *bounds)
    elif len(argv) >= 4 and argv[0] == "sweep":
        workers = int(argv[5]) if len(argv) > 5 and argv[4] == "--workers" else 4
        sweep(argv[1], argv[2], argv[3], workers)
    elif len(argv) >= 3 and argv[0] == "diff":
        diff(argv[1], argv[2], int(argv[3]) if len(argv) > 3 else 3)
    elif len(argv) >= 4 and argv[0] == "sample":
        snapshots, fields, extra, rest = [], [], 300, argv[3:]
        while rest:
            if rest[0] == "--with":
                fields.append(rest[1])
                rest = rest[2:]
            elif rest[0] == "--random":
                extra = int(rest[1])
                rest = rest[2:]
            else:
                snapshots.append(rest.pop(0))
        sample(argv[1], argv[2], snapshots, fields, extra)
    else:
        raise SystemExit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
