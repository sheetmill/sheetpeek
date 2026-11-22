"""The machine's screen geometry: recorded once per machine, checked before every window capture.

Where a dash or a fill tile starts follows Excel's window on the screen (a dash's phase moved with
the window at zoom 200), so every screen capture of one machine is taken in one geometry: the main
screen's size in points and its backing scale, Excel's window maximized (its bounds in points and
window state), and the window's usable width and height (Excel's ``usable width`` and ``usable
height``: the sheet area, in points). The expected geometry is a small JSON file outside the
repository, one per machine:

    python -m capture.screen geometry show      # what Excel shows now, and the expected file
    python -m capture.screen geometry record    # write what Excel shows now as the expected

The file is ``$SHEETPEEK_CAPTURE_GEOMETRY``, else ``~/.sheetpeek-capture/geometry.json``. A
capture records the geometry it found in its header (``geometry``) and refuses to run when no
file is recorded or when what it finds differs from the file in any field.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

FIELDS = ("screen_pt", "display_scale", "window_pt", "usable_pt", "window_state")
"""What a geometry holds: sizes and bounds in points, rounded to hundredths."""


class GeometryError(RuntimeError):
    """The machine's geometry is not recorded, or differs from what was recorded."""


def expected_path() -> Path:
    """The machine's expected-geometry file."""
    found = os.environ.get("SHEETPEEK_CAPTURE_GEOMETRY")
    return Path(found) if found else Path.home() / ".sheetpeek-capture" / "geometry.json"


def normalized(geometry: dict[str, Any]) -> dict[str, Any]:
    """``geometry`` with exactly :data:`FIELDS`, numbers rounded to hundredths."""

    def value(v: Any) -> Any:
        if isinstance(v, (list, tuple)):
            return [value(x) for x in v]
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return round(float(v), 2)
        return v

    missing = [f for f in FIELDS if f not in geometry]
    if missing:
        raise GeometryError(f"geometry lacks {', '.join(missing)}")
    return {f: value(geometry[f]) for f in FIELDS}


def text(geometry: dict[str, Any]) -> str:
    """One line for a capture header: the normalized geometry as compact JSON, keys sorted."""
    return json.dumps(normalized(geometry), sort_keys=True, separators=(",", ":"))


def expected(path: Path | None = None) -> dict[str, Any]:
    """The recorded geometry; refuses when none is recorded."""
    path = path or expected_path()
    if not path.exists():
        raise GeometryError(
            f"no expected geometry at {path}: record it once on this machine with "
            "`python -m capture.screen geometry record`"
        )
    return normalized(json.loads(path.read_text(encoding="utf-8")))


def differences(found: dict[str, Any], wanted: dict[str, Any]) -> list[str]:
    """Each field in which ``found`` differs from ``wanted``, as ``field: found != wanted``."""
    a, b = normalized(found), normalized(wanted)
    return [f"{f}: {a[f]} != {b[f]}" for f in FIELDS if a[f] != b[f]]


def check(found: dict[str, Any], path: Path | None = None) -> None:
    """Refuse a capture whose geometry differs from the machine's recorded one."""
    wanted = expected(path)
    differ = differences(found, wanted)
    if differ:
        raise GeometryError(
            "Excel's window is not in this machine's recorded geometry ("
            + "; ".join(differ)
            + f"; file {path or expected_path()})"
        )


def record(found: dict[str, Any], path: Path | None = None) -> Path:
    """Write ``found`` as the machine's expected geometry."""
    path = path or expected_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(normalized(found), indent=1, sort_keys=True) + "\n", encoding="utf-8"
    )
    return path
