"""Run a screen capture: build a workbook, open it in Excel, capture its window, map cells.

    caps = capture(build, cells=["B4", "C4"], out=Path(...), markers=("B2", "J20"), zooms=(100,))

``build(ws)`` fills the first worksheet of a new openpyxl workbook (``ws.parent`` is the
workbook). ``capture`` adds the two marker cells (magenta fills, in different rows and columns;
nothing else may be magenta), saves the workbook inside Excel's container, opens it in one
session, and for each zoom parks the selection off-view, scrolls to A1, captures the window,
converts the image to sRGB, and maps cells to pixels from the markers. What Excel displays for
each cell is read once (it does not depend on the zoom). The default markers, B2 and J20, keep
off row 1 and column A: a marker touching the headings starts a pixel late (scale 1.995, not
2.000, measured).

``out`` receives ``formats.tsv`` (a header of ``environment``'s fields, then each cell's
properties) and ``window_<zoom>.png``; ``load(out)`` measures a run again without Excel. Every
measuring helper reports points first (sheet points at the capture's zoom: image pixels divided
by ``scale``), with image pixels beside them.
"""

from __future__ import annotations

import contextlib
import os
import subprocess
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from capture.screen import png
from capture.screen.session import CONTAINER, ExcelSession, environment
from capture.screen.window import (
    GRIDLINE,
    MARKER,
    Grid,
    capture_window,
    locate,
    nominal_scale,
    platform_of,
)

SRGB = Path("/System/Library/ColorSync/Profiles/sRGB Profile.icc")
PARK = "AZ300"


def to_srgb(source: Path, target: Path) -> None:
    """Convert a capture from the display's profile to sRGB (``sips --matchTo``)."""
    result = subprocess.run(
        ["sips", "--matchTo", str(SRGB), str(source), "--out", str(target)],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode or not target.exists():
        raise RuntimeError(f"sips failed: {result.stderr.strip()[:200]}")


@dataclass
class Capture:
    zoom: int
    image: png.Image
    grid: Grid
    check: dict
    geometry: dict[str, dict[str, float]]
    rows: list[tuple[str, str, str]] = field(repr=False)
    platform: str
    """``mac`` or ``windows``: every capture states the platform it was taken on."""

    @property
    def gridline(self) -> tuple[int, int, int] | None:
        """The default gridline's gray on this run's platform, in sRGB; ``None`` until recorded."""
        return GRIDLINE[self.platform]

    @property
    def scale(self) -> float:
        """Image pixels per sheet point at this zoom (2 at zoom 100 on a 2x display)."""
        return self.grid.scale

    def edge(self, cell: str, which: str) -> float:
        """A cell's edge or center in image pixels: left, right, top, bottom, cx, cy."""
        if cell not in self.geometry:
            raise KeyError(f"{cell}: its geometry was not read; add it to capture(cells=...)")
        c = self.geometry[cell]
        return {
            "left": self.grid.x(c["left"]),
            "right": self.grid.x(c["left"] + c["width"]),
            "top": self.grid.y(c["top"]),
            "bottom": self.grid.y(c["top"] + c["height"]),
            "cx": self.grid.x(c["left"] + c["width"] / 2),
            "cy": self.grid.y(c["top"] + c["height"] / 2),
        }[which]

    def pixel(self, x: float, y: float) -> tuple[int, int, int]:
        return self.image.pixel(round(x), round(y))

    def pt(self, pixels: float) -> str:
        """``"1.5 pt (3 px)"``: points first, image pixels in brackets."""
        return f"{pixels / self.scale:g} pt ({pixels:g} px)"

    def value(self, cell: str, key: str) -> str | None:
        for c, k, v in self.rows:
            if c == cell and k == key:
                return v
        return None


def _geometry(rows: list[tuple[str, str, str]]) -> dict[str, dict[str, float]]:
    g: dict[str, dict[str, float]] = {}
    for cell, key, value in rows:
        if key.startswith("geometry.") and key != "geometry.text":
            with contextlib.suppress(ValueError):  # ERR: Excel refused that read
                g.setdefault(cell, {})[key.split(".", 1)[1]] = float(value)
    return g


def capture(
    build: Callable,
    cells: list[str],
    out: Path,
    *,
    edition: str | None = None,
    markers: tuple[str, str] = ("B2", "J20"),
    zooms: tuple[int, ...] = (100,),
    folder: str = "sheetpeek-capture",
    name: str = "capture",
    postprocess: Callable[[Path], None] | None = None,
    top_row: int | None = None,
    state: bool = False,
) -> list[Capture]:
    """Run one capture session; write ``formats.tsv`` and ``window_<zoom>.png`` (sRGB) to ``out``.

    ``postprocess(path)``, if given, edits the saved workbook before Excel opens it: for XML that
    openpyxl cannot write (``x14`` extensions, threaded comments). Use ``patch_part`` for that.
    ``top_row``, if given, scrolls each view so that row is the window's top row, and keeps the
    window's visible range as read ``("window", "visible.<zoom>", address)``. ``state`` keeps
    the window's state after each view as read ``("window", "state.<zoom>", state)``
    (``ExcelSession.window``: the zoom, the view, the panes).
    """
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill

    edition = edition or os.environ.get("EXCEL_EDITION")
    if not edition:
        raise SystemExit("state the Excel edition (Excel > About Excel) in EXCEL_EDITION")
    out.mkdir(parents=True, exist_ok=True)
    (CONTAINER / folder).mkdir(parents=True, exist_ok=True)
    path = CONTAINER / folder / f"{name}_{os.getpid()}.xlsx"
    header = environment(edition)
    cells = list(dict.fromkeys([*markers, *cells]))
    session = ExcelSession(path, evidence=out / "excel_timeouts")
    shots = []
    try:
        wb = Workbook()
        build(wb.active)
        for cell in markers:
            wb.active[cell].fill = PatternFill("solid", fgColor="FF" + MARKER)
        wb.save(path)
        if postprocess is not None:
            postprocess(path)
        with session:
            session.view(PARK, zooms[0])
            rows = session.read_cells(cells)
            for zoom in zooms:
                session.view(PARK, zoom)
                if state:
                    rows.append(("window", f"state.{zoom}", session.window()))
                if top_row is not None:
                    rows.append(("window", f"visible.{zoom}", session.scroll(top_row)))
                raw = out / f"window_{zoom}_display.png"
                window, shot = capture_window(path.stem, raw)
                shots.append((zoom, window, raw))
                header.setdefault("display_scale", f"{shot.width / window['w']:.3f}")
    finally:
        path.unlink(missing_ok=True)
    header["zooms"] = ",".join(str(z) for z in zooms)
    header["colors"] = "captures converted to sRGB (sips --matchTo sRGB Profile.icc)"
    header["suite"] = name
    header["markers"] = ",".join(markers)
    header["workbooks_open_after"] = session.open_after or "unknown"
    header["platform"] = "mac"
    write_tsv(out / "formats.tsv", header, rows)
    geometry = _geometry(rows)
    results = []
    for zoom, window, raw in shots:
        srgb = out / f"window_{zoom}.png"
        to_srgb(raw, srgb)
        raw.unlink(missing_ok=True)
        image = png.read(srgb)
        display = image.width / window["w"]
        nominal = nominal_scale(zoom, round(display, 3), "mac")
        grid, check = locate(image, geometry[markers[0]], geometry[markers[1]], nominal)
        check["display_scale"] = round(display, 3)
        results.append(Capture(zoom, image, grid, check, geometry, rows, "mac"))
    return results


def load(
    out: Path,
    markers: tuple[str, str] | None = None,
    *,
    display_scale: float | None = None,
    one_marker: bool = False,
) -> list[Capture]:
    """Reload a saved run (``formats.tsv`` and ``window_<zoom>.png``) without Excel.

    The markers are the ones the run's header records, unless given. The nominal scale needs the
    display's scale: the header's ``display_scale``, else ``display_scale`` (runs before the
    field); without either, both markers must show. ``one_marker`` measures from the first marker
    and the nominal scale even where both show (:func:`capture.screen.window.locate`).
    """
    found = header(out)
    if markers is None:
        first, _, second = found.get("markers", "B2,J20").partition(",")
        markers = (first, second)
    platform = platform_of(found)
    display = float(found["display_scale"]) if "display_scale" in found else display_scale
    rows = []
    for line in (out / "formats.tsv").read_text(encoding="utf-8").splitlines(keepends=True):
        if line.startswith("# ") or line.startswith("cell\t"):
            continue
        rows.append(tuple([*line.rstrip("\n").split("\t"), "", ""][:3]))
    geometry = _geometry(rows)
    results = []
    for image_path in sorted(out.glob("window_*.png")):
        parts = image_path.stem.split("_")
        if len(parts) != 2 or not parts[1].isdigit():
            # window_<zoom>_display.png is the unconverted display-profile shot a run leaves when it
            # fails before converting (sips --matchTo sRGB); it must not be measured as a capture
            continue
        stem = parts[1]
        image = png.read(image_path)
        nominal = nominal_scale(int(stem), display, platform) if display else None
        grid, check = locate(
            image, geometry[markers[0]], geometry[markers[1]], nominal, one_marker=one_marker
        )
        results.append(Capture(int(stem), image, grid, check, geometry, rows, platform))
    return sorted(results, key=lambda c: c.zoom)


def patch_part(path: Path, edits: dict[str, Callable[[str], str] | str]) -> None:
    """Rewrite parts of a saved xlsx: ``{part name: new text, or a function of the old text}``.

    A part not present is added. The ``[Content_Types].xml`` and relationship parts are the
    caller's to edit like any other part. A function sees its part in the standard library's form
    (:func:`capture.xmlform.standard_form`), whichever writer openpyxl used.
    """
    import shutil
    import zipfile

    from capture.xmlform import serialized, standard_form

    tmp = path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        names = src.namelist()
        for name in names:
            data = src.read(name)
            if name in edits:
                edit = edits[name]
                text = data.decode("utf-8")
                if callable(edit) and serialized(name):
                    text = standard_form(text)
                data = (edit(text) if callable(edit) else edit).encode("utf-8")
            dst.writestr(name, data)
        for name, edit in edits.items():
            if name not in names:
                dst.writestr(name, (edit("") if callable(edit) else edit).encode("utf-8"))
    shutil.move(tmp, path)


# -- measuring helpers ---------------------------------------------------------------------------


def luma(rgb: tuple[int, int, int]) -> float:
    r, g, b = rgb
    return 0.299 * r + 0.587 * g + 0.114 * b


def off_white(rgb: tuple[int, int, int], tolerance: int = 12) -> bool:
    return any(255 - c > tolerance for c in rgb)


def runs(values: list[bool]) -> list[tuple[bool, int, int]]:
    """Run-length encode booleans: [(value, start, length)]."""
    out: list[list] = []
    for i, v in enumerate(values):
        if out and out[-1][0] == v:
            out[-1][2] += 1
        else:
            out.append([v, i, 1])
    return [tuple(r) for r in out]


def scan_column(cap: Capture, x: float, y0: float, y1: float) -> list[tuple[int, int, int]]:
    return [cap.pixel(x, y) for y in range(round(y0), round(y1))]


def scan_row(cap: Capture, y: float, x0: float, x1: float) -> list[tuple[int, int, int]]:
    return [cap.pixel(x, y) for x in range(round(x0), round(x1))]


def box(
    cap: Capture,
    x0: float,
    y0: float,
    x1: float,
    y1: float,
    test: Callable[[tuple[int, int, int]], bool],
) -> tuple[int, int, int, int] | None:
    """Bounding box (inclusive, image pixels) of pixels passing ``test`` inside a region."""
    pts = [
        (x, y)
        for y in range(round(y0), round(y1))
        for x in range(round(x0), round(x1))
        if test(cap.image.pixel(x, y))
    ]
    if not pts:
        return None
    return (
        min(p[0] for p in pts),
        min(p[1] for p in pts),
        max(p[0] for p in pts),
        max(p[1] for p in pts),
    )


def write_tsv(path: Path, header: dict[str, str], rows: list[tuple[str, str, str]]) -> None:
    """A capture's cell reads: ``# key: value`` header lines, then cell, property, value."""
    with open(path, "w", encoding="utf-8") as out:
        for key, value in header.items():
            out.write(f"# {key}: {value}\n")
        out.write("cell\tproperty\tvalue\n")
        for cell, key, value in rows:
            out.write(f"{cell}\t{key}\t{value.replace(chr(9), ' ').replace(chr(10), ' ')}\n")


def header(out: Path) -> dict[str, str]:
    """A saved capture's header (``formats.tsv``'s ``# key: value`` lines)."""
    found = {}
    for line in (out / "formats.tsv").read_text(encoding="utf-8").splitlines(keepends=True):
        if not line.startswith("# "):
            break
        key, _, value = line[2:].rstrip("\n").partition(": ")
        found[key] = value
    return found


def read_only(
    build: Callable[[Any], None],
    cells: list[str],
    out: Path,
    *,
    edition: str | None = None,
    folder: str = "sheetpeek-capture",
    name: str = "capture",
    postprocess: Callable[[Path], None] | None = None,
) -> list[tuple[str, str, str]]:
    """Open a built workbook and read what Excel shows for ``cells``, without a capture.

    For what Excel's own reads settle (row heights, the text a cell shows): no window has to
    hold the cells. Writes ``formats.tsv`` to ``out`` as ``capture`` does; ``postprocess`` edits
    the saved workbook before Excel opens it, as in ``capture``.
    """
    from openpyxl import Workbook

    edition = edition or os.environ.get("EXCEL_EDITION")
    if not edition:
        raise SystemExit("state the Excel edition (Excel > About Excel) in EXCEL_EDITION")
    out.mkdir(parents=True, exist_ok=True)
    (CONTAINER / folder).mkdir(parents=True, exist_ok=True)
    path = CONTAINER / folder / f"{name}_{os.getpid()}.xlsx"
    found = environment(edition)
    session = ExcelSession(path, evidence=out / "excel_timeouts")
    try:
        wb = Workbook()
        build(wb.active)
        wb.save(path)
        if postprocess is not None:
            postprocess(path)
        with session:
            rows = session.read_cells(cells)
    finally:
        path.unlink(missing_ok=True)
    found["suite"] = name
    found["workbooks_open_after"] = session.open_after or "unknown"
    write_tsv(out / "formats.tsv", found, rows)
    return rows
