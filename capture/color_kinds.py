"""Capture which kind Excel for Mac takes from a cell color that states two (read from Excel).

    EXCEL_EDITION="Microsoft 365" uv run --group capture python -m capture.color_kinds run
    uv run --group capture python -m capture.color_kinds read        # the saved answer again
    uv run --group capture python -m capture.color_kinds build PATH  # the workbook only

ARCHITECTURE.md §3.3: where a color element states more than one of ``indexed``, ``theme``,
``auto`` and ``rgb``, sheetpeek takes them in that order for cell colors; for tab colors Excel
draws theme before rgb before indexed, and rgb over auto (``tab_strip``). This suite asks the same
of a fill's foreground, a font color and a border color.

One sheet, ``Kinds``. Row 2 holds a control of each single kind (theme 5, rgb 0070C0, indexed
17), row 4 the pairs: theme with rgb (C), rgb with indexed (D), theme with indexed (E), auto with
rgb (F). Each cell has a solid fill, a font color and a thin bottom border in the same color
element, written over openpyxl's XML. The workbook part carries Excel's ``fileVersion``.

``run`` opens the workbook alone (Excel may refuse it: the refusal is the answer), reads each
cell's display fill and font color and its bottom border's color (``ExcelSession.read_cells``),
and saves the answer to ``capture/build/color_kinds/answer.tsv``; ``read`` writes
``tests/captures/color_kinds.tsv``. No window, no clipboard.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Border, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "capture" / "build" / "color_kinds"
OUTPUT = ROOT / "tests" / "captures" / "color_kinds.tsv"
SHEET = "Kinds"

THEME, RGB, INDEXED = 'theme="5"', 'rgb="FF0070C0"', 'indexed="17"'
CASES: dict[str, str] = {
    "C2": THEME,
    "D2": RGB,
    "E2": INDEXED,
    "C4": f"{THEME} {RGB}",
    "D4": f"{RGB} {INDEXED}",
    "E4": f"{THEME} {INDEXED}",
    "F4": f'auto="1" {RGB}',
}
"""Cell → the attributes of its color element (fill, font and bottom border alike)."""
COLUMNS = ("cell", "attributes", "fill", "font_color", "border_color")


def _marker(k: int) -> str:
    """A placeholder ARGB openpyxl writes for case ``k``, replaced after saving."""
    return f"FF0A0A{k:02X}"


def build(path: Path) -> list[str]:
    """Write the workbook to ``path``; the cells written."""
    from capture.screen.run import patch_part
    from capture.spill_neighbors import excel_saved

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    for k, cell in enumerate(CASES, 1):
        color = _marker(k)
        ws[cell] = "HH"
        ws[cell].fill = PatternFill("solid", fgColor=color)
        ws[cell].font = Font(color=color)
        ws[cell].border = Border(bottom=Side("thin", color=color))
    wb.save(path)

    def styles(text: str) -> str:
        for k, attributes in enumerate(CASES.values(), 1):
            marker = f'rgb="{_marker(k)}"'
            if marker not in text:
                raise ValueError(f"{marker} not written")
            text = text.replace(marker, attributes)
        return text

    patch_part(path, {"xl/styles.xml": styles, "xl/workbook.xml": excel_saved})
    return list(CASES)


def _rgb(value: str) -> str:
    """Excel's ``r, g, b`` answer as ``RRGGBB``, else the answer as given."""
    parts = [p.strip() for p in value.split(",")]
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        return "".join(f"{int(p):02X}" for p in parts)
    return value


def measure(answers: list[tuple[str, str, str]]) -> list[dict[str, str]]:
    """One record per case from ``read_cells``' (cell, property, value) answers."""
    found: dict[str, dict[str, str]] = {c: {"cell": c, "attributes": a} for c, a in CASES.items()}
    wanted = {
        "display.fill": "fill",
        "display.font_color": "font_color",
        "own.border.bottom.color": "border_color",
    }
    for cell, prop, value in answers:
        if cell in found and prop in wanted:
            found[cell][wanted[prop]] = _rgb(value)
    return list(found.values())


def run() -> None:
    """Open the workbook alone in Excel and save its answer (needs Excel)."""
    from capture.screen.session import CONTAINER, ExcelError, ExcelSession, environment

    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    path = CONTAINER / "color_kinds.xlsx"
    build(path)
    header = environment(os.environ.get("EXCEL_EDITION", "unknown"))
    lines = [f"# {k}: {v}" for k, v in header.items()]
    try:
        with ExcelSession(path, BUILD_DIR) as session:
            answers = session.read_cells(list(CASES))
    except ExcelError as exc:
        (BUILD_DIR / "answer.tsv").write_text("\n".join([*lines, f"refused\t{exc}"]) + "\n")
        print(f"refused: {exc}", file=sys.stderr)
        return
    body = ["\t".join(a) for a in answers]
    (BUILD_DIR / "answer.tsv").write_text("\n".join([*lines, *body]) + "\n", encoding="utf-8")
    write()


def write() -> None:
    """Measure the saved answer again and write ``color_kinds.tsv``."""
    text = (BUILD_DIR / "answer.tsv").read_text(encoding="utf-8").splitlines()
    head = [line for line in text if line.startswith("# ")]
    rows = [tuple(line.split("\t")) for line in text if not line.startswith("# ")]
    out = [*head, "# suite: color_kinds", "\t".join(COLUMNS)]
    if rows and rows[0][0] == "refused":
        out.append("refused\t" + rows[0][1])
    else:
        records = measure([r for r in rows if len(r) == 3])  # type: ignore[misc]
        out += ["\t".join(r.get(c, "") for c in COLUMNS) for r in records]
    OUTPUT.write_text("\n".join(out) + "\n", encoding="utf-8")
    print(f"{len(out) - len(head) - 2} rows written to {OUTPUT}")


def main(argv: list[str]) -> int:
    if len(argv) == 2 and argv[0] == "build":
        print(f"{len(build(Path(argv[1])))} cells in {argv[1]}")
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
