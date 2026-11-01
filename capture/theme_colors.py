"""Capture the colors Excel for Mac resolves for theme colors with tints (ARCHITECTURE.md §3.3).

    uv run --group capture python -m capture.theme_colors run

The workbook uses openpyxl's default theme. Row ``i + 1`` holds theme index ``i`` (0 to 11), and
each column a tint; every cell has that theme color and tint as its solid fill and as its font
color, and an ``x``. ``theme_colors.applescript`` reads the fill and font color Excel resolves
for each cell. The output, ``tests/captures/theme_colors.tsv``, has one record per cell.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "tests" / "captures" / "theme_colors.tsv"
CONTAINER_DIR = (
    Path.home() / "Library" / "Containers" / "com.microsoft.Excel" / "Data" / "Documents"
    / "sheetpeek"
)  # fmt: skip
WORKBOOK = CONTAINER_DIR / "theme_colors.xlsx"
THEMES = range(12)
TINTS = (-0.5, -0.25, -0.1, 0.0, 0.2, 0.4, 0.6, 0.8)


def build(path: Path = WORKBOOK) -> Path:
    """Write the workbook: theme index by row, tint by column."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "theme"
    for theme in THEMES:
        for k, tint in enumerate(TINTS):
            cell = ws.cell(theme + 1, k + 1, "x")
            cell.fill = PatternFill("solid", fgColor=Color(theme=theme, tint=tint))
            cell.font = Font(color=Color(theme=theme, tint=tint))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def cases() -> list[tuple[str, int, float]]:
    """(address, theme index, tint) in the order Excel reads the cells: row by row."""
    return [
        (f"{get_column_letter(k + 1)}{theme + 1}", theme, tint)
        for theme in THEMES
        for k, tint in enumerate(TINTS)
    ]


def run() -> None:
    """Build, have Excel read the colors, write the TSV."""
    build()
    script = Path(__file__).with_name("theme_colors.applescript")
    result = subprocess.run(
        ["osascript", str(script), str(WORKBOOK), *(address for address, _, _ in cases())],
        capture_output=True, text=True, check=False,
    )  # fmt: skip
    if result.returncode != 0:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    lines = result.stdout.strip().split("\n")
    expected = cases()
    if len(lines) != len(expected):
        sys.exit(f"expected {len(expected)} cells, got {len(lines)}: {result.stdout!r}")
    version = subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to get version'],
        capture_output=True, text=True, check=True,
    ).stdout.strip()  # fmt: skip
    with OUTPUT.open("w", encoding="utf-8", newline="\n") as out:
        out.write(f"# suite: theme_colors (capture.theme_colors); Excel {version} for Mac, en_US\n")
        out.write(
            "# openpyxl's default theme; each cell's solid fill and font color are the theme\n"
        )
        out.write("# index and tint given; fill and font: the RRGGBB Excel resolves (color of\n")
        out.write("# interior object, color of font object)\n")
        out.write(f"# captured: {dt.date.today().isoformat()}; method: AppleScript, per cell\n")
        out.write("cell\ttheme\ttint\tfill\tfont\n")
        for (address, theme, tint), line in zip(expected, lines, strict=True):
            got, fill, font = line.split("\t")
            if got != address:
                sys.exit(f"Excel answered for {got}, expected {address}")
            out.write(f"{address}\t{theme}\t{tint!r}\t{_hex(fill)}\t{_hex(font)}\n")
    print(f"wrote {OUTPUT}")


def _hex(rgb: str) -> str:
    return "".join(f"{int(v):02X}" for v in rgb.split(","))


if __name__ == "__main__":
    {"build": build, "run": run}[sys.argv[1]]()
