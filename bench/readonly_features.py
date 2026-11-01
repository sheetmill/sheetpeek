"""What openpyxl exposes in normal and read-only mode (ARCHITECTURE.md §5.1).

    uv run python bench/readonly_features.py [DIR]

Writes one small workbook into DIR (default: a temporary directory) and prints, for each load
mode, which worksheet attributes exist and what the cell styles resolve to.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

ATTRIBUTES = (
    "merged_cells",
    "column_dimensions",
    "row_dimensions",
    "sheet_format",
    "dimensions",
    "freeze_panes",
    "sheet_view",
    "conditional_formatting",
    "max_row",
    "max_column",
)


def build(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 1234.5
    ws["A1"].number_format = "#,##0.00"
    ws["A2"] = "merged"
    ws.merge_cells("A2:C3")
    ws["A2"].font = Font(name="Arial", size=14, bold=True, color="FF00FF00")
    ws["A2"].fill = PatternFill("solid", fgColor="FFFFFF00")
    ws["A2"].alignment = Alignment(horizontal="center", wrap_text=True)
    ws.column_dimensions["B"].width = 20
    ws.row_dimensions[3].height = 30
    wb.save(path)


def report(path: Path, **options: bool) -> None:
    wb = openpyxl.load_workbook(path, **options)
    ws = wb.active
    print(f"\n{type(ws).__name__} {options or '(normal mode)'}")
    for name in ATTRIBUTES:
        print(f"  {name:24s} {'yes' if hasattr(ws, name) else 'missing'}")
    for row in ws.iter_rows(min_row=1, max_row=2, max_col=1):
        cell = row[0]
        print(
            f"  {cell.coordinate}: format={cell.number_format!r} font={cell.font.name},"
            f"{cell.font.sz},bold={cell.font.b} fill={cell.fill.fill_type} "
            f"align={cell.alignment.horizontal},wrap={cell.alignment.wrap_text}"
        )
    wb.close()


def main() -> None:
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(tempfile.mkdtemp())
    path = folder / "readonly_features.xlsx"
    build(path)
    print(f"openpyxl {openpyxl.__version__}, {path}")
    report(path)
    report(path, read_only=True)


if __name__ == "__main__":
    main()
