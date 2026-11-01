# Usage

## Reading a sheet

{func}`sheetpeek.read` returns a {class}`sheetpeek.Sheet`: what the file says about one worksheet,
before any layout. The examples on this page use a small workbook written with openpyxl.

```python
>>> import tempfile, pathlib, openpyxl, sheetpeek
>>> path = pathlib.Path(tempfile.mkdtemp()) / "sales.xlsx"
>>> wb = openpyxl.Workbook()
>>> ws = wb.active
>>> ws.append(["Region", "Sales"])
>>> ws.append(["North", 1234567.891])
>>> ws.append(["A note long enough to spill"])
>>> ws["B2"].number_format = "#,##0.00"
>>> wb.save(path)
>>> sheet = sheetpeek.read(path)
>>> sheet.cell(2, 2).value, sheet.cell(2, 2).number_format
(1234567.891, '#,##0.00')
```

## Showing it

{func}`sheetpeek.render` shows a sheet with a backend: `text`, `rich`, `html`, `markdown`,
`csv` or `tsv`. Fixed-width output gives every column its width in characters, as Excel lays it out: a
number that does not fit shows `#`, and text spills into empty neighbors.

```python
>>> print(sheetpeek.render(sheet, "text"))
Region   Sales
North    ########
A note long enough to spill
<BLANKLINE>
>>> print(sheetpeek.render(sheet, "markdown"))
| Region | Sales |
| --- | ---: |
| North | 1,234,567.89 |
| A note long enough to spill |  |
<BLANKLINE>
```

Markdown and CSV have no column widths, so every cell shows its full text. Text that spills past
the last used column runs on into empty columns added for it, as it does in Excel; the layout
marks them as added.

## The rendered cells

{meth}`sheetpeek.Sheet.layout` gives the layout itself, the same data every backend draws from.
In `characters` mode each cell records the characters drawn in its column and the cell they come
from.

```python
>>> rendered = sheet.layout("characters")
>>> cell = rendered.cell(2, 2)
>>> cell.text, cell.full_text, cell.overflow, cell.align
('########', '1,234,567.89', True, 'right')
>>> spilled = rendered.cell(3, 2)
>>> spilled.lines, spilled.source
(('ng enoug',), (3, 1))
```

## Excel's geometry

In `excel` mode the layout is in points, as Excel for Mac lays the sheet out: each column's x
and width, each row's y and height, and each text as a {class}`sheetpeek.TextRun` with its
start, baseline, width and clip window. A number shows as many `#` as Excel shows.

```python
>>> excel = sheet.layout("excel")
>>> [(c.letter, c.x_pt, c.width_pt) for c in excel.columns]
[('A', 0.0, 53.0), ('B', 53.0, 53.0), ('C', 106.0, 53.0)]
>>> run = excel.rows[2].runs[0]
>>> run.text, run.x, run.y, run.clip
('A note long enough to spill', 3.0, 42.0, (1.0, 31.0, 159.0, 45.0))
>>> excel.rows[2].crossed
(True, True)
>>> excel.cell(2, 2).text
'#########'
```

Column C was added for the spill: it lies past the last used column, where the sheet is empty.

## HTML

The `html` backend draws that geometry as a static page, or as a fragment to embed; the
{doc}`html` page describes its options and the CSS names a stylesheet can restyle.

```python
>>> page = sheetpeek.render(sheet, "html")
>>> page.splitlines()[0]
'<!doctype html>'
```

## From the command line

```bash
sheetpeek view sales.xlsx                 # rich on a terminal, plain text when piped
sheetpeek view sales.xlsx --column-headings --row-headings   # letters and numbers
sheetpeek view sales.xlsx -t markdown     # or html, csv, tsv, text, rich
sheetpeek view sales.xlsx --list          # the sheet names
sheetpeek serve reports/                  # every workbook under reports/ as HTML pages
```

`sheetpeek serve` is a small read-only web server, from the standard library: an index of the
`.xlsx` files under the folder, and one page per sheet. It binds to 127.0.0.1 unless `--host`
says otherwise, and it reads only `.xlsx` files inside the folder.
