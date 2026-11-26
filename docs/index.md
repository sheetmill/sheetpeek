# sheetpeek

sheetpeek shows an Excel workbook the way Excel shows it: as HTML, Markdown, terminal output
or plain text. It is pure Python, needs no office suite, and gives programs access to the
rendered cells. Number formats are applied by `xlformat`. The defaults follow Excel for Windows
at 96 dpi with en-US regional settings; `platform="mac"` shows a workbook as Excel for Mac does.

```bash
pip install sheetpeek
pip install "sheetpeek[rich]"   # with styled terminal output
```

Python 3.11 or later. The source is at <https://github.com/sheetmill/sheetpeek>.

sheetpeek is part of the [sheetmill](https://github.com/sheetmill) project, with
[xlio](https://github.com/sheetmill/xlio), which reads and writes the workbooks,
[xlformat](https://github.com/sheetmill/xlformat), which applies Excel number formats, and
[xlrecalc](https://github.com/sheetmill/xlrecalc), a formula engine.

```{toctree}
:maxdepth: 2

usage
html
api
changelog
```
