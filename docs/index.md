# sheetpeek

sheetpeek shows an Excel workbook the way Excel shows it: as HTML, Markdown, terminal output
or plain text. It is pure Python, needs no office suite, and gives programs access to the
rendered cells. Number formats are applied by `xlformat`.

```bash
pip install sheetpeek
pip install "sheetpeek[rich]"   # with styled terminal output
```

Python 3.11 or later. The source is at <https://github.com/sheetmill/sheetpeek>.

sheetpeek is part of the [sheetmill](https://github.com/sheetmill) project, with
[xlformat](https://github.com/sheetmill/xlformat), which applies Excel number formats, and
[xlrecalc](https://github.com/sheetmill/xlrecalc), a formula engine.

```{toctree}
:maxdepth: 2

usage
html
api
changelog
```
