"""The page's stylesheet and the Calibri font faces (``ARCHITECTURE.md`` §7.3)."""

from __future__ import annotations

# Calibri runs use a named face that is only Calibri or Carlito, installed locally, so that the
# letter-spacing computed for Calibri lands on Calibri's glyphs; without either, a local Arial
# or Helvetica (metric-compatible with each other and with Liberation Sans) scaled to Calibri's
# widths and given its vertical metrics, as Next.js builds fallback fonts.
_CALIBRI = "'sheetpeek-calibri', 'sheetpeek-calibri-fallback', sans-serif"
# size-adjust: Carlito's widths over Arial's on spreadsheet text (digits, numbers, labels),
# measured once; the overrides give the fallback Calibri's ascent 0.952 em and descent 0.269 em.
_FALLBACK_SIZE = {"regular": 91.6, "bold": 90.2, "italic": 91.4, "bold italic": 90.0}
_STYLE_NAMES = {  # style -> (CSS weight, CSS style, suffixes of the local font names)
    "regular": ("400", "normal", ("", "-Regular")),
    "bold": ("700", "normal", (" Bold", "-Bold")),
    "italic": ("400", "italic", (" Italic", "-Italic", " Oblique", "-Oblique")),
    "bold italic": ("700", "italic", (" Bold Italic", "-BoldItalic", " Bold Oblique")),
}
_LOCAL = ("Calibri", "Carlito")
_LOCAL_FALLBACK = ("Arial", "Helvetica", "Liberation Sans")


def _local(families: tuple[str, ...], suffixes: tuple[str, ...]) -> str:
    return ", ".join(f'local("{f}{s}")' for f in families for s in suffixes)


def _font_faces() -> str:
    rules = []
    for style, (weight, slant, suffixes) in _STYLE_NAMES.items():
        size = _FALLBACK_SIZE[style]
        rules.append(
            f'@font-face {{ font-family: "sheetpeek-calibri"; src: {_local(_LOCAL, suffixes)}; '
            f"font-weight: {weight}; font-style: {slant}; }}\n"
            '@font-face { font-family: "sheetpeek-calibri-fallback"; '
            f"src: {_local(_LOCAL_FALLBACK, suffixes)}; font-weight: {weight}; "
            f"font-style: {slant}; size-adjust: {size}%; "
            f"ascent-override: {95.2 / size * 100:.2f}%; "
            f"descent-override: {26.9 / size * 100:.2f}%; line-gap-override: 0%; }}\n"
        )
    return "".join(rules)


# Calibri's and Carlito's vertical metrics in em: ascent and descent, and where the baseline
# sits in a 1 em line box (half the leading above the ascent).
_ASCENT, _DESCENT = 0.952, 0.269
_BASELINE = 0.8418

CSS = (
    _font_faces()
    + """\
.sheetpeek {
  --sheetpeek-paper: #ffffff;
  --sheetpeek-ink: #000000;
  --sheetpeek-grid: #d4d4d4;
  --sheetpeek-heading-paper: #f3f3f3;
  --sheetpeek-heading-ink: #444444;
  --sheetpeek-heading-grid: #c8c8c8;
  position: relative; display: inline-block; vertical-align: top;
  background: var(--sheetpeek-paper); color-scheme: light;
}
.sheetpeek[data-color-scheme="dark"] {
  --sheetpeek-heading-paper: #2b2b2b;
  --sheetpeek-heading-ink: #cccccc;
  --sheetpeek-heading-grid: #3d3d3d;
}
.sheetpeek.sp-nogrid { --sheetpeek-grid: transparent; }
.sheetpeek table {
  table-layout: fixed; border-collapse: separate; border-spacing: 0;
  border-right: 1px solid var(--sheetpeek-grid); border-bottom: 1px solid var(--sheetpeek-grid);
  color: var(--sheetpeek-ink);
}
.sheetpeek td {
  padding: 0; box-sizing: border-box;
  border-left: 1px solid var(--sheetpeek-grid); border-top: 1px solid var(--sheetpeek-grid);
}
.sheetpeek th {
  padding: 0; box-sizing: border-box; overflow: hidden; white-space: nowrap;
  font-size: 9pt; font-weight: normal; text-align: center; vertical-align: middle;
  background: var(--sheetpeek-heading-paper); color: var(--sheetpeek-heading-ink);
  border-left: 1px solid var(--sheetpeek-heading-grid);
  border-top: 1px solid var(--sheetpeek-heading-grid);
}
.sheetpeek .sp-crossed { border-left-color: transparent; }
.sheetpeek .sp-clip { position: absolute; overflow: hidden; }
.sheetpeek .sp-run {
  position: absolute; white-space: pre; line-height: 1;
  font-kerning: none; font-variant-ligatures: none;
}
.sheetpeek .sp-center { transform: translateX(-50%); }
"""
)
"""The stylesheet every page and fragment carries."""
