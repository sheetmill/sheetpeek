"""Route C: what Excel for Mac draws on the screen, from window captures (``CAPTURING.md``).

``session`` drives Excel through ``excel.applescript``; ``window`` captures the window and maps
cells to pixels through two marker cells; ``run`` builds a suite's workbook, captures it and
reloads a saved run; ``png`` reads the images. ``python -m capture.screen`` runs suites one at a
time. Standard library and openpyxl only.
"""
