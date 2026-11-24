-- Prints Excel's standard font and size (Preferences > General), tab separated: what Excel
-- uses for new workbooks. Reads two application properties; opens nothing, alerts untouched,
-- no clipboard.
tell application "Microsoft Excel"
	set fontName to standard font as text
	set fontSize to standard font size as text
end tell
return fontName & (ASCII character 9) & fontSize
