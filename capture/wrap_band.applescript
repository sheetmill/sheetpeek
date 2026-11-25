-- Arguments in pairs: a workbook's path, then the rows to read in it, comma separated ("1,2,3").
-- Opens each workbook, reads, in column A of each row given, the font name and size Excel
-- applied, the row height it lays out and column A's width (column width in characters and width
-- in points), and closes the workbook without saving. Returns one tab-separated line per row:
-- file name, row, font name, font size, row height, column width, width.
-- Non-interactive: alerts are off while it runs and restored on every exit path. Touches no
-- other open workbook; uses no clipboard.
on run argv
	set tabChar to ASCII character 9
	set output to {}
	if ((count of argv) mod 2) is not 0 then error "expected pairs of a path and its rows"
	tell application "Microsoft Excel"
		-- A workbook of one of these names already open in Excel is someone's: stop first.
		repeat with k from 1 to (count of argv) by 2
			set xlName to do shell script "basename " & quoted form of (item k of argv as text)
			if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		end repeat
		set display alerts to false
		repeat with k from 1 to (count of argv) by 2
			set xlPath to item k of argv as text
			set xlName to do shell script "basename " & quoted form of xlPath
			set AppleScript's text item delimiters to ","
			set rowList to text items of (item (k + 1) of argv as text)
			set AppleScript's text item delimiters to ""
			try
				open workbook workbook file name xlPath
				set xlWb to workbook xlName
				set xlWs to worksheet 1 of xlWb
				repeat with rowRef in rowList
					set rowText to contents of rowRef
					set target to range ("A" & rowText) of xlWs
					set fontName to name of font object of target
					set fontSize to font size of font object of target
					set h to row height of target
					set cw to column width of target
					set wpt to width of target
					set end of output to xlName & tabChar & rowText & tabChar & fontName & tabChar & (fontSize as text) & tabChar & (h as text) & tabChar & (cw as text) & tabChar & (wpt as text)
				end repeat
				close xlWb saving no
			on error errMsg number errNum
				try
					close workbook xlName saving no
				end try
				set display alerts to true
				error errMsg number errNum
			end try
		end repeat
		set display alerts to true
	end tell
	set AppleScript's text item delimiters to linefeed
	set answer to output as text
	set AppleScript's text item delimiters to ""
	return answer
end run
