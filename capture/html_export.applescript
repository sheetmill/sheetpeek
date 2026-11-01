-- Open a workbook (argument 1), read the displayed text of a range (argument 3), save the
-- workbook as a web page (argument 2), close it without saving. Non-interactive: alerts are off
-- while it runs and restored on every exit path. Touches no other open workbook.
-- Output: the displayed text of every cell of the range, row by row, one per line.
on run argv
	set xlPath to item 1 of argv
	set xlHtml to item 2 of argv
	set xlRange to item 3 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set xlOut to {}
	tell application "Microsoft Excel"
		-- A workbook of this name already open in Excel is someone's: leave it and stop.
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlWb to workbook xlName
			set xlWs to worksheet 1 of xlWb
			set xlOut to string value of every cell of range xlRange of xlWs
			save workbook as xlWb filename xlHtml file format HTML file format
			set xlHtmlName to do shell script "basename " & quoted form of xlHtml
			close workbook xlHtmlName saving no
			set display alerts to true
		on error errMsg number errNum
			set display alerts to true
			error errMsg number errNum
		end try
	end tell
	set AppleScript's text item delimiters to linefeed
	set xlText to xlOut as text
	set AppleScript's text item delimiters to ""
	return xlText
end run
