-- Open each workbook given as an argument, read the row heights Excel lays out for rows without
-- a stored height, close it without saving. Returns one tab-separated line per workbook: file
-- name, the font name and size Excel applied to A1, the row height of rows 1, 2, 3, 5 and 100,
-- and the worksheet's standard height ("n/a" if Excel does not report it). Non-interactive: alerts are off while it runs and restored on every
-- exit path. Touches no other open workbook; uses no clipboard.
on run argv
	set tabChar to ASCII character 9
	set output to {}
	tell application "Microsoft Excel"
		-- A workbook of one of these names already open in Excel is someone's: stop first.
		repeat with xlPath in argv
			set xlName to do shell script "basename " & quoted form of (xlPath as text)
			if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		end repeat
		set display alerts to false
		repeat with xlPath in argv
			set xlPath to xlPath as text
			set xlName to do shell script "basename " & quoted form of xlPath
			try
				open workbook workbook file name xlPath
				set xlWb to workbook xlName
				set xlWs to worksheet 1 of xlWb
				set fontName to name of font object of range "A1" of xlWs
				set fontSize to font size of font object of range "A1" of xlWs
				set h1 to row height of range "A1" of xlWs
				set h2 to row height of range "A2" of xlWs
				set h3 to row height of range "A3" of xlWs
				set h5 to row height of range "A5" of xlWs
				set h100 to row height of range "A100" of xlWs
				try
					set std to (standard height of xlWs) as text
				on error
					set std to "n/a"
				end try
				close xlWb saving no
				set end of output to xlName & tabChar & fontName & tabChar & (fontSize as text) & tabChar & (h1 as text) & tabChar & (h2 as text) & tabChar & (h3 as text) & tabChar & (h5 as text) & tabChar & (h100 as text) & tabChar & std
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
