-- Open a workbook (argument 1), read the string value Excel shows for each range given as
-- "sheet!range" (arguments 2 and on), close it without saving. Bulk reads, one per range.
-- Returns, per range, one line per cell: the range argument, the cell's position in the range
-- (row,column), and the string value. Non-interactive: alerts are off while it runs and
-- restored on every exit path. Touches no other open workbook; uses no clipboard.
on run argv
	set xlPath to item 1 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set tabChar to ASCII character 9
	set output to {}
	tell application "Microsoft Excel"
		-- A workbook of this name already open in Excel is someone's: leave it and stop.
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlWb to workbook xlName
			repeat with k from 2 to count of argv
				set spec to item k of argv
				set AppleScript's text item delimiters to "!"
				set sheetName to text item 1 of spec
				set cellRange to text item 2 of spec
				set AppleScript's text item delimiters to ""
				set shown to string value of range cellRange of worksheet sheetName of xlWb
				if class of shown is not list then set shown to {{shown}}
				repeat with r from 1 to count of shown
					set rowValues to item r of shown
					if class of rowValues is not list then set rowValues to {rowValues}
					repeat with c from 1 to count of rowValues
						set end of output to spec & tabChar & r & "," & c & tabChar & (item c of rowValues)
					end repeat
				end repeat
			end repeat
			close xlWb saving no
			set display alerts to true
		on error errMsg number errNum
			try
				close workbook xlName saving no
			end try
			set display alerts to true
			error errMsg number errNum
		end try
	end tell
	set AppleScript's text item delimiters to linefeed
	set answer to output as text
	set AppleScript's text item delimiters to ""
	return answer
end run
