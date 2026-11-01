-- Open a workbook (argument 1), read the fill and font color Excel resolves for each cell whose
-- address follows (arguments 2 and on), close it without saving. Returns one tab-separated line
-- per cell: its address, the fill color and the font color as r,g,b. Each cell is read on its
-- own: a bulk read of colors returns missing value for every cell (xlformat's capture found
-- this). Non-interactive: alerts are off while it runs and restored on every exit path. Touches
-- no other open workbook, stops before opening anything if one of the same name is open, and
-- uses no clipboard.
on run argv
	set xlPath to item 1 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set tabChar to ASCII character 9
	set output to {}
	set step to "opening"
	tell application "Microsoft Excel"
		-- A workbook of this name already open in Excel is someone's: leave it and stop.
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlWb to workbook xlName
			set xlWs to worksheet 1 of xlWb
			repeat with k from 2 to count of argv
				set cellRef to item k of argv
				set step to "range " & cellRef
				set xlCell to range cellRef of xlWs
				set step to "fill of " & cellRef
				set fillColor to color of interior object of xlCell
				set step to "font of " & cellRef
				set fontColor to color of font object of xlCell
				set end of output to cellRef & tabChar & my joined(fillColor) & tabChar & my joined(fontColor)
			end repeat
			close xlWb saving no
			set display alerts to true
		on error errMsg number errNum
			try
				close workbook xlName saving no
			end try
			set display alerts to true
			error errMsg & " (" & step & ")" number errNum
		end try
	end tell
	set AppleScript's text item delimiters to linefeed
	set answer to output as text
	set AppleScript's text item delimiters to ""
	return answer
end run

on joined(rgbList)
	set AppleScript's text item delimiters to ","
	set answer to rgbList as text
	set AppleScript's text item delimiters to ""
	return answer
end joined
