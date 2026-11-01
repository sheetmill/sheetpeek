-- Open a workbook (argument 1), save it as PDF (argument 2), close it without saving.
-- Non-interactive: alerts are off while it runs and restored on every exit path. Touches no
-- other open workbook, and stops before
-- opening anything if a workbook of the same name is already open.
on run argv
	set xlPath to item 1 of argv
	set xlPdf to item 2 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	tell application "Microsoft Excel"
		-- A workbook of this name already open in Excel is someone's: leave it and stop.
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlWb to workbook xlName
			save workbook as xlWb filename xlPdf file format PDF file format
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
	return "ok"
end run
