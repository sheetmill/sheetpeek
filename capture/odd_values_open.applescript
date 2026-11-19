-- Open one workbook (argument 1), read A1's text and text orientation, B1's number format and
-- the text B1 shows (string value) on its first sheet, close it without saving. Returns
-- "opened" (or "repaired", when Excel opens it under another name) and those, tab-separated,
-- or "refused<TAB>error number". A property Excel will not give reads as "ERR " and its
-- message. A repaired copy is this script's own and is closed too. Non-interactive: alerts are
-- off while it runs and restored on every exit path. Touches no other open workbook; uses no
-- clipboard.
on run argv
	set xlPath to item 1 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set tabChar to ASCII character 9
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set priorNames to name of every workbook
		set display alerts to false
		try
			try
				open workbook workbook file name xlPath
			on error errMsg number errNum
				set answer to "refused" & tabChar & (errNum as text)
				my closeNew(priorNames)
				set display alerts to true
				return answer
			end try
			if exists workbook xlName then
				set opened to {xlName}
				set outcome to "opened"
			else
				set opened to my newNames(priorNames)
				set outcome to "repaired"
			end if
			if (count of opened) is 0 then
				set display alerts to true
				return "refused" & tabChar & "no workbook"
			end if
			set xlWb to workbook (item 1 of opened)
			set xlSheet to worksheet 1 of xlWb
			set xlA1 to range "A1" of xlSheet
			set xlB1 to range "B1" of xlSheet
			try
				set xlText to (string value of xlA1) as string
			on error errText
				set xlText to "ERR " & errText
			end try
			try
				set xlOrientation to (text orientation of xlA1) as string
			on error errText
				set xlOrientation to "ERR " & errText
			end try
			try
				set xlFormat to (number format of xlB1) as string
			on error errText
				set xlFormat to "ERR " & errText
			end try
			try
				set xlShown to (string value of xlB1) as string
			on error errText
				set xlShown to "ERR " & errText
			end try
			set answer to outcome & tabChar & xlText & tabChar & xlOrientation & tabChar & xlFormat & tabChar & xlShown
			close xlWb saving no
			my closeNew(priorNames)
			set display alerts to true
			return answer
		on error errMsg number errNum
			my closeNew(priorNames)
			set display alerts to true
			error errMsg number errNum
		end try
	end tell
end run

on newNames(priorNames)
	tell application "Microsoft Excel"
		set found to {}
		repeat with xlName in (name of every workbook)
			if (xlName as text) is not in priorNames then set end of found to (xlName as text)
		end repeat
		return found
	end tell
end newNames

on closeNew(priorNames)
	tell application "Microsoft Excel"
		repeat with xlName in my newNames(priorNames)
			try
				close workbook (xlName as text) saving no
			end try
		end repeat
	end tell
end closeNew
