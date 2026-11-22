-- Requests to Excel for the screen captures (capture/screen/session.py runs it, one request a
-- call). The first argument names the request:
--   count                      how many workbooks are open
--   open PATH                  open a workbook; stops if one of its name is already open
--   close NAME                 close a workbook without saving
--   alerts                     turn display alerts back on
--   view NAME PARK ZOOM        bring the workbook forward, set the zoom, select PARK (a cell off
--                              the measured area) and scroll back to A1
--   scroll NAME ROW            scroll the active window so ROW is its top row (column A at the
--                              left); returns the address of the window's visible range
--   window NAME                the workbook's active window's state, read and not set: zoom,
--                              view, freeze panes, split row and column, scroll row and column
--   maximize NAME              maximize the workbook's active window; returns its window state
--   geometry NAME              the active window's usable width and height in points and its
--                              window state, joined by ";"
--   save NAME PATH             save the workbook as PATH (Excel's own file)
--   read NAME CELL...          what Excel shows for each cell: its display format, its own
--                              format, its borders and its geometry, one record a property
--   follow NAME ACTIVE SHEET CELL
--                              make sheet ACTIVE the active sheet (A1 selected), read the
--                              hyperlinks Excel holds on CELL of SHEET and the cell's formula,
--                              follow the first one, and read the active sheet and cell after
--                              it; one record a property. For links into the workbook only: a
--                              link to a URL or a file would open another application
-- Non-interactive: display alerts are off while a workbook opens, closes or saves, and restored
-- on every exit path. Opens only the workbook it is given; uses no clipboard.
on run argv
	set request to item 1 of argv
	if request is "count" then
		tell application "Microsoft Excel" to return (count of workbooks) as string
	else if request is "alerts" then
		tell application "Microsoft Excel" to set display alerts to true
		return ""
	else if request is "open" then
		return my openBook(item 2 of argv)
	else if request is "close" then
		return my closeBook(item 2 of argv)
	else if request is "view" then
		return my viewBook(item 2 of argv, item 3 of argv, (item 4 of argv) as integer)
	else if request is "scroll" then
		return my scrollBook(item 2 of argv, (item 3 of argv) as integer)
	else if request is "window" then
		return my windowState(item 2 of argv)
	else if request is "maximize" then
		return my maximizeBook(item 2 of argv)
	else if request is "geometry" then
		return my windowGeometry(item 2 of argv)
	else if request is "save" then
		return my saveBook(item 2 of argv, item 3 of argv)
	else if request is "read" then
		return my readCells(item 2 of argv, items 3 thru -1 of argv)
	else if request is "follow" then
		return my followLink(item 2 of argv, item 3 of argv, item 4 of argv, item 5 of argv)
	end if
	error "unknown request: " & request
end run

on openBook(xlPath)
	set xlName to do shell script "basename " & quoted form of xlPath
	tell application "Microsoft Excel"
		-- A workbook of this name already open in Excel is someone's: leave it and stop.
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlWb to workbook xlName
		on error errText
			set display alerts to true
			error errText
		end try
		set display alerts to true
	end tell
	return xlName
end openBook

on closeBook(xlName)
	tell application "Microsoft Excel"
		set display alerts to false
		try
			close workbook xlName saving no
		end try
		set display alerts to true
		return (count of workbooks) as string
	end tell
end closeBook

on viewBook(xlName, park, zoomLevel)
	tell application "Microsoft Excel"
		activate
		set xlWb to workbook xlName
		activate object xlWb
		-- the zoom first: a zoom change recenters the window on the active cell
		set zoom of active window to zoomLevel
		select range park of worksheet 1 of xlWb
		set scroll row of active window to 1
		set scroll column of active window to 1
		return (zoom of active window) as string
	end tell
end viewBook

on maximizeBook(xlName)
	tell application "Microsoft Excel"
		activate
		activate object workbook xlName
		set window state of active window to window state maximized
		return (window state of active window) as string
	end tell
end maximizeBook

on windowGeometry(xlName)
	tell application "Microsoft Excel"
		activate object workbook xlName
		set xlWindow to active window
		return ((usable width of xlWindow) as string) & ";" & ((usable height of xlWindow) as string) & ";" & ((window state of xlWindow) as string)
	end tell
end windowGeometry

on scrollBook(xlName, topRow)
	tell application "Microsoft Excel"
		activate object workbook xlName
		set scroll row of active window to topRow
		set scroll column of active window to 1
		return (get address of (visible range of active window)) as string
	end tell
end scrollBook

-- The window's state as key=value fields joined by ";"; a field Excel refuses is "ERR". The
-- window's view comes back as Excel names it ("normal view"), or as a raw constant for page break
-- preview, which the dictionary leaves out of its enumeration.
on windowState(xlName)
	set found to {}
	tell application "Microsoft Excel"
		activate object workbook xlName
		set xlWindow to active window
		try
			set end of found to "zoom=" & ((zoom of xlWindow) as string)
		on error
			set end of found to "zoom=ERR"
		end try
		try
			set end of found to "view=" & ((view of xlWindow) as string)
		on error
			set end of found to "view=ERR"
		end try
		try
			set end of found to "freeze_panes=" & ((freeze panes of xlWindow) as string)
		on error
			set end of found to "freeze_panes=ERR"
		end try
		try
			set end of found to "split_row=" & ((split row of xlWindow) as string)
		on error
			set end of found to "split_row=ERR"
		end try
		try
			set end of found to "split_column=" & ((split column of xlWindow) as string)
		on error
			set end of found to "split_column=ERR"
		end try
		try
			set end of found to "scroll_row=" & ((scroll row of xlWindow) as string)
		on error
			set end of found to "scroll_row=ERR"
		end try
		try
			set end of found to "scroll_column=" & ((scroll column of xlWindow) as string)
		on error
			set end of found to "scroll_column=ERR"
		end try
	end tell
	set saved to AppleScript's text item delimiters
	set AppleScript's text item delimiters to ";"
	set output to found as string
	set AppleScript's text item delimiters to saved
	return output
end windowState

on saveBook(xlName, xlPath)
	tell application "Microsoft Excel"
		set display alerts to false
		try
			save workbook as workbook xlName filename xlPath file format workbook normal file format
		on error errText
			set display alerts to true
			error errText
		end try
		set display alerts to true
	end tell
	return ""
end saveBook

on readCells(xlName, cellRefs)
	set us to character id 31
	set rs to character id 30
	set saved to AppleScript's text item delimiters
	set AppleScript's text item delimiters to ","
	set output to ""
	set propertyNames to {"fill", "pattern", "font_color", "bold", "italic", "underline", "strike", "font_name", "font_size", "number_format", "h_align", "v_align", "wrap", "orientation", "indent", "shrink"}
	set edges to {"top", "bottom", "left", "right", "diag_down", "diag_up"}
	tell application "Microsoft Excel"
		set xlWs to worksheet 1 of workbook xlName
		repeat with cellRef in cellRefs
			set cellName to cellRef as string
			set xlCell to range cellName of xlWs
			set shown to display format of xlCell
			repeat with propertyName in propertyNames
				set output to output & cellName & us & "display." & propertyName & us & my cellProperty(shown, propertyName as string) & rs
				set output to output & cellName & us & "own." & propertyName & us & my cellProperty(xlCell, propertyName as string) & rs
			end repeat
			repeat with edgeName in edges
				repeat with partName in {"style", "weight", "color"}
					set readValue to my borderPart(xlCell, edgeName as string, partName as string)
					set output to output & cellName & us & "own.border." & edgeName & "." & partName & us & readValue & rs
				end repeat
			end repeat
			repeat with partName in {"left", "top", "width", "height", "text"}
				set output to output & cellName & us & "geometry." & partName & us & my geometryPart(xlCell, partName as string) & rs
			end repeat
		end repeat
	end tell
	set AppleScript's text item delimiters to saved
	return output
end readCells

-- One property of a cell or of its display format, as text; "ERR " and the message when Excel
-- refuses it. Colors come back as r,g,b (the text item delimiters are a comma).
on cellProperty(target, propertyName)
	tell application "Microsoft Excel"
		try
			if propertyName is "fill" then return (color of interior object of target) as string
			if propertyName is "pattern" then return (pattern of interior object of target) as string
			if propertyName is "font_color" then return (color of font object of target) as string
			if propertyName is "bold" then return (bold of font object of target) as string
			if propertyName is "italic" then return (italic of font object of target) as string
			if propertyName is "underline" then return (underline of font object of target) as string
			if propertyName is "strike" then return (strikethrough of font object of target) as string
			if propertyName is "font_name" then return (name of font object of target) as string
			if propertyName is "font_size" then return (font size of font object of target) as string
			if propertyName is "number_format" then return (number format of target) as string
			if propertyName is "h_align" then return (horizontal alignment of target) as string
			if propertyName is "v_align" then return (vertical alignment of target) as string
			if propertyName is "wrap" then return (wrap text of target) as string
			if propertyName is "orientation" then return (text orientation of target) as string
			if propertyName is "indent" then return (indent level of target) as string
			if propertyName is "shrink" then return (shrink to fit of target) as string
		on error errText
			return "ERR " & errText
		end try
	end tell
	return "ERR unknown property"
end cellProperty

-- A cell's own border on one edge: its line style, weight or color. Display formats carry no
-- borders in Excel's dictionary, so displayed borders come from the screen.
on borderPart(xlCell, edgeName, partName)
	tell application "Microsoft Excel"
		try
			if edgeName is "top" then set b to get border xlCell which border edge top
			if edgeName is "bottom" then set b to get border xlCell which border edge bottom
			if edgeName is "left" then set b to get border xlCell which border edge left
			if edgeName is "right" then set b to get border xlCell which border edge right
			if edgeName is "diag_down" then set b to get border xlCell which border diagonal down
			if edgeName is "diag_up" then set b to get border xlCell which border diagonal up
			if partName is "style" then return (line style of b) as string
			if partName is "weight" then return (weight of b) as string
			if partName is "color" then return (color of b) as string
		on error errText
			return "ERR " & errText
		end try
	end tell
	return "ERR unknown partName"
end borderPart

-- Follow a cell's first hyperlink with another sheet possibly active, and say where Excel went.
-- Each step that Excel refuses is recorded as "ERR " and its message; the steps after it still run.
on followLink(xlName, activeName, sheetName, cellRef)
	set us to character id 31
	set rs to character id 30
	set output to ""
	tell application "Microsoft Excel"
		set xlWb to workbook xlName
		activate object xlWb
		set display alerts to false
		try
			activate object worksheet activeName of xlWb
			select range "A1" of worksheet activeName of xlWb
			set sheetBefore to (name of active sheet) as string
		on error errText
			set sheetBefore to "ERR " & errText
		end try
		set output to output & "active.before" & us & sheetBefore & rs
		try
			set xlCell to range cellRef of worksheet sheetName of xlWb
		on error errText
			set display alerts to true
			error errText
		end try
		try
			set found to (count of hyperlinks of xlCell) as string
		on error errText
			set found to "ERR " & errText
		end try
		set output to output & "links.count" & us & found & rs
		try
			set formulaText to (formula of xlCell) as string
		on error errText
			set formulaText to "ERR " & errText
		end try
		set output to output & "formula" & us & formulaText & rs
		try
			set xlLink to hyperlink 1 of xlCell
			set output to output & "link.address" & us & ((address of xlLink) as string) & rs
			set output to output & "link.sub_address" & us & ((sub address of xlLink) as string) & rs
			follow xlLink
			set followed to "yes"
		on error errText
			set followed to "ERR " & errText
		end try
		set output to output & "followed" & us & followed & rs
		try
			set output to output & "after.sheet" & us & ((name of active sheet) as string) & rs
			set output to output & "after.cell" & us & ((get address of active cell) as string) & rs
		on error errText
			set output to output & "after.sheet" & us & "ERR " & errText & rs
		end try
		set display alerts to true
	end tell
	return output
end followLink

-- Where the cell is, in points from the sheet's origin, and the text it shows.
on geometryPart(xlCell, partName)
	tell application "Microsoft Excel"
		try
			if partName is "left" then return (left position of xlCell) as string
			if partName is "top" then return (top of xlCell) as string
			if partName is "width" then return (width of xlCell) as string
			if partName is "height" then return (height of xlCell) as string
			if partName is "text" then return (string value of xlCell) as string
		on error errText
			return "ERR " & errText
		end try
	end tell
	return "ERR unknown partName"
end geometryPart
