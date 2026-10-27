-- Capture template. capture/capture.py fills the {{...}} slots and writes the result next to
-- the workbook as <suite>.capture.applescript, then runs it with osascript.
--
-- What it does: opens the suite workbook by path, optionally switches it to the 1904 date
-- system, records Excel's version and fonts, then for each capture column either copies the
-- column to the clipboard (width-less rendering, read by capture.py with pbpaste) or sets the
-- Excel column width and reads `string value of every cell` (one Apple event per column).
-- Optionally reads each cell's display-format font color (one event per cell) and autofits
-- the measure cells and reads their column widths back. Closes the
-- workbook without saving. Touches no other workbook. Non-interactive. Restores `display
-- alerts` on every exit path. The workbook must live inside Excel's sandbox container, or
-- macOS asks for access with a dialog (CAPTURING.md §2).
--
-- Output: records separated by ASCII 30, fields by ASCII 31.
on run
	set xlPath to "{{PATH}}"
	set xlName to "{{NAME}}"
	set nRows to {{NROWS}}
	set colSpecs to {{COLUMNS}}
	set measureCells to {{MEASURE_CELLS}}
	set useDate1904 to {{DATE1904}}
	set useColors to {{COLORS}}
	set RS to ASCII character 30
	set US to ASCII character 31
	set xlOut to {}
	tell application "Microsoft Excel"
		set display alerts to false
		try
		open workbook workbook file name xlPath
		set xlWb to workbook xlName
		if useDate1904 then set date 1904 of xlWb to true
		set xlWs to worksheet "cases" of xlWb
		set end of xlOut to "excel_version" & US & (version as text)
		set end of xlOut to "standard_font" & US & (standard font as text) & " " & (standard font size as text)
		set end of xlOut to "workbook_font" & US & (name of font object of range "B1" of xlWs) & " " & (font size of font object of range "B1" of xlWs)
		set end of xlOut to "date1904" & US & (date 1904 of xlWb as text)
		activate object xlWs
		repeat with spec in colSpecs
			set xlLetter to item 1 of spec
			set xlWidth to item 2 of spec
			set xlMethod to item 3 of spec
			set xlRange to range (xlLetter & "1:" & xlLetter & nRows) of xlWs
			if xlMethod is "string_value" then
				set column width of range (xlLetter & ":" & xlLetter) of xlWs to xlWidth
				set xlShown to string value of every cell of xlRange
				set AppleScript's text item delimiters to US
				set xlJoined to xlShown as text
				set AppleScript's text item delimiters to ""
				set end of xlOut to "column" & US & xlLetter & US & (column width of range (xlLetter & "1") of xlWs as text) & US & xlJoined
				set end of xlOut to "colpoints" & US & xlLetter & US & (width of range (xlLetter & "1") of xlWs as text)
			else
				copy range xlRange
				set end of xlOut to "clipboard" & US & xlLetter & US & (my readClipboard())
			end if
		end repeat
		if useColors then
			-- Bulk reads of display format return missing values (capture/probe_color_bulk.applescript),
			-- so one Apple event per cell: about 10 ms each.
			set xlColors to {}
			repeat with i from 1 to nRows
				set xlRGB to color of font object of display format of range ("B" & i) of xlWs
				set end of xlColors to ((item 1 of xlRGB) as text) & "," & ((item 2 of xlRGB) as text) & "," & ((item 3 of xlRGB) as text)
			end repeat
			set AppleScript's text item delimiters to US
			set xlJoined to xlColors as text
			set AppleScript's text item delimiters to ""
			set end of xlOut to "colors" & US & xlJoined
		end if
		if (count of measureCells) > 0 then
			set xlMs to worksheet "measure" of xlWb
			set firstCell to item 1 of measureCells
			set lastCell to item -1 of measureCells
			autofit (entire column of range (firstCell & ":" & lastCell) of xlMs)
			set xlWidths to {}
			set xlPoints to {}
			repeat with c in measureCells
				set end of xlWidths to (column width of range c of xlMs) as text
				set end of xlPoints to (width of range c of xlMs) as text
			end repeat
			set AppleScript's text item delimiters to US
			set xlJoined to xlWidths as text
			set xlJoinedPoints to xlPoints as text
			set AppleScript's text item delimiters to ""
			set end of xlOut to "measure" & US & xlJoined
			set end of xlOut to "measure_points" & US & xlJoinedPoints
		end if
		close xlWb saving no
		set display alerts to true
		on error errMsg number errNum
			set display alerts to true
			if exists workbook xlName then close workbook xlName saving no
			error errMsg number errNum
		end try
	end tell
	set AppleScript's text item delimiters to RS
	set xlText to xlOut as text
	set AppleScript's text item delimiters to ""
	return xlText
end run

on readClipboard()
	-- pbpaste gives the clipboard text; cells are separated by linefeeds. Encode as base64 so
	-- leading/trailing spaces and the record separators survive the trip through osascript.
	return do shell script "pbpaste | base64"
end readClipboard
