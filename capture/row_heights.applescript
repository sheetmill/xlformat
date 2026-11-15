-- Read the height of rows 1 to N, and the sheet's standard height, from one workbook (argument 1,
-- a path inside Excel's container, argument 3 its file name) that capture/row_heights.py built;
-- the workbook is found by that name, never as the active one; close it without saving.
-- Alerts are off while the workbook is open and back on at every exit, errors included. Opens
-- only the workbook it is given; no clipboard.
--
-- Output: fields separated by ASCII 31: "standard", the standard height, then each row's height
-- in points, rows 1 to N.
on run argv
	set xlPath to item 1 of argv
	set nRows to (item 2 of argv) as integer
	set xlName to item 3 of argv
	set US to ASCII character 31
	tell application "Microsoft Excel"
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlWb to workbook xlName
			set xlWs to worksheet 1 of xlWb
			set xlOut to "standard" & US & ((standard height of xlWs) as text)
			repeat with r from 1 to nRows
				set xlOut to xlOut & US & ((row height of row r of xlWs) as text)
			end repeat
			close xlWb saving no
			set display alerts to true
		on error xlMessage number xlNumber
			set display alerts to true
			error xlMessage number xlNumber
		end try
	end tell
	return xlOut
end run
