-- Open one workbook of built-in number format ids (argument 1, a path inside Excel's container;
-- argument 2 its file name; argument 3 the number of rows; argument 4 the number of value columns),
-- read each row's number format (as stored and as shown locally) and every value cell's displayed
-- text, and close it without saving. capture/builtin_formats.py runs it. Alerts are off while the
-- workbook is open and back on at every exit, errors included; no clipboard. A workbook Excel
-- refuses makes the script fail with Excel's error, which the caller records.
--
-- Output: one line per row, fields separated by ASCII 31: number format, number format local, then
-- each value cell's string value.
on run argv
	set xlPath to item 1 of argv
	set xlName to item 2 of argv
	set nRows to (item 3 of argv) as integer
	set nCols to (item 4 of argv) as integer
	set US to ASCII character 31
	set xlOut to {}
	tell application "Microsoft Excel"
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlWs to worksheet 1 of workbook xlName
			set column width of range ("B1:" & (item (nCols + 1) of {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"}) & "1") of xlWs to 40
			repeat with r from 1 to nRows
				set xlFirst to cell r of column 2 of xlWs
				set xlLine to ((number format of xlFirst) as text) & US & ((number format local of xlFirst) as text)
				repeat with c from 2 to nCols + 1
					set xlLine to xlLine & US & ((string value of cell r of column c of xlWs) as text)
				end repeat
				set end of xlOut to xlLine
			end repeat
			close workbook xlName saving no
			set display alerts to true
		on error xlMessage number xlNumber
			try
				if exists workbook xlName then close workbook xlName saving no
			end try
			set display alerts to true
			error xlMessage number xlNumber
		end try
	end tell
	set AppleScript's text item delimiters to linefeed
	set xlText to xlOut as text
	set AppleScript's text item delimiters to ""
	return xlText
end run
