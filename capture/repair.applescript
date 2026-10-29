-- Open a workbook Excel will want to repair (argument 1), WITH alerts on, so Excel shows its
-- repair dialogs and a person answers them; then read every case row back and close the
-- repaired workbook without saving. capture/repair.py runs it, announced, with the person at
-- the keyboard. The only interactive capture script: alerts stay on (as they were) on every
-- path. Opens only the workbook it is given, by path, inside Excel's container.
--
-- Output: one line per row, fields separated by ASCII 31: row, then per value column the
-- displayed text and the number format; the first line is the repaired workbook's name.
on run argv
	set xlPath to item 1 of argv
	set nRows to (item 2 of argv) as integer
	set nCols to (item 3 of argv) as integer
	set US to ASCII character 31
	set xlOut to {}
	tell application "Microsoft Excel"
		set display alerts to true
		activate
		with timeout of 600 seconds
			open workbook workbook file name xlPath
		end timeout
		set xlWb to active workbook
		set end of xlOut to "workbook" & US & (name of xlWb)
		set xlWs to worksheet 1 of xlWb
		repeat with r from 1 to nRows
			set xlLine to (r as text)
			repeat with c from 2 to nCols + 1
				set xlCell to cell r of column c of xlWs
				set xlLine to xlLine & US & (string value of xlCell) & US & (number format of xlCell)
			end repeat
			set end of xlOut to xlLine
		end repeat
		close xlWb saving no
	end tell
	set AppleScript's text item delimiters to linefeed
	set xlText to xlOut as text
	set AppleScript's text item delimiters to ""
	return xlText
end run
