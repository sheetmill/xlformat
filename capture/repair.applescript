-- Read a workbook Excel has repaired, found by name (argument 1), then close it without saving.
-- capture/repair.py runs it only once Excel's repair dialog has been answered and `read` is run:
-- nothing is sent to Excel while that dialog may be open. Alerts are left as they are (on).
--
-- Output: one line per row, fields separated by ASCII 31: row, then per value column the
-- displayed text and the number format; the first line is "window" and the captions of the
-- workbook's windows (a repaired one's says "Repaired").
on run argv
	set xlName to item 1 of argv
	set nRows to (item 2 of argv) as integer
	set nCols to (item 3 of argv) as integer
	set US to ASCII character 31
	set xlOut to {}
	tell application "Microsoft Excel"
		set xlWb to workbook xlName
		set xlCaptions to ""
		try
			repeat with xlWin in (every window)
				set xlCaption to (caption of xlWin) as text
				if xlCaption contains xlName then set xlCaptions to xlCaptions & US & xlCaption
			end repeat
		end try
		set end of xlOut to "window" & xlCaptions
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
