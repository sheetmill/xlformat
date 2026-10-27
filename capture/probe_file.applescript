-- Probe: open a workbook written by openpyxl (path given as argument), read displayed text and
-- the number format Excel kept for every row, close without saving. Non-interactive.
-- Output: one line per row: id <TAB> value repr <TAB> format repr <TAB> displayed <TAB> kept format
on run argv
	set xlTab to ASCII character 9
	set xlPath to item 1 of argv
	set xlRows to (item 2 of argv) as integer
	set xlName to do shell script "basename " & quoted form of xlPath
	set xlOut to {}
	tell application "Microsoft Excel"
		set display alerts to false
		try
		open workbook workbook file name xlPath
		set xlWb to workbook xlName
		set xlWs to worksheet 1 of xlWb
		set xlShown to string value of every cell of range ("A1:A" & xlRows) of xlWs
		set xlMeta to value of range ("B1:D" & xlRows) of xlWs
		repeat with i from 1 to xlRows
			set xlRec to item i of xlMeta
			set xlKept to number format of range ("A" & i) of xlWs
			set end of xlOut to (item 1 of xlRec) & xlTab & (item 2 of xlRec) & xlTab & (item 3 of xlRec) & xlTab & (item i of xlShown) & xlTab & xlKept
		end repeat
		close xlWb saving no
		set display alerts to true
		on error errMsg number errNum
			set display alerts to true
			error errMsg number errNum
		end try
	end tell
	set AppleScript's text item delimiters to linefeed
	set xlText to xlOut as text
	set AppleScript's text item delimiters to ""
	return xlText
end run
