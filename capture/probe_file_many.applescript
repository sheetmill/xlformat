-- Probe: open every inv_*.xlsx in a directory (built by openpyxl, one dialog-rejected format
-- each), record displayed text and the number format Excel kept, or the open error. Closes
-- only workbooks it opened, never saves. Non-interactive.
on run argv
	set xlTab to ASCII character 9
	set xlDir to item 1 of argv
	set xlRows to (item 2 of argv) as integer
	set xlNames to paragraphs of (do shell script "ls " & quoted form of xlDir & " | grep '^inv_.*xlsx$'")
	log "files: " & (count of xlNames)
	set xlOut to {}
	tell application "Microsoft Excel"
		set display alerts to false
		try
		repeat with xlNameRef in xlNames
			set xlName to xlNameRef as text
			set xlPath to xlDir & "/" & xlName
			try
				open workbook workbook file name xlPath
				set xlWb to workbook xlName
				set xlWs to worksheet 1 of xlWb
				set xlShown to string value of every cell of range ("A1:A" & xlRows) of xlWs
				set xlMeta to value of range ("B1:D" & xlRows) of xlWs
				repeat with i from 1 to xlRows
					set xlRec to item i of xlMeta
					set xlKept to number format of range ("A" & i) of xlWs
					set end of xlOut to xlName & xlTab & (item 1 of xlRec) & xlTab & (item 2 of xlRec) & xlTab & (item 3 of xlRec) & xlTab & (item i of xlShown) & xlTab & xlKept
				end repeat
				close xlWb saving no
			on error e number num
				set end of xlOut to xlName & xlTab & "ERROR " & num & ": " & e
				if exists workbook xlName then close workbook xlName saving no
			end try
		end repeat
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
