-- Probe: can the number-format color be read in bulk? Opens the lone_zero workbook from the
-- container (built by capture/build.py), tries `color of font object of display format of every
-- cell`, falls back to a per-cell loop, and reports both with timings. Non-interactive; alerts
-- restored on every path; closes without saving; touches no other workbook.
on run
	set xlPath to (POSIX path of (path to home folder)) & "Library/Containers/com.microsoft.Excel/Data/Documents/xlformat/lone_zero.xlsx"
	set xlName to "lone_zero.xlsx"
	set US to ASCII character 31
	set xlOut to {}
	tell application "Microsoft Excel"
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlWb to workbook xlName
			set xlWs to worksheet "cases" of xlWb
			set xlRange to range "B1:B96" of xlWs
			set t0 to current date
			try
				set xlBulk to color of font object of display format of every cell of xlRange
				set end of xlOut to "bulk_ok" & US & ((count of xlBulk) as text) & US & ((item 1 of xlBulk) as text)
			on error errMsg
				set end of xlOut to "bulk_failed" & US & errMsg
			end try
			set t1 to current date
			set end of xlOut to "bulk_seconds" & US & ((t1 - t0) as text)
			-- also try display format of the range as a whole, and font color of the range
			try
				set xlOne to color of font object of display format of range "B1" of xlWs
				set end of xlOut to "single_ok" & US & (xlOne as text)
			on error errMsg
				set end of xlOut to "single_failed" & US & errMsg
			end try
			set t2 to current date
			set xlColors to {}
			repeat with i from 1 to 96
				set end of xlColors to (color of font object of display format of range ("B" & i) of xlWs) as text
			end repeat
			set t3 to current date
			set end of xlOut to "loop_seconds" & US & ((t3 - t2) as text)
			set AppleScript's text item delimiters to US
			set end of xlOut to "loop" & US & (xlColors as text)
			set AppleScript's text item delimiters to ""
			-- what does a plain cell font color read as, and the range-level font color?
			set end of xlOut to "cell_font" & US & ((color of font object of range "B1" of xlWs) as text)
			close xlWb saving no
			set display alerts to true
		on error errMsg number errNum
			set display alerts to true
			if exists workbook xlName then close workbook xlName saving no
			error errMsg number errNum
		end try
	end tell
	set AppleScript's text item delimiters to linefeed
	set xlText to xlOut as text
	set AppleScript's text item delimiters to ""
	return xlText
end run
