-- One request to Excel for capture/screen_vertical.py, which runs it once per step:
--   open PATH NAME     open the workbook (alerts off only while opening), bring Excel to the front,
--                      zoom 100, select a cell outside the cases; returns the workbook's name
--   heights NAME N     the height in points of rows 1 to N, separated by commas
--   scroll NAME ROW    make ROW the top visible row; returns the scroll row Excel reports
--   close NAME         close the workbook without saving and turn alerts on
-- Opens only the workbook it is given, by path, inside Excel's container, and finds it by name;
-- the run starts with no workbook open, so the active window is that workbook's.
on run argv
	set xlCommand to item 1 of argv
	tell application "Microsoft Excel"
		if xlCommand is "open" then
			set xlPath to item 2 of argv
			set xlName to item 3 of argv
			set display alerts to false
			try
				open workbook workbook file name xlPath
			on error xlMessage number xlNumber
				set display alerts to true
				error xlMessage number xlNumber
			end try
			set display alerts to true
			activate
			set zoom of active window to 100
			select range "F1" of worksheet 1 of workbook xlName
			return name of workbook xlName
		else if xlCommand is "heights" then
			set xlWs to worksheet 1 of workbook (item 2 of argv)
			set xlCount to (item 3 of argv) as integer
			set xlOut to {}
			repeat with r from 1 to xlCount
				set end of xlOut to (row height of row r of xlWs) as text
			end repeat
			set AppleScript's text item delimiters to ","
			set xlText to xlOut as text
			set AppleScript's text item delimiters to ""
			return xlText
		else if xlCommand is "scroll" then
			set scroll row of active window to ((item 3 of argv) as integer)
			return (scroll row of active window) as text
		else if xlCommand is "close" then
			close workbook (item 2 of argv) saving no
			set display alerts to true
			return "closed"
		end if
	end tell
	error "unknown command " & xlCommand
end run
