-- Open a workbook (argument 1), save it as PDF (argument 2), close it without saving.
-- capture/alignment.py runs it. Non-interactive: alerts are off while it runs and restored on
-- every exit path. Opens only the workbook it is given, by path, inside Excel's container;
-- touches no other open workbook.
on run argv
	set xlPath to item 1 of argv
	set xlPdf to item 2 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	tell application "Microsoft Excel"
		set display alerts to false
		try
			set xlVersion to version as text
			open workbook workbook file name xlPath
			set xlWb to workbook xlName
			save workbook as xlWb filename xlPdf file format PDF file format
			close xlWb saving no
			set display alerts to true
		on error errMsg number errNum
			try
				if exists workbook xlName then close workbook xlName saving no
			end try
			set display alerts to true
			error errMsg number errNum
		end try
	end tell
	return xlVersion
end run
