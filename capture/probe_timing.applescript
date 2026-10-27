-- Probe: throughput of per-cell `string value` versus bulk forms over 1,000 cells.
-- Creates one new workbook, never saves, closes it. Non-interactive.
on now()
	return (do shell script "perl -MTime::HiRes=time -e 'printf \"%.3f\", time'") as real
end now

on sh(cmd)
	return do shell script cmd
end sh

tell application "Microsoft Excel"
	set display alerts to false
	try
	set out to {}
	set wb to make new workbook
	tell active sheet of wb
		set formula of range "A1:A1000" to "=ROW()*1234.5678"
		set number format of range "A1:A1000" to "#,##0.00"
		set formula of range "B1:B1000" to "=ROW()"
		set t0 to my now()
		set acc to {}
		repeat with i from 1 to 1000
			set end of acc to (string value of range ("A" & i))
		end repeat
		set t1 to my now()
		set end of out to "per-cell 1000 string value: " & (t1 - t0) & " s; sample " & (item 1000 of acc)
		set t0 to my now()
		set bulk to string value of every cell of range "A1:A1000"
		set t1 to my now()
		set end of out to "every-cell 1000 string value: " & (t1 - t0) & " s; count " & (count of bulk) & "; sample " & (item 1000 of bulk)
		set t0 to my now()
		set vals to value of range "B1:B1000"
		set t1 to my now()
		set end of out to "value of 1000-cell range: " & (t1 - t0) & " s; count " & (count of vals)
		set t0 to my now()
		copy range range "A1:A1000"
		set nlines to my sh("pbpaste | wc -l | tr -d ' '")
		set t1 to my now()
		set end of out to "copy range + pbpaste: " & (t1 - t0) & " s; lines " & nlines
		set t0 to my now()
		set bulk2 to string value of every cell of range "A1:A200"
		set t1 to my now()
		set end of out to "every-cell 200 string value: " & (t1 - t0) & " s"
	end tell
	close wb saving no
	set AppleScript's text item delimiters to linefeed
	set t to out as text
	set AppleScript's text item delimiters to ""
	return t
	set display alerts to true
	on error errMsg number errNum
		set display alerts to true
		error errMsg number errNum
	end try
end tell
