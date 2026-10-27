-- Probe: does `string value` reflect column width; bulk-read forms; autofit measurement;
-- number-format color via display format; dialog-equivalent rejection of lower-case e+.
-- Creates one new workbook, never saves, closes it. Non-interactive.
on j(lst)
	set AppleScript's text item delimiters to " | "
	set t to lst as text
	set AppleScript's text item delimiters to ""
	return t
end j

tell application "Microsoft Excel"
	set display alerts to false
	try
	set out to {}
	set end of out to "excel version: " & (version as text)
	set end of out to "standard font: " & (standard font as text) & " " & (standard font size as text)
	set wb to make new workbook
	tell active sheet of wb
		set end of out to "cell font: " & (name of font object of range "A1") & " " & (font size of font object of range "A1")
		set value of range "A1" to -1
		set number format of range "A1" to "m/d/yy"
		set value of range "A2" to 1234567.911
		set number format of range "A2" to "General"
		set value of range "A3" to 1234.5
		set number format of range "A3" to "0.00*-"
		set value of range "A4" to 123456789
		set number format of range "A4" to "0.00"
		set value of range "A5" to "hello world text"
		set number format of range "A5" to "@"
		set value of range "A6" to 0.5
		set number format of range "A6" to "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"
		set value of range "A7" to 123456789012
		set number format of range "A7" to "General"
		set value of range "A8" to -1234.5
		set number format of range "A8" to "[Red]0.00;[Blue]-0.00"
		set column width of range "A:A" to 6
		set end of out to "A column width 6 -> width pts " & (width of range "A1" as text)
		set narrow to {}
		repeat with i from 1 to 8
			set end of narrow to (string value of range ("A" & i))
		end repeat
		set end of out to "narrow per-cell: " & my j(narrow)
		set end of out to "narrow every-cell: " & my j(string value of every cell of range "A1:A8")
		set column width of range "A:A" to 30
		set end of out to "A column width 30 -> width pts " & (width of range "A1" as text)
		set wide to {}
		repeat with i from 1 to 8
			set end of wide to (string value of range ("A" & i))
		end repeat
		set end of out to "wide per-cell: " & my j(wide)
		set end of out to "wide every-cell: " & my j(string value of every cell of range "A1:A8")
		set end of out to "wide A1 length: " & (length of (string value of range "A1"))
		try
			set m to string value of range "A1:A8"
			set end of out to "multi-cell string value class: " & (class of m as text) & " -> " & my j(m)
		on error e
			set end of out to "multi-cell string value error: " & e
		end try
		try
			set end of out to "A8 display format font color: " & my j(color of font object of display format of range "A8")
		on error e
			set end of out to "A8 display format font color error: " & e
		end try
		try
			set end of out to "A8 font color: " & my j(color of font object of range "A8")
		on error e
			set end of out to "A8 font color error: " & e
		end try
		set end of out to "A8 horizontal alignment: " & (horizontal alignment of range "A8" as text) & " / A5: " & (horizontal alignment of range "A5" as text)
		-- autofit measurement: one text per column, autofit, read widths back
		set texts to {"Tuesday, January 4, 2000", "Sunday, December 3, 2000", "1234567890", "0000000000", "iiiiiiiiii", "WWWWWWWWWW", "0", "00000000000000000000"}
		set cols to {"C", "D", "E", "F", "G", "H", "I", "J"}
		repeat with k from 1 to 8
			set value of range ((item k of cols) & "1") to (item k of texts)
			set number format of range ((item k of cols) & "1") to "@"
		end repeat
		try
			autofit (entire column of range "C1:J1")
			set end of out to "autofit form: entire column of range"
		on error e1
			try
				repeat with k from 3 to 10
					autofit column k
				end repeat
				set end of out to "autofit form: column k (entire-column form failed: " & e1 & ")"
			on error e2
				set end of out to "autofit failed: " & e1 & " // " & e2
			end try
		end try
		repeat with k from 1 to 8
			set end of out to "autofit " & (item k of texts) & " -> column width " & (column width of range ((item k of cols) & "1") as text) & ", width pts " & (width of range ((item k of cols) & "1") as text)
		end repeat
		set value of range "K1" to 1234567890
		set number format of range "K1" to "0"
		try
			autofit (entire column of range "K1")
		on error
			autofit column 11
		end try
		set end of out to "autofit number 1234567890 -> column width " & (column width of range "K1" as text) & ", width pts " & (width of range "K1" as text)
		set column width of range "L:L" to 8.43
		set end of out to "column width 8.43 -> width pts " & (width of range "L1" as text)
		set column width of range "L:L" to 1
		set end of out to "column width 1 -> width pts " & (width of range "L1" as text)
		set column width of range "L:L" to 10
		set end of out to "column width 10 -> width pts " & (width of range "L1" as text)
		-- dialog-equivalent check for lower-case exponent
		try
			set number format of range "B1" to "0.00e+00"
			set end of out to "set number format 0.00e+00 accepted; read back: " & (number format of range "B1")
		on error e
			set end of out to "set number format 0.00e+00 error: " & e
		end try
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
