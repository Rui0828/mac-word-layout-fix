-- mac-word-layout-fix (AppleScript version)
-- Rescales HTML table row heights in the front Word document (px -> pt, x0.75).
-- Install to ~/Library/Application Scripts/com.microsoft.Word/ to show it in Word's Scripts menu.

on fixTable(t)
	tell application "Microsoft Word"
		repeat with i from 1 to (count of rows of t)
			set r to row i of t
			try
				set h to height of r
				if h is not missing value and h > 0 then set height of r to (h * 0.75)
			end try
			repeat with j from 1 to (count of cells of r)
				set c to cell j of r
				set k to 1
				repeat
					try
						set nt to table k of c
					on error
						exit repeat
					end try
					my fixTable(nt)
					set k to k + 1
				end repeat
			end repeat
		end repeat
	end tell
end fixTable

tell application "Microsoft Word"
	if (count of documents) = 0 then return
	set doc to active document
	if (save format of doc) is not format HTML then
		display dialog "This document is not an HTML-based file, so no fix is needed." buttons {"OK"} default button 1
		return
	end if
	repeat with n from 1 to (count of tables of doc)
		my fixTable(table n of doc)
	end repeat
end tell
