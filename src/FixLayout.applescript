-- mac-word-layout-fix (AppleScript version)
-- Word for Mac reads unitless HTML table sizes (<td height="36">, cellpadding="3")
-- as points; Word for Windows reads them as 96-dpi pixels (x0.75).
-- This rescales row heights and cell padding in the front Word document.
--
-- Installed to ~/Library/Application Scripts/com.microsoft.Word/ so that it shows in
-- Word's Scripts menu. When the add-in is installed it fixes documents as they open, and
-- this menu item only confirms that; otherwise it runs the (slower) AppleScript fix.

on scalePadding(t)
	tell application "Microsoft Word"
		try
			set top padding of t to (top padding of t) * 0.75
			set bottom padding of t to (bottom padding of t) * 0.75
			set left padding of t to (left padding of t) * 0.75
			set right padding of t to (right padding of t) * 0.75
		end try
	end tell
end scalePadding

on fixTable(t)
	tell application "Microsoft Word"
		my scalePadding(t)
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

-- Slow fallback (one Apple Event per row and cell). Returns "fixed" or "skipped".
on fixActiveDocument(mode)
	tell application "Microsoft Word"
		if (count of documents) = 0 then return "skipped"
		set doc to active document
		if (save format of doc) is not format HTML then return "skipped"
		set wasSaved to saved of doc
		repeat with n from 1 to (count of tables of doc)
			my fixTable(table n of doc)
		end repeat
		-- Only the on-screen layout changed; don't prompt to save on close.
		-- A just-opened document has no user edits yet, so it is safe to mark it saved.
		if mode is not "manual" or wasSaved then set saved of doc to true
	end tell
	return "fixed"
end fixActiveDocument

-- Scripts menu entry point.
on run
	tell application "Microsoft Word"
		if (count of documents) = 0 then return
		set addinLoaded to false
		try
			set addinLoaded to installed of add in "MacWordLayoutFix.dotm"
		end try
		-- The add-in already fixed this document when it opened; running again would shrink it twice.
		if addinLoaded and (save format of active document) is format HTML then
			display dialog "The mac-word-layout-fix add-in already fixed this document when it opened." buttons {"OK"} default button 1
			return
		end if
	end tell
	if fixActiveDocument("manual") is "skipped" then
		tell application "Microsoft Word"
			display dialog "This document is not an HTML-based file, so no fix is needed." buttons {"OK"} default button 1
		end tell
	end if
end run
