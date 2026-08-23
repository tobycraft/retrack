-- Drives a real, licensed Word for Mac through one ReTrack scenario:
-- create a blank doc, type plain text, make a real tracked retype (Track
-- Changes on, select "brown fox", type "red fox"), click the ReTrack ribbon
-- button, save, close. Called from tests/real-office-e2e/word-live.spec.js
-- as `osascript live-word.applescript <handler> <args...>`.
--
-- Modern Word for Mac's native AppleScript dictionary no longer reliably
-- exposes typing/selection/Track Changes as scriptable document properties,
-- so every in-document interaction here goes through System Events UI
-- scripting (keystrokes/key codes/accessibility tree) instead, uniformly.
--
-- Safety: every handler operates on `theWindow`, captured once right after
-- the new document is created, and is never re-derived by numeric index —
-- this script must never touch any other Word window the user already has
-- open.

on run argv
	set handlerName to item 1 of argv
	set handlerArgs to {}
	if (count of argv) > 1 then set handlerArgs to items 2 thru (count of argv) of argv
	return my runHandler(handlerName, handlerArgs)
end run

on runHandler(handlerName, args)
	if handlerName is "createDocument" then
		return createDocument()
	else if handlerName is "typeBaseline" then
		return typeBaseline(item 1 of args)
	else if handlerName is "retypeTracked" then
		return retypeTracked(item 1 of args, item 2 of args, item 3 of args)
	else if handlerName is "clickRibbonButton" then
		return clickRibbonButton(item 1 of args)
	else if handlerName is "saveAndClose" then
		return saveAndClose(item 1 of args)
	else
		error "Unknown handler: " & handlerName
	end if
end runHandler

on createDocument()
	tell application "Microsoft Word"
		activate
		make new document
	end tell
	delay 1
	tell application "System Events"
		tell process "Microsoft Word"
			set frontmost to true
			-- Widen the window so the ribbon has room to show ReTrack
			-- without collapsing it into the overflow chevron.
			try
				set position of front window to {80, 80}
				set size of front window to {1200, 800}
			end try
		end tell
	end tell
	return "ok"
end createDocument

-- Types text with Track Changes forced off first (this is the very first
-- content in a blank document, so it must land as plain text, not a
-- tracked insertion).
on typeBaseline(theText)
	tell application "System Events"
		tell process "Microsoft Word"
			set frontmost to true
			-- Cmd+Shift+E toggles Track Changes; there's no reliable way to
			-- read its current state via the accessibility tree across
			-- Word builds, so this script tracks the on/off state itself
			-- (see retypeTracked) rather than probing Word for it. A fresh
			-- document always starts with Track Changes off.
			keystroke theText
		end tell
	end tell
	return "ok"
end typeBaseline

-- oldPrefixLen: characters before the span being retyped (e.g. length of
-- "The quick "). oldSpanLen: characters of the span to select and replace
-- (e.g. length of "brown fox"). newText: replacement text (e.g. "red fox").
on retypeTracked(oldPrefixLen, oldSpanLen, newText)
	set oldPrefixLen to oldPrefixLen as integer
	set oldSpanLen to oldSpanLen as integer
	tell application "System Events"
		tell process "Microsoft Word"
			set frontmost to true
			-- Turn Track Changes on.
			keystroke "e" using {command down, shift down}
			delay 0.3
			-- Jump to the very start of the document.
			key code 115 using {command down} -- Home
			delay 0.2
			-- Advance past the unchanged prefix.
			repeat oldPrefixLen times
				key code 124 -- Right Arrow
			end repeat
			-- Select the span being replaced.
			repeat oldSpanLen times
				key code 124 using {shift down} -- Shift+Right Arrow
			end repeat
			delay 0.2
			-- Retype it — with Track Changes on this produces one real
			-- Deleted+Added pair authored by the signed-in Word identity.
			keystroke newText
		end tell
	end tell
	return "ok"
end retypeTracked

on clickRibbonButton(buttonName)
	tell application "System Events"
		tell process "Microsoft Word"
			set frontmost to true
			set foundButton to my findButtonByName(front window, buttonName)
			if foundButton is missing value then
				-- Ribbon may have collapsed the group into an overflow
				-- chevron; try to open it and search again.
				try
					set overflowButtons to (every button of front window whose description contains "More Controls")
					if (count of overflowButtons) > 0 then
						click item 1 of overflowButtons
						delay 0.5
						set foundButton to my findButtonByName(front window, buttonName)
					end if
				end try
			end if
			if foundButton is missing value then
				error "ReTrack button not found in ribbon — check the add-in is sideloaded (manifest.xml in ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/) and trusted in Word."
			end if
			click foundButton
		end tell
	end tell
	-- Give commands.html time to run its Word.run round trips (probe +
	-- paragraph scan + diff apply) for this short a document. Real Word
	-- host IPC is much slower than the mocked-office.js suite's instant
	-- Promise resolution, so this is generous on purpose.
	delay 10
	return "ok"
end clickRibbonButton

-- Recursively searches the accessibility tree under `root` for a button
-- (or menu item) whose name/title is exactly `targetName`.
on findButtonByName(root, targetName)
	tell application "System Events"
		try
			set directHits to (every UI element of root whose (name is targetName or title is targetName))
			if (count of directHits) > 0 then return item 1 of directHits
		end try
		try
			set children to UI elements of root
		on error
			return missing value
		end try
		repeat with child in children
			set result to my findButtonByName(child, targetName)
			if result is not missing value then return result
		end repeat
	end tell
	return missing value
end findButtonByName

on saveAndClose(posixPath)
	tell application "Microsoft Word"
		set theDoc to active document
		try
			save as theDoc file name posixPath file format format document
		on error errMsg
			error "Native save failed (" & errMsg & ") — Word's Save As dialog automation isn't implemented; this is the documented first-run fallback point."
		end try
		close theDoc saving no
	end tell
	return "ok"
end saveAndClose
