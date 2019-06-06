set locationsList to {}
tell application "Microsoft Outlook"
	set this_account to default account
	set this_inbox to inbox
	set this_calendar to default calendar of this_account
	
	set locationsListRef to a reference to locationsList
	
	
	repeat with this_event in every calendar event of this_calendar
		--log every properties of this_event
		if (exists location of this_event) then
			try
				
				set this_location to location of this_event
				copy this_location to the end of locationsListRef
				log this_location
			on error number -1728
				log "no location"
				-- statements to write to the file
			end try
		end if
	end repeat
	
	
	--return number of calendar events of this_calendar
	
end tell

try
	set locationsFile to POSIX file "/Users/companjenba/locations.txt"
	set fp to open for access locationsFile with write permission
	set AppleScript's text item delimiters to {linefeed}
	write (locationsList as text) to fp as text
	close access fp
on error
	
	close access fp
	
end try

