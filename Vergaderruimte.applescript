tell application "Microsoft Outlook"
	set room_emails to {}
	repeat with this_group in groups
		if (name of this_group = "Zalen") then
			repeat with room in expand this_group
				copy address of room to the end of room_emails
			end repeat
		end if
	end repeat
	log room_emails

	repeat with this_contact in contacts
		if (display type of this_contact isn't meeting room) then
			if (job title of this_contact = "kamer") then
				set display type of this_contact to meeting room
			end if
		end if
	end repeat
end tell
