tell application "Microsoft Outlook"
	
	repeat with this_contact in contacts
		if (display type of this_contact = meeting room) then
			log "meeting room"
			log vcard data of this_contact as text
		end if
	end repeat
end tell