tell application "Microsoft Outlook"
	repeat with this_contact in contacts
		if (company of this_contact = "UU") then
			set company of this_contact to "Universiteit Utrecht"
		end if
		if (job title of this_contact = "Drs.") then
			set job title of this_contact to ""
			set title of this_contact to "Drs."
		else if (job title of this_contact = "MA") then
			set job title of this_contact to ""
			set title of this_contact to "MA"
		else if (job title of this_contact = "MSc") then
			set job title of this_contact to ""
			set title of this_contact to "MSc"
		else if (job title of this_contact = "BA") then
			set job title of this_contact to ""
			set title of this_contact to "BA"
		else if (job title of this_contact = "Dr.") then
			set job title of this_contact to ""
			set title of this_contact to "Dr."
		end if
	end repeat
end tell