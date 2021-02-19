tell application "Microsoft Outlook"
	set job_titles to {"LLM", "Drs.", "MA", "MSc", "BSc", "BA", "Ir.", "Prof.dr.", "Prof.ir.", "Dr.ir.", "Dr.", "Dr. Ph.D.", "Ph.D.", "MSc BA", "BEc", "Drs. BEd", "Mr.", "Mr.drs."}
	repeat with this_contact in contacts
		log company of this_contact as string
		if (company of this_contact = "UU") then
			set company of this_contact to "Universiteit Utrecht"
			set business city of this_contact to "Utrecht"
		else if (company of this_contact = "VU") then
			set company of this_contact to "Vrije Universiteit"
			set business city of this_contact to "Amsterdam"
		else if (company of this_contact = "KB") then
			set company of this_contact to "Koninklijke Bibliotheek"
			set business city of this_contact to "Den Haag"
		else if (company of this_contact = "UvA") then
			set company of this_contact to "Universiteit van Amsterdam"
			set business city of this_contact to "Amsterdam"
		else if (company of this_contact = "RUG") then
			set company of this_contact to "Rijksuniversiteit Groningen"
			set business city of this_contact to "Groningen"
		else if (company of this_contact = "UvT") then
			set company of this_contact to "Universiteit van Tilburg"
			set business city of this_contact to "Tilburg"
		end if
		log job title of this_contact as string
		if (job title of this_contact is in job_titles) then
			set title of this_contact to job title of this_contact
			set job title of this_contact to ""
		end if		
	end repeat
end tell