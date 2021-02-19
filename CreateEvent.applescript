tell application "Microsoft Outlook"
	set acc to first exchange account
	set my_event to make new calendar event with properties {start time:date "25-11-20, 15:00", end time: date "25-11-20, 16:00", subject:"SWIB"}
	--set start time of my_event to (current date) + 1 * days
	open my_event
	log short date string of start time of my_event
end tell