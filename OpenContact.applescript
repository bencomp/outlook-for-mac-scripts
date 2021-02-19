tell application "Microsoft Outlook"
	set locations to {""}
	repeat with this_message in meeting messages
		set this_message to first meeting message
		(* open contact email address sender of this_message *)
		set the_event to the meeting of this_message
		set the_location to the location of the_event
		set end of locations to the_location
	end repeat
	--set this_sender to sender of this_message
	--set this_add to address of this_sender
	--open contact email address this_add
end tell
