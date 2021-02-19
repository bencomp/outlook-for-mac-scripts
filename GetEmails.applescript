
try
	set theDate to current date
	set file_name to "/Users/companjenba/Documents/emails-" & short date string of theDate & ".tsv"
	set fp to open for access POSIX file file_name with write permission
	write "Account_name	Time_sent	Name	Email
" to fp
	tell application "Microsoft Outlook"
		
		set acc to first exchange account
		repeat with del_acc in delegated accounts of acc
			--set del_acc to first delegated account of acc
			set acc_name to name of del_acc
			set mail_dir to sent items of del_acc
			repeat with mess in outgoing messages of mail_dir
				repeat with rec in recipients of mess
					set e to email address of rec
					write "" & acc_name & "	" to fp --as text
					write "" & time sent of mess & "	" to fp --as text
					write name of e & "	" to fp --as text
					write address of e & "
" to fp as text
				end repeat
			end repeat
		end repeat
	end tell
	close access fp
on error errStr number errorNumber
	log errStr
	close access fp
end try
--try
--end try