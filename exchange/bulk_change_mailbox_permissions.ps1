# Bulk change mailbox permissions.
# This example focuses on the calendar.
# (c) 2020, vk
$import = import-csv "C:\tmp\Users.csv" -Delimiter ";"

# Permission level
$perm = "Author"

# Target user/group
$tUsr = "Default"

foreach ($usr in $import) {
	$Alias = $usr.username
	$fname = $usr.firstname
	$lname = $usr.lastname

    $finitial = $fname.SubString(0,1)

   	# Construct Mail adress
	$mail = "$finitial$lname"
	
    # Change calendar permission	
	Set-MailboxFolderPermission -Identity $mail@CHANGEME:\Kalender -User $tUsr -AccessRights $perm
	Set-MailboxFolderPermission -Identity $mail@CHANGEME:\Calendar -User $tUsr -AccessRights $perm
       
    write-Host "Set calendar perm level $perm for $fname $lname."
}
