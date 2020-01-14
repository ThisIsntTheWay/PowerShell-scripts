$import = import-csv "C:\tmp\Users.csv" -Delimiter ";"

$domain = "CHANGEME"

foreach ($usr in $import) {
	$Alias = $usr.username
	$fname = $usr.firstname
	$lname = $usr.lastname
	
	# Get first character of fname
	$finitial = $fname.SubString(0,1)
	
	# Construct Mail adress
	$mail = "$finitial$lname"
	
	Set-Mailbox -identity $Alias -EmailAddresses "SMTP:$mail@$domain" -EmailAddressPolicyEnabled $false
	
	write-Host "Mail set for $fname $lname -> $mail@$domain"
}
