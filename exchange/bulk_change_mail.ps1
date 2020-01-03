# (c) 2020, v. klopfenstein

$import = import-csv "C:\PATH\TO\FILE.csv" -Delimiter ";"

foreach ($usr in $import) {
	$Alias = $usr.username
	$fname = $usr.firstname
	$lname = $usr.lastname
	
	# Get first character of fname
	$finitial = $fname.SubString(0,1)
	
	# Construct Mail adress
	$mail = "$finitial$lname"
	
	Set-Mailbox -identity $Alias -EmailAddresses "SMTP:$mail@example.com" -EmailAddressPolicyEnabled $false
}
