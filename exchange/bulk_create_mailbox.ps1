#(c) 2020, v. klopfenstein

Enable-Mailbox -Identity

$import = import-csv "C:\PATH\TO\FILE.csv" -Delimiter ";"

foreach ($usr in $import) {
	$Alias = $usr.username
	$fname = $usr.firstname
	$lname = $usr.lastname
	
	Enable-Mailbox -identity $Alias -Displayname "$fname $lname"
}
