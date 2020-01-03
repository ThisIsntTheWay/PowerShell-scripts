# (c) 2020, v. klopfenstein
#
# Gather inventory of computer that this script is running on.
# CAUTION: Only partially works on Win7 due to Win32 object designation discrepancies.

Import-Module ActiveDirectory
if (test-path -Path "<path>$name-sw_info.csv") {
	Write-Host "$name has already been processed. Skipping."
} else {
	# Acquire architecture, install date, build version, operating system, hostname
	$arch = (Get-WmiObject Win32_OperatingSystem).OSArchitecture
	$date = gcim Win32_OperatingSystem | select InstallDate #using GCIM for date formatting
	$winb = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId #BUILD
	$winn = (Get-WmiObject -class Win32_OperatingSystem).Caption
	$name = hostname

	$o = new-object PSObject
	$o | add-member NoteProperty Hostname $name
	$o | add-member NoteProperty winCaption $winn
	$o | add-member NoteProperty winBuild $winb
	$o | add-member NoteProperty Architecture $arch
	$o | add-member NoteProperty InstallDate $date

	Write-Host "Creating .csv"

	$o | export-csv "<path>$name-sw_info.csv" -notypeinformation

	Write-Host "Done!"
}
