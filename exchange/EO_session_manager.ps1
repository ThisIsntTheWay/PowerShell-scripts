# ---------------------------------------------------------------------------------------------------
# @AUTHOR
#	Simple script to quickly initiate an Exchange Online session in PS.
# 	Built to quickly access creds of commonly used tenants.
# 	Also has some additional, gimmicky stuff.
#
# 	(c) 2019, vk
#
# @USAGE
# 	This script can be run either as is or with the following, optional parameters:
# 	  -Connect <INT>
# 	  -Disconnect
#
# 	When using '-Connect', a number is expected to be passed along.
# 	The value will reference a connection profile as shown in Write-MainMenu.
#
# 	Example:
# 	".\Script.ps1 -Connect 1" will load connection profile USER 1, as it is located at index '1'.
#
# 	When using '-Disconnect', no value is expected to be passed along.
# 	If used, the script will terminate all PSSessions at outlook.office365.com
# ---------------------------------------------------------------------------------------------------

# ---------------------------------------------------------------------------------------------------
# INIT Phase

# Parameters
param (
	[String]$Connection = "NULL",
	[Switch]$Disconnect
)

$scrBuild 	= "1.4"
$scrBuildDate 	= "03.01.2020"

Write-Host "Version: $scrBuild | Build date: $scrBuildDate" -ForegroundColor Gray
Write-Host "© 2020, V. Klopfenstein" -ForegroundColor Gray

# Check if a session at outlook.office365.com exists already
$chk = Get-PSSession | Where-Object {$_.ComputerName -eq "outlook.office365.com"}

# Disconnect session if parameter '-Disconnect' is passed
if ($Disconnect -eq $true) {
	Write-Host "All active sessions at outlook.office365.com will be disconnected."
	Remove-PSSession $chk
	
	if($?) {
		Write-Host ""
		Write-Host "Session(s) disconnected!" -ForegroundColor Green
		$host.ui.RawUI.WindowTitle = "Windows PowerShell"
	   
	} else {
		Write-Host "An error has occurred." -ForegroundColor Red -BackgroundColor Black
		   
		exit
	}	
	exit
}

# Check if parameter, once supplied, is valid
if ($Connection -ne "NULL") {
	if ( !($Connection -match '[1-9]') ) {
		Write-Host ""
		
		Write-Host "Cannot continue:" -ForegroundColor Red -BackgroundColor Black
		Write-Host "Parameter '$Connection' is invalid." -ForegroundColor Red -BackgroundColor Black
		Write-Host "'-Connection' must be a number." -ForegroundColor Red -BackgroundColor Black
		
		exit
	}
}

# Prompt user for choice if a session has been detected.
# Continuity of script is determined based on the value of $i.
# 'y' = terminate, 'n' = continue 
if ($chk.ComputerName -eq "outlook.office365.com") {

	$i = "y"
	
	Write-Host "Cannot continue:" -ForegroundColor Red -BackgroundColor Black
	Write-Host "At least one active PSSession already exists." -ForegroundColor Red -BackgroundColor Black
	
	Write-Host ""
	
	$choice = Read-Host "Disconnect all sessions at outlook.office365.com? (y/n)"
	
	if ($choice -eq "y") {
	
		# Get PSSession and copy $t as $d
		$t = Get-PsSession | Where ComputerName -eq "outlook.office365.com"
		$d = $t
		
		# Disconnect/Remove PSSession
		Remove-PSSession $chk | Out-Null
	   
		if($?) {
		  	# Write all now active sessions into Session variable
			$Session = Get-PSSession
			
			$host.ui.RawUI.WindowTitle = "Windows PowerShell"
			
			Write-Host "Session(s) discconected!" -ForegroundColor Green
			Write-Host ""
			$i = "n"
		   
		} else {
			Write-Host "An error has occurred!" -ForegroundColor Red -BackgroundColor Black
			Write-Host "Try creating a new runspace." -ForegroundColor Red -BackgroundColor Black
			$i = "y"
		}		
	}	
	if ($i -eq "y") { exit }
}

# Prompt main menu
function Write-MainMenu {
	Write-Host ""
	Write-Host "======================================"
	Write-Host "| Please select a credential object: |"
	Write-Host "| 1: OMITTED                         |"
	Write-Host "| 2: OMITTED                         |"
	Write-Host "| 3: OMITTED                         |"
	Write-Host "|                                    |"
	Write-Host "| 9: Manually specified              |"
	Write-Host "| ---------------------------------- |"
	Write-Host "| g: Encrypt SecureString            |"
	Write-Host "| q: Quit                            |"
	Write-Host "======================================"
	Write-Host ""
}

# ---------------------------------------------------------------------------------------------------
# === MAIN

# Implement check to see if user has provided the $Connection param.
# If yes, skip menu segment.

if ($Connection -ne "NULL") {
	# Do nothing (yet?)
} else {
	Write-MainMenu
	Write-Host "Use 'c' to prompt for command, 'r' for printing this menu again." -ForegroundColor Gray
}

# Menu/Param logic
do {
	$break = "n"
	
	# If a parameter has been defined, skip Read-Host
	if ($Connection -ne "NULL") {
		# Copy $connection into $choice, simulating user input
		$choice = $Connection
		
	} else {
		$choice = Read-Host "Input"
	}
	
	switch ($choice) {
		
		# == Tenants ==
		
		'1' {
			# USER 1
			$break = "y"
			
			$choiceName = "USER 1"
			
			$domPWenc = ""
			$domPWstr = ConvertTo-SecureString -String $domPWenc
			
			$objCred = New-Object System.Management.Automation.PSCredential (".onmicrosoft.com", $domPWstr)
		}
		'2' {
			# USER 2
			$break = "y"
			
			$choiceName = "USER 2"
			
			$domPWenc = ""
			$domPWstr = ConvertTo-SecureString -String $domPWenc
			
			$objCred = New-Object System.Management.Automation.PSCredential (".onmicrosoft.com", $domPWstr)		
		}
		'3' {
			# USER 3
			$break = "y"
			
			$choiceName = "USER 3"
			
			$domPWenc = ""
			$domPWstr = ConvertTo-SecureString -String $domPWenc
			
			$objCred = New-Object System.Management.Automation.PSCredential (".onmicrosoft.com", $domPWstr)		
		}
		'9' {
			# Other / User specified
			$break = "y"
			
			$choiceName = "manually specified"
			
			$objCred = Get-Credential "admin@tenant.onmicrosoft.com"
		}
		
		# == Miscellaneous commands ==
		
		'c' {
			# Prompt command used for establishing O365 session
			$break = "n"
			$noErr = "y"
			
			Write-Host "The command used for establishing an O365 session in PS is:"
			Write-Host "> $([char]0x0024)CREDENTIALS = Read-Credential" -Foregroundcolor Gray
			Write-Host "> $([char]0x0024)SESSION = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $([char]0x0024)CREDENTIALS -Authentication Basic -AllowRedirection" -Foregroundcolor Gray
			Write-Host "> Import-PSSession $([char]0x0024)SESSION -DisableNameChecking" -Foregroundcolor Gray
			Write-Host ""
		}
		'g' {
			# Generate and encrypt SecureString
			$break = "n"
			$noErr = "y"
			
			Write-Host ""
			Write-Host "Convert input to SecureString and encrypt." -ForegroundColor Cyan
			$sString = Read-Host "String (obfuscated)" -AsSecureString
			$eString = ConvertFrom-SecureString -SecureString $sString

			Write-Host ""
			Write-Host "Encrypted string:" -Foregroundcolor Cyan
			Write-Host "$eString" -Foregroundcolor Yellow
			Write-Host ""
		}
		'r' {
			# Prompt for menu again
			$break = "n"
			$noErr = "y"

			Write-MainMenu
		}
		
		# == Terminating commands ==
		
		'q' {
			# Quit
			$break = "y"
			
			Write-Host ""
			Write-Host "Goodbye!" -Foregroundcolor Magenta
			Write-Host ""
			
			exit
		}
	}
	# Handling of unrecognized input
	if ($break -ne 'y') {		
		if ($noErr -ne "y") {
			
			Write-Host "Invalid input." -ForegroundColor Red -BackgroundColor Black
			Write-Host ""
			
			# Check if a parameter was provided.
			# If yes, exit to prevent a loop.
			if ($connection -ne "NULL") {				
				exit
			}
		}
		$noErr = "n"
	}
}
until ($break -eq 'y')

# ---------------------------------------------------------------------------------------------------
# === EXEC

Write-Host ""
Write-Host "Importing session as $choiceName..." -ForegroundColor Cyan

# Import PSSession
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $objCred -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

# Error handling
if($?) {
   Write-Host "Ready!" -ForegroundColor Green
   $host.ui.RawUI.WindowTitle = "Active O365 session: $choiceName"
} else {
   Write-Host "An error has occurred!" -ForegroundColor Red -BackgroundColor Black
   Write-Host "Script cannot continue." -ForegroundColor Red -BackgroundColor Black
   Write-Host ""
   Write-Host "Last choice: $choice ($choiceName)" -ForegroundColor Yellow
}
