# -------------------------------------------------------------------------
# Simple script to quickly initiate an Exchange Online session in PS.
# Built to quickly access creds of commonly used tenants.
# Also has some additional, gimmicky stuff.
#
# (c) 2019, vk
# -------------------------------------------------------------------------

$scrBuild 		= "1.2"
$scrBuildDate 	= "31.12.2019"

# Check if a session exists already

$chk = Get-PSSession | Where-Object {$_.ComputerName -eq "outlook.office365.com"}

if ($chk.ComputerName -eq "outlook.office365.com") {

	$i = "y"
	
	Write-Host "Cannot continue!" -ForegroundColor Red
	Write-Host "A session already exists." -ForegroundColor Red
	
	Write-Host ""
	
	$choice = Read-Host "Disconnect all sessions at outlook.office365.com? (y/n)"
	
	if ($choice -eq "y") {
	
		# Get PSSession and copy $t as $d
		$t = Get-PsSession | Where ComputerName -eq "outlook.office365.com"
		$d = $t
		
		# Remove PSSession
		# The following commands achieve the same thing, but are present for redundancy
		Remove-PSSession $chk | Out-Null
	   
		if($?) {
		  	# Write all now active sessions into Session variable
			$Session = Get-PSSession
			
			$host.ui.RawUI.WindowTitle = "Windows PowerShell"
			
			Write-Host "Session(s) discconected!" -ForegroundColor Green
			Write-Host ""
			$i = "n"
		   
		} else {
			Write-Host "An error has occurred!" -ForegroundColor Red
			Write-Host "Try creating a new runspace." -ForegroundColor Red
		   
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
	Write-Host "| 1: USER 1                          |"
	Write-Host "| 2: USER 2                          |"
	Write-Host "| 3: USER 3                          |"
	Write-Host "|                                    |"
	Write-Host "| 9: Manually specified              |"
	Write-Host "| ---------------------------------- |"
	Write-Host "| g: Encrypt SecureString            |"
	Write-Host "| q: Quit                            |"
	Write-Host "======================================"
	Write-Host ""
}

# ---------------------------------------------------------------------------------------------------

Write-Host "Version: $scrBuild | Build date: $scrBuildDate" -ForegroundColor Gray
Write-Host "© 2020, V. Klopfenstein - sdag" -ForegroundColor Gray
Write-MainMenu
Write-Host "Use 'c' to prompt for command, 'r' for printing this menu again." -ForegroundColor Gray

# Menu logic
do {
	$break = "n"
	$choice = Read-Host "Input"
	
	switch ($choice) {
		
		# == Tenants ==
		
		'1' {
			# USER 1
			$break = "y"
			
			$choiceName = "<OMITTED>"
			
			$domPWenc = "<OMITTED>"
			$domPWstr = ConvertTo-SecureString -String $domPWenc
			
			$objCred = New-Object System.Management.Automation.PSCredential ("<OMITTED>", $domPWstr)
		}
		'2' {
			# USER 2
			$break = "y"
			
			$choiceName = "<OMITTED>"
			
			$domPWenc = "<OMITTED>"
			$domPWstr = ConvertTo-SecureString -String $domPWenc
			
			$objCred = New-Object System.Management.Automation.PSCredential ("<OMITTED>", $domPWstr)		
		}
		'3' {
			# USER 3
			$break = "y"
			
			$choiceName = "<OMITTED>"
			
			$domPWenc = "<OMITTED>"
			$domPWstr = ConvertTo-SecureString -String $domPWenc
			
			$objCred = New-Object System.Management.Automation.PSCredential ("<OMITTED>", $domPWstr)		
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
			$Secure = Read-Host "String (obfuscated)" -AsSecureString
			$Encrypted = ConvertFrom-SecureString -SecureString $Secure

			Write-Host ""
			Write-Host "Encrypted string:" -Foregroundcolor Cyan
			Write-Host "$Encrypted" -Foregroundcolor Yellow
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
			Write-Host "Invalid input." -ForegroundColor Red
			Write-Host ""
		}
		
		$noErr = "n"
	}
}
until ($break -eq 'y')

# ---------------------------------------------------------------------------------------------------

# Execution

Write-Host ""
Write-Host "Importing session as $choiceName..." -ForegroundColor Cyan

# Import PSSession based on previous input
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $objCred -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

# Error handling
if($?) {
   Write-Host "Ready!" -ForegroundColor Green
   $host.ui.RawUI.WindowTitle = "Active O365 session: $choiceName"
   
} else {
   Write-Host "An error has occurred." -ForegroundColor Red
   Write-Host "Please consult last exception." -ForegroundColor Red
   Write-Host ""
   Write-Host "Last input: $choice ($choiceName)" -ForegroundColor Yellow
   
}
