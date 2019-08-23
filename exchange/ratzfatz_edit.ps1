# ---------------------------------------------
# Author: Valentin Klopfenstein
# Made in August 2019
#
# Use at will without credit
# ---------------------------------------------

# Set up inbound params
param(
	[string] $enableLog
)

# Constant variables
$globalPath = Get-Location

# Define main menu
function Main-Menu {
    param (
        [string]$Title = 'Exchange Ratzfatz edit'
    )
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "[!] Only use with v2010 or higher!"
    Write-Host ""
    
    Write-Host "1: Assign Mailbox permissions"
    Write-Host "2: Export Mailbox"
    Write-Host "3: Get all Memmbers of the global OAB"
    Write-Host "Q: Abort"
	
	if ($enableLog = "ToLog") {
		Write-Host ''
		Write-Host '[i] Logging is active!'
		Write-Host "[i] Logs are saved in $globalPath"
	}
}

# Assign Mailbox permission
function assignPerm {
        Write-Host ''
        Write-Host '================================='
        Write-Host '> Assign Mailbox permissions'
        Write-Host ''
		
        $ident = Read-Host "Define 'identity'"
        $user = Read-Host "Define 'user'"
		
        Write-Host ''
		
        Write-Host "[i] For reference on possible AccessRights params, consult the following:"
        Write-Host '[>] https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/set-mailboxfolderpermission?view=exchange-ps#parameters'
        $acr = Read-Host "Set Access level"

        Add-MailboxPermissions -Identity "$ident" -User "$user" -AccessRights "$acr"

        Write-Host "Done! Press RETURN to proceed"
        pause >NUL

        main
}

# Export Mailbox
function exportMail {
    Write-Host ''
    Write-Host '[i] Caution! Only possible in on-premises configuration!'
    Write-Host ''

    $target = Read-Host "Target Mailbox"
    $path = Read-Host "Path (Must be UNC)"

    New-MailboxExportRequest -Mailbox "$target" -FilePath "$path"
    Get-MailboxExportRequest -Status InProgress

    Write-Host "Done! Press RETURN to proceed"
    pause >NUL

    main
}

# Acquire global offline address book
function getGOAB {
	Write-Host ''
	
	$a = get-addresslist "Offline Global Address List"
	get-recipient -filter $a.recipientfilter
	
    Write-Host "Done! Press RETURN to proceed"
    pause >NUL

    main
}

# Actual program
function main {
cls
    do
     {
        Main-Menu

        Write-Host ""

        $selection = Read-Host "Awaiting input"
        switch ($selection)
        {
            '1' {

                assignPerm

            } '2' {

                exportMail

            } '3' {
              'NOT YET COMPLETE'

            } default {

                'Invalid input!'

            } 'Q' {
            Write-Host 'Good day!'
            exit }
        }
        
        Write-Host 'Press RETURN to proceed.'
        pause >NUL

     }
     until ($selection -eq 'q')
 }

 # Actual program
 main
