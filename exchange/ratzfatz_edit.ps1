# Function definition

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
    Write-Host "3: Get all Memmbers of the OGAL"
    Write-Host "Q: Abort"
}

# Assign Mailbox permission
function assignPerm {
        ''
        '================================='
        '> Assign Mailbox permissions'
        ''
        $ident = Read-Host "Define 'identity'"
        $user = Read-Host "Define 'user'"
        ''
        "[i] For reference on possible AccessRights params, consult the following:"
        '[>] https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/set-mailboxfolderpermission?view=exchange-ps#parameters'
        $acr = Read-Host "Set Access level"

        Add-MailboxPermissions -Identity "$ident" -User "$user" -AccessRights "$acr"

        Write-Host "Done! Press RETURN to proceed"
        pause >NUL

        main
}

# Export Mailbox
function exportMail {
    ''
    '[i] Caution! Only possible in on-premises configuration!'
    ''
    $target = Read-Host "Target Mailbox"
    $path = Read-Host "Path (Must be UNC)"

    New-MailboxExportRequest -Mailbox "$target" -FilePath "$path"
    Get-MailboxExportRequest -Status InProgress

    Write-Host "Done! Press RETURN to proceed"
    pause >NUL

    main
}

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
          'You chose option #3'
        }
        }
        pause
     }
     until ($selection -eq 'q')
 }

 # Actual program
 main
