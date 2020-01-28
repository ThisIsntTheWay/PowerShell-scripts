# Bulk change locale of Exchange mailboxes.
# CAUTION: Inapplicable if target user already has folders in target lanugage.
#          In that case, conflicts must be resolved manually.
# (c) 2020, vk

# This will assume a Timezone of CET

# ===========================================================================================
# MAIN

Write-Host "Please specify a locale ID." -ForegroundColor Yellow
Write-Host "(Commonly used: 2055 [de_CH], 1033 [en_US], 4108 [fr_CH])" -ForegroundColor DarkGray
$l = Read-Host "Locale ID: $l"

Write-Host ""
Write-Host "Selected locale ID is: $l"
Write-Host "Please select an action." -ForegroundColor Yellow
Write-Host "1 - Change langauge of EVERY mailbox."
Write-Host "2 - Change langauge of SPECIFIC mailbox."
$c = Read-Host "Choice:"

switch($c) {
  '1' {
    Write-Host "Changing EVERY mailbox..." -ForegroundColor Cyan
    Get-mailbox -ResultSize unlimited | Set-MailboxRegionalConfiguration -Language $l -TimeZone "W. Europe Standard Time" -LocalizeDefaultFolderName
  }
  
  '2' {
    Write-Host "Changing SPECIFIC mailbox..." -ForegroundColor Cyan
    $d = Read-Host "Specify target:"
    Set-MailboxRegionalConfiguration $d -Language $l -TimeZone "W. Europe Standard Time" -LocalizeDefaultFolderName
  } 
}

Write-Host ""
Write-Host "Script has concluded." -ForegroundColor Magenta
