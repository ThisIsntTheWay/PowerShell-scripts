# (c) 2020, v. klopfenstein
# Get last logon date of specified computer objects.
# Export result into specified CSV.

import-module activedirectory
Get-Content ".\computers.txt" | Foreach-Object {Get-ADComputer -identity $_ -Properties Name,LastLogonDate | select-object -property Name,LastLogonDate } | Export-Csv -Notype "computers.csv"
