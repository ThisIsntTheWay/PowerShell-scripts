# Generate random strings with numbers and upper-, lower-case characters.
# Can be used for temporary passwords, for example.
# (c) 2020, VK

function Generate-String {
  # $ran1 is used for characters, $ran2 for numbers

  [string]$ran1 = -join ((65..90) + (97..122) | Get-Random -Count 5 | % {[char]$_})
  [string]$ran2 = (10..100) | Get-Random 3
  
  # Separate all (3) generated numbers into variables.
  $a,$b,$c = $ran2.split()
  
  # Construct random string and output to host
  Write-Host "$a$ran1$b$c" -ForegroundColor Yellow
}
