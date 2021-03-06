# Change Windows wallpaper based on time of day
# (c) 2020, vk

##################################
# TODO LIST
# [X] Implement 24h format
# [ ] Test
##################################

# ******************************

[string]$wpPathRoot = "C:\Path\to\wallpaper\folder\"
[string]$wpFileSuffix = ".jpg"
[string]$wpFile = "NULL"

[string]$time = "NULL"

[bool]$wpFileValid = $false

$curTimeH = (get-date).ToString('HH') -as [int]

function Set-Wallpaper([string] $file) {

	$wpFilePath = "$wpPathRoot\$file"

	if (Test-Path $wpFilePath) {
		$wpFileValid = $true
	}

	if ($wpFileValid = $true) {
		New-ItemProperty "HKCU:\Control Panel\Desktop" -Name "Wallpaper" -Type "REG_SZ" -Value "$wpPathRoot\$file" -Force
		Start-Sleep -s 10
		
		# Set Wallpaper as defined in in Registry
		Start-Process -FilePath "rundll32.exe" user32.dll -ArgumentList "UpdatePerUserSystemParameters, 0, $false"
		if (!($?)) {
			Write-Host "An error occurred during rundll32.exe call" -ForegroundColor Red -BackgroundColor Black
		}
	} else {
		Write-Host "Could not set wallpaper:" -ForegroundColor Red -BackgroundColor Black
		Write-Host "File $wpFile not found in $wpPathRoot." -ForegroundColor Red -BackgroundColor Black
	}
}

# ******************************
# Main routine

# Time checks
if ($curTimeH -ge 21) {
	$time = "Night"
	
	$wpFile = "night.jpg" + $wpFileSuffix
	Set-Wallpaper($wpFile)
	
} elseif ($curTimeH -ge 17) {
	$time = "Evening"
	
	$wpFile = "evening" + $wpFileSuffix
	Set-Wallpaper($wpFile)
	
} elseif ($curTimeH -ge 13) {
	$time = "Afternoon"
	
	$wpFile = "afternoon" + $wpFileSuffix
	Set-Wallpaper($wpFile)
	
} elseif ($curTimeH = 12) {
	$time = "Midday"
	
	$wpFile = "midday" + $wpFileSuffix
	Set-Wallpaper($wpFile)
	
} elseif ($curTimeH -ge 10) {
	$time = "Noon"
	
	$wpFile = "noon" + $wpFileSuffix
	Set-Wallpaper($wpFile)
	
} elseif ($curTimeH -ge 9) {
	$time = "Morning"
	
	$wpFile = "morning" + $wpFileSuffix
	Set-Wallpaper($wpFile)
	
} elseif ($curTimeH -ge 7) {
	$time = "Sunrise"
	
	$wpFile = "sunrise" + $wpFileSuffix
	Set-Wallpaper($wpFile)
	
}

if ($time = "NULL") {
	throw "Current time $curTimeH did not match any case!"
}

Write-Host "Current hour (24h): $curTimeH / $time"
Write-Host "Wallpaper set to $wpFile"
