<# 
.SYNOPSIS
Automates the cleanup of browser cache and refreshes network settings for a smoother and more optimized system experience

.NOTES
Author: Damon Sih Boon Kiat
#>

Write-Output "Starting browser process termination..."

# Task kill for Browsers
Write-Output "Stopping all browser processes..."
Stop-Process -Name "msedge", "chrome", "firefox" -Force -ErrorAction SilentlyContinue
Write-Output "All browser processes stopped successfully."

# Clear cache from Edge
Write-Output "Checking Edge cache folder..."
$EdgeCachePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache\"
if (Test-Path $EdgeCachePath) {
    Write-Output "Edge cache found. Clearing cache..."
    Remove-Item -Path "$EdgeCachePath\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Output "Edge cache cleared successfully."
} else {
    Write-Output "Edge cache folder not found, skipping."
}

# Clear cache from Chrome
Write-Output "Checking Chrome cache folder..."
$ChromeCachePath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache\"
if (Test-Path $ChromeCachePath) {
    Write-Output "Chrome cache found. Clearing cache..."
    Remove-Item -Path "$ChromeCachePath\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Output "Chrome cache cleared successfully."
} else {
    Write-Output "Chrome cache folder not found, skipping."
}

# Clear cache from Firefox
Write-Output "Checking Firefox cache folder..."
$FirefoxCachePath = "$env:APPDATA\Mozilla\Firefox\Profiles\"
if (Test-Path $FirefoxCachePath) {
    Write-Output "Firefox cache found. Clearing cache..."
    Get-ChildItem "$FirefoxCachePath\*\cache2\" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Write-Output "Firefox cache cleared successfully."
} else {
    Write-Output "Firefox cache folder not found, skipping."
}

# Network Cleanup Commands
Write-Output "Starting network cleanup..."
ipconfig /flushdns
Write-Output "DNS cache flushed."

nbtstat -R
Write-Output "NetBIOS names refreshed."

nbtstat -c
Write-Output "NetBIOS cache cleared."

ipconfig /release
Write-Output "IP address released."

ipconfig /renew
Write-Output "IP address renewed."

Write-Output "All cleanup operations completed."
