<# 
.SYNOPSIS
Automates Windows Update download and installation process with enhanced efficiency and sends an email report after completion.

.PARAMETER EmailTo
Recipient email address

.PARAMETER EmailFrom
Sender email address

.PARAMETER SmtpServer
SMTP server address

.PARAMETER AutoRestart
Enable automatic restart if required (default: true)

.EXAMPLE
.\WindowsUpdate.ps1 -EmailTo "admin@example.com" -EmailFrom "updates@example.com" -SmtpServer "smtp.office365.com" -AutoRestart $true

.NOTES
Author: Damon Sih Boon Kiat
#>

#Requires -RunAsAdministrator

param (
    [string]$EmailTo,
    [string]$EmailFrom,
    [string]$SmtpServer,
    [bool]$AutoRestart = $true
)

#Logging
$logPath = Join-Path $PSScriptRoot "WindowsUpdate.log"
function Write-Log {
    param([string]$message, [string]$level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp][$level] $message"
    Add-Content -Path $logPath -Value $logEntry
    Write-Host $logEntry -ForegroundColor $(if ($level -eq "ERROR") { "Red" } else { "White" })
}

#Email Functions
function Send-EmailReport {
    param (
        [string]$subject,
        [string]$body
    )

    try {
        $mailParams = @{
            To         = $EmailTo
            From       = $EmailFrom
            SmtpServer = $SmtpServer
            Subject    = $subject
            Body       = $body
            BodyAsHtml = $true
            ErrorAction = 'Stop'
        }

        Send-MailMessage @mailParams
        Write-Log "Email report sent successfully"
    }
    catch {
        Write-Log "Failed to send email report: $_" -level "ERROR"
    }
}

function Format-HTMLReport {
    param($updates, $installResult, $machineName)

    $html = @"
<html>
<head><style>
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    tr:nth-child(even) { background-color: #f2f2f2; }
    .success { color: green; }
    .error { color: red; }
</style></head>
<body>
<h2>Windows Update Report for $machineName</h2>
<table>
    <tr><th>Update Title</th><th>Status</th><th>Details</th></tr>
"@

    for ($i = 0; $i -lt $updates.Count; $i++) {
        $result = $installResult.GetUpdateResult($i)
        $status = if ($result.ResultCode -eq 2) {
            "<span class='success'>Installed</span>"
        } else {
            "<span class='error'>Failed ($($result.ResultCode))</span>"
        }
        $html += "<tr><td>$($updates[$i].Title)</td><td>$status</td><td>$($result.HResult)</td></tr>"
    }

    $html += "</table></body></html>"
    return $html
}


#Update Functions
function Invoke-WithRetry {
    param (
        [ScriptBlock]$Action,
        [int]$MaxRetries = 3,
        [int]$RetryDelay = 30
    )

    $attempt = 1
    do {
        try {
            return & $Action
        }
        catch {
            if ($attempt -ge $MaxRetries) { throw }
            Write-Log "Attempt $attempt failed: $_ - Retrying in $RetryDelay seconds..." -level "WARNING"
            Start-Sleep -Seconds $RetryDelay
            $attempt++
        }
    } while ($true)
}

function Invoke-WindowsUpdate {
    try {
        $session = New-Object -ComObject 'Microsoft.Update.Session'
        $searcher = $session.CreateUpdateSearcher()
        $searcher.ServerSelection = 0

        # Get machine name
        $machineName = $env:COMPUTERNAME

        # Search updates
        Write-Log "Searching for updates..."
        $searchResult = Invoke-WithRetry { $searcher.Search("IsInstalled=0") }
        if ($searchResult.Updates.Count -eq 0) {
            Write-Log "System is up to date"
            Send-EmailReport -subject "Windows Update Report - No Updates" -body "No updates available on $machineName"
            return
        }

        # Prepare updates
        $updates = $searchResult.Updates
        $downloadCollection = New-Object -ComObject 'Microsoft.Update.UpdateColl'
        $updates | ForEach-Object {
            if (-not $_.EulaAccepted) { $_.AcceptEula() }
            $downloadCollection.Add($_) | Out-Null
        }

        # Download updates
        Write-Log "Downloading $($downloadCollection.Count) updates..."
        $downloader = $session.CreateUpdateDownloader()
        $downloader.Updates = $downloadCollection
        $downloadResult = Invoke-WithRetry { $downloader.Download() }

        # Filter downloaded updates
        $installCollection = New-Object -ComObject 'Microsoft.Update.UpdateColl'
        $downloadCollection | Where-Object IsDownloaded | ForEach-Object {
            $installCollection.Add($_) | Out-Null
        }

        # Install updates
        if ($installCollection.Count -gt 0) {
            Write-Log "Installing $($installCollection.Count) updates..."
            $installer = $session.CreateUpdateInstaller()
            $installer.Updates = $installCollection
            $installResult = Invoke-WithRetry { $installer.Install() }

            # Generate report
            $htmlReport = Format-HTMLReport $installCollection $installResult $machineName
            $subject = "Windows Update Report for $machineName - " + $(if ($installResult.RebootRequired) {
                "Reboot Required" } else { "Success" })

            Send-EmailReport -subject $subject -body $htmlReport

            if ($installResult.RebootRequired -and $AutoRestart) {
                Write-Log "System will reboot in 2 minutes"
                shutdown /r /f /t 120 /c "Automated update completed on $machineName"
            }
        }
        else {
            Write-Log "No updates downloaded successfully"
            Send-EmailReport -subject "Windows Update Report for $machineName - Download Failed" -body "All update downloads failed on $machineName"
        }
    }
    catch {
        Write-Log "Update process failed: $_" -level "ERROR"
        Send-EmailReport -subject "Windows Update Report for $machineName - Critical Error" -body "Error during update process on $machineName<br><pre>$($_)</pre>"
        exit 1
    }
}


#main
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "Elevation required. Run script as Administrator." -level "ERROR"
    exit 1
}

try {
    Invoke-WindowsUpdate
    Write-Log "Update process completed successfully"
}
catch {
    Write-Log "Fatal error in main execution: $_" -level "ERROR"
    exit 1
}
