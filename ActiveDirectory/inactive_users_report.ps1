<#
.SYNOPSIS
  Generates and sends inactive user report for specified Organizational Unit.

.DESCRIPTION
  This script identifies inactive users (no logon in 60 days) in a specified AD OU and generates a CSV report, and emails it to designated recipients.

.PARAMETER SearchBase
  Target Organizational Unit distinguished name

.PARAMETER ReportPath
  Full path for temporary report file

.PARAMETER ToEmail
  Recipient email address(es)

.PARAMETER FromEmail
  Sender email address

.PARAMETER SmtpServer
  SMTP server hostname/IP

.PARAMETER RetentionDays
Days since last logon to consider inactive (default: 60)

.EXAMPLE
  .\inactive_users_report.ps1 -OU "OU=Users,DC=contoso,DC=com" `
      -ReportPath "C:\Reports\InactiveUsers.csv" `
      -ToEmail "it@contoso.com" `
      -FromEmail "noreply@contoso.com" `
      -SmtpServer "smtp.contoso.com"
    
.NOTES
	Author: Damon Sih Boon Kiat | License: CC0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ 
        try { Get-ADObject $_ } catch { throw "Invalid OU: $_" }
    })]
    [string]$SearchBase,

    [Parameter(Mandatory)]
    [ValidateScript({
        $parent = Split-Path $_ -Parent
        if(-not (Test-Path $parent -PathType Container)) {
            throw "Invalid report path directory"
        }
        $true
    })]
    [string]$ReportPath,

    [Parameter(Mandatory)]
    [ValidatePattern('^.+@.+\..+$')]
    [string[]]$ToEmail,
    
    [Parameter(Mandatory)]
    [ValidatePattern('^.+@.+\..+$')]
    [string]$FromEmail,

    [Parameter(Mandatory)]
    [string]$SmtpServer,

    [ValidateRange(1, 365)]
    [int]$RetentionDays = 60
)

# "Initialize Error Collection
$ErrorActionPreference = 'Stop'
$script:Errors = [System.Collections.Generic.List[object]]::new()

function Get-InactiveUsers {
    param(
        [string]$SearchBase,
        [int]$InactiveDays
    )

    $thresholdDate = (Get-Date).AddDays(-$InactiveDays)
    $properties = @(
        'SamAccountName',
        'DisplayName',
        'EmailAddress',
        'EmployeeID',
        'Company',
        'Department',
        'Title',
        'LastLogonDate',
        'Enabled'
    )

    try {
        $adParams = @{
            SearchBase = $SearchBase
            Filter = "Enabled -eq 'true'"
            Properties = $properties
        }

        Get-ADUser @adParams | Where-Object {
            $_.LastLogonDate -and 
            $_.LastLogonDate -lt $thresholdDate -and
            $_.EmailAddress -match '^.+@.+\..+$'
        } | Select-Object $properties

    } catch {
        $script:Errors.Add([PSCustomObject]@{
            Time  = Get-Date
            Type  = 'AD Query'
            Error = $_.Exception.Message
        })
        throw
    }
}

function Export-UserReport {
    param(
        [object[]]$Users,
        [string]$Path
    )

    try {
        if ($Users.Count -eq 0) {
            Write-Verbose "No inactive users found"
            return $null
        }

        $Users | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        Write-Output "Report exported to $Path"
        return $Path

    } catch {
        $script:Errors.Add([PSCustomObject]@{
            Time  = Get-Date
            Type  = 'File Export'
            Error = $_.Exception.Message
        })
        throw
    }
}

function Send-ReportNotification {
    param(
        [string]$ReportFile,
        [string[]]$Recipients,
        [string]$Sender,
        [string]$Server
    )

    try {
        $mailParams = @{
            From         = $Sender
            To           = $Recipients
            Subject      = "Inactive Users Report - $(Get-Date -Format 'yyyy-MM-dd')"
            Body         = "Attached is the inactive users report."
            SmtpServer   = $Server
            ErrorAction = 'Stop'
        }

        if ($ReportFile -and (Test-Path $ReportFile)) {
            $mailParams['Attachments'] = $ReportFile
            $mailParams['Body'] = "Found $(@(Import-Csv $ReportFile).Count) inactive users."
        } else {
            $mailParams['Body'] = "No inactive users found during this scan."
        }

        Send-MailMessage @mailParams
        Write-Output "Report notification sent to $($Recipients -join ', ')"

    } catch {
        $script:Errors.Add([PSCustomObject]@{
            Time  = Get-Date
            Type  = 'Email'
            Error = $_.Exception.Message
        })
        throw
    } finally {
        # Cleanup report file
        if ($ReportFile -and (Test-Path $ReportFile)) {
            Remove-Item $ReportFile -Force -ErrorAction SilentlyContinue
        }
    }
}

# Main
try {
    # Retrieve inactive users
    $inactiveUsers = Get-InactiveUsers -SearchBase $OU -InactiveDays $RetentionDays
    
    # Generate report
    $reportFile = Export-UserReport -Users $inactiveUsers -Path $ReportPath
    
    # send email
    Send-ReportNotification -ReportFile $reportFile `
                          -Recipients $ToEmail `
                          -Sender $FromEmail `
                          -Server $SmtpServer

} catch {
    Write-Error "Processing failed: $_"
}

# Error Report to Admin
if ($script:Errors.Count -gt 0) {
    $errorReport = $script:Errors | Format-List | Out-String
    Write-Warning "Encountered $($script:Errors.Count) errors:`n$errorReport"
    
    try {
        Send-MailMessage -From $FromEmail -To $ToEmail `
                       -Subject "Report Errors - $(Get-Date -Format 'yyyy-MM-dd')" `
                       -Body $errorReport -SmtpServer $SmtpServer
    } catch {
        Write-Error "Failed to send error report: $_"
    }
}
