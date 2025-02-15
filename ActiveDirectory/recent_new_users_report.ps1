<#
.SYNOPSIS
  Generate and send recent new user report

.DESCRIPTION
  This script is used to retrieve newly created Active Directory user accounts within a specified time period, generate a CSV report, and send it to specified recipients via email.

.PARAMETER DaysBack
  Query users created within the last number of days (default 7 days).

.PARAMETER SearchBase
  Target Organizational Unit distinguished name

.PARAMETER ReportPath
  Full path for temporary report file

.PARAMETER ToEmail
  Recipient email addresses (multiple addresses separated by commas)

.PARAMETER FromEmail
  Sender email address

.PARAMETER SmtpServer
  SMTP server hostname/IP

.EXAMPLE
  .\recent_new_users_report.ps1 -DaysBack 3 `
      -SearchBase "OU=Users,DC=contoso,DC=com" `
      -ReportPath "C:\Reports\NewUsers.csv" `
      -ToEmail "hr@contoso.com" `
      -FromEmail "noreply@contoso.com" `
      -SmtpServer "smtp.office365.com"

.NOTES
	Author: Damon Sih Boon Kiat
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 365)]
    [int]$DaysBack = 7,

    [Parameter(Mandatory)]
    [ValidateScript({
        try { Get-ADObject $_ } catch { 
            throw "Invalid OU: $_" 
        }
    })]
    [string]$SearchBase,

    [Parameter(Mandatory)]
    [ValidateScript({
        $parentDir = Split-Path $_ -Parent
        if(-not (Test-Path $parentDir -PathType Container)) {
            throw "Invalid report path directory"
        }
        $true
    })]
    [string]$ReportPath,

    [Parameter(Mandatory)]
    [ValidatePattern('^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$')]
    [string[]]$ToEmail,

    [Parameter(Mandatory)]
    [ValidatePattern('^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$')]
    [string]$FromEmail,

    [Parameter(Mandatory)]
    [string]$SmtpServer
)

# Initialize Error Collection
$ErrorActionPreference = 'Stop'
$script:ErrorList = [System.Collections.Generic.List[object]]::new()

function Get-NewADUsers {
    param(
        [int]$Days,
        [string]$OU
    )

    try {
        $startDate = (Get-Date).AddDays(-$Days)
        $properties = @(
            'Created',
            'SamAccountName',
            'DisplayName',
            'UserPrincipalName',
            'EmailAddress',
            'Department',
            'Title',
            'Enabled'
        )

        $filter = "Created -ge '$($startDate.ToString('yyyy-MM-dd'))'"
        
        Get-ADUser -Filter $filter `
            -SearchBase $OU `
            -Properties $properties `
            -ErrorAction Stop |
        Select-Object $properties |
        Sort-Object Created -Descending

    } catch {
        $script:ErrorList.Add([PSCustomObject]@{
            Timestamp = Get-Date
            Type      = 'AD查询'
            Message   = $_.Exception.Message
        })
        throw
    }
}

function Export-NewUserReport {
    param(
        [object[]]$Users,
        [string]$Path
    )

    try {
        if ($Users.Count -eq 0) {
            Write-Verbose "No new users found"
            return $null
        }

        $reportData = $Users | ForEach-Object {
            [PSCustomObject]@{
                whenCreated    = $_.Created.ToString("yyyy-MM-dd HH:mm")
                SamAccountName = $_.SamAccountName
                Name           = $_.DisplayName
                UPN            = $_.UserPrincipalName
                Email address  = $_.EmailAddress
                Deparment      = $_.Department
                Title          = $_.Title
                Account status = if ($_.Enabled) { "Enabled" } else { "Disabled" }
            }
        }

        $reportData | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        Write-Output "Report exported：$Path"
        return $Path

    } catch {
        $script:ErrorList.Add([PSCustomObject]@{
            Timestamp = Get-Date
            Type      = 'File Export'
            Message   = $_.Exception.Message
        })
        throw
    }
}

function Send-NewUserNotification {
    param(
        [string]$ReportFile,
        [string[]]$Recipients,
        [string]$Sender,
        [string]$SmtpHost
    )

    try {
        $mailParams = @{
            From       = $Sender
            To         = $Recipients
            Subject    = "New User Report - Last $DaysBack Days ($(Get-Date -Format 'yyyy-MM-dd'))"
            Body       = "The attachment contains a list of user accounts created in the last $DaysBack days."
            SmtpServer = $SmtpHost
            Priority   = 'High'
        }

        if ($ReportFile -and (Test-Path $ReportFile)) {
            $mailParams['Attachments'] = $ReportFile
            $userCount = @(Import-Csv $ReportFile).Count
            $mailParams.Body += " (A total of $userCount new accounts were found.)"
        } else {
            $mailParams.Body = "No new user accounts were found created in the last $DaysBack days."
        }

        Send-MailMessage @mailParams -ErrorAction Stop
        Write-Output "Report has been sent to: $($Recipients -join ', ')"

    } catch {
        $script:ErrorList.Add([PSCustomObject]@{
            Timestamp = Get-Date
            Type      = 'Email'
            Message   = $_.Exception.Message
        })
        throw
    } finally {
        if ($ReportFile -and (Test-Path $ReportFile)) {
            Remove-Item $ReportFile -Force -ErrorAction SilentlyContinue
        }
    }
}

# Main
try {
    # Retrieve new users in the last $DaysBack
    $newUsers = Get-NewADUsers -Days $DaysBack -OU $SearchBase
    
    # Generate report
    $reportFile = Export-NewUserReport -Users $newUsers -Path $ReportPath
    
    # Email report
    Send-NewUserNotification -ReportFile $reportFile `
                           -Recipients $ToEmail `
                           -Sender $FromEmail `
                           -SmtpHost $SmtpServer

} catch {
    Write-Error "Processing failed: $_"
}

# Error Report
if ($script:ErrorList.Count -gt 0) {
    $errorReport = "Encountered $($script:ErrorList.Count) errors：`n"
    $errorReport += $script:ErrorList | Format-Table -AutoSize | Out-String
    
    Write-Warning $errorReport
    
    try {
        Send-MailMessage -From $FromEmail -To $ToEmail `
                       -Subject "Users Report Errors - $(Get-Date -Format 'yyyy-MM-dd')" `
                       -Body $errorReport -SmtpServer $SmtpServer
    } catch {
        Write-Error "Failed to send error report: $_"
    }
}
