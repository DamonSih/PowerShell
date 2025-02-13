<#
.SYNOPSIS
	Gather and export user properties from AD into a CSV file and send this report via email
.DESCRIPTION
	This sample demonstrates how to gather and export user properties from Active Directory (AD) into a CSV file, and then send this report via email.
.EXAMPLE
	PS> ./AD_users_report.ps1
	https://github.com/themazdarati/PowerShell/ActiveDirectory
.NOTES
	Author: Damon Sih Boon Kiat | License: CC0
#>

$Properties = @(
    "Name",
    "SamAccountName",
    "UserPrincipalName",
    "EmployeeID",
    "EmailAddress",
    "extensionAttribute5",
    "Manager",
    "Department",
    "extensionAttribute1",
    "Company",
    "extensionAttribute10",
    "extensionAttribute12",
    "Office",
    "Country",
    "PasswordLastSet",
    "msDS-UserPasswordExpiryTimeComputed",
    "extensionAttribute11",
    "Enabled",
    "WhenCreated",
    "WhenChanged",
    "targetAddress",
    "proxyAddresses",
    "accountExpires",
    "LastLogonDate",
    "lastLogonTimestamp",
	"CanonicalName"
)

# Lookup user's properties
$Users = Get-ADUser -Filter * -Properties $Properties |
Select-Object @{Name="Full Name"; Expression={$_.Name}},
              @{Name="Sam Account Name"; Expression={$_.SamAccountName}},
              @{Name="Logon Name"; Expression={$_.UserPrincipalName}},
              @{Name="Employee ID"; Expression={$_.EmployeeID}},
              @{Name="Email Address"; Expression={$_.EmailAddress}},
              @{Name="Hire Date (extensionAtrribute5)"; Expression={$_.extensionAttribute5}},
              @{Name="Manager"; Expression={(Get-ADUser $_.Manager -ErrorAction SilentlyContinue).Name}},
              @{Name="Department"; Expression={$_.Department}},
              @{Name="Company Code (extensionAttribute1)"; Expression={$_.extensionAttribute1}},
              @{Name="Company"; Expression={$_.Company}},
              @{Name="Owner (extensionAttribute10)"; Expression={$_.extensionAttribute10}},
              @{Name="Cost Centre (extensionAttribute12)"; Expression={$_.extensionAttribute12}},
			        @{Name="Office"; Expression={$_.Office}},
			        @{Name="CanonicalName"; Expression={$_.CanonicalName}},
			        @{Name="Country"; Expression={$_.Country}},
              @{Name="Password Last Set"; Expression={$_.PasswordLastSet.ToString("MM/dd/yyyy")}},
              @{Name="Password Expiry Date"; Expression={ 
                  if ($_."msDS-UserPasswordExpiryTimeComputed" -ne 9223372036854775807) {
                      [datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed").ToString("MM/dd/yyyy")
                  } else { "Never Expires" }
              }},
              @{Name="Account Type (extensionAttribute11)"; Expression={$_.extensionAttribute11}},
              @{Name="Account Status"; Expression={if ($_.Enabled) { "Enabled" } else { "Disabled" }}},
              @{Name="When Created"; Expression={$_.WhenCreated.ToString("MM/dd/yyyy")}},
              @{Name="When Modified"; Expression={$_.WhenChanged.ToString("MM/dd/yyyy")}},
              @{Name="targetAddress"; Expression={$_.targetAddress}},
              @{Name="Proxy Email Address"; Expression={($_.proxyAddresses | Where-Object {$_ -like "SMTP:*"}) -replace "SMTP:",""}},
              @{Name="Account Expiry"; Expression={
                  if ($_.accountExpires -ne 9223372036854775807) {
                      [datetime]::FromFileTime($_.accountExpires).ToString("MM/dd/yyyy")
                  } else { "Never Expires" }
              }},
              @{Name="Last Logon Date"; Expression={$_.LastLogonDate.ToString("MM/dd/yyyy")}},
              @{Name="Last Logon Timestamp"; Expression={[datetime]::FromFileTime($_.lastLogonTimestamp).ToString("MM/dd/yyyy")}}

$tempFile = "C:\Temp\AD_Users_Full_Export.csv"
$Users | Export-Csv -Path $tempFile -NoTypeInformation -Encoding UTF8
Write-Host "Export is completed successfully! File Path: $tempFile" -ForegroundColor Green

# SMTP configuration
$SMTPServer = "your_smtp_server"
$From = "from_address"
$To = "to_address","to_address1"
$Subject = "Daily Report - AD Users Report"
$messageBody = "The report generated from your_server as per your schedule has been sent in this email as an attachement."

Send-MailMessage -SmtpServer $SMTPServer -From $From -To $To -Subject $Subject -Body $messageBody -Attachments $tempFile

Write-Host "Email sent successfully!" -ForegroundColor Green

# Clean up the temporary file
Remove-Item $tempFile

Write-Host "Clean up is completed!" -ForegroundColor Green
