<#
.SYNOPSIS
	Automates the process of notifying users about impending password expirations.
.DESCRIPTION
	This sample demonstrates how administrators ensure users are aware of upcoming password changes and can take action to reset their password on time.
.EXAMPLE
	PS> ./password_expiry_notifier.ps1
.NOTES
	Author: Damon Sih Boon Kiat | License: CC0
#>

$ErrorActionPreference = 'SilentlyContinue'

# Define SMTP Server settings
$SMTPFrom = "from_address"
$SMTPServer = "your_smtp_server"
$PasswordResetURL = "https://passwordreset.microsoftonline.com/"
$ExpiryDaysNotice = @(14, 7, 1, 2, 3)  # Days notice for expiration warnings

# Function to send email notifications
function Send-PasswordExpiryNotification {
    param (
        [Parameter(Mandatory = $true)]
        [PSObject] $User,
        [Parameter(Mandatory = $true)]
        [int] $DaysRemaining
    )

    $Emailbody = @"
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Password Expiry Notification</title>
        <style>
            body { font-family: Arial, sans-serif; }
            .header { width: 100%; background-color: #00a19a; padding: 10px; }
            .notification-bar { background-color: #4b4b4b; color: white; padding: 10px; text-align: center; font-weight: bold; }
            .content { padding: 20px; }
            .content p { margin-bottom: 20px; font-size: 14px; }
            .content a { color: #0066cc; }
            .footer { background-color: #EA2B29; color: white; margin-top: 40px; font-size: 12px; text-align: center; }
            .footer-note { font-size: 11px; color: gray; margin-top: 30px; }
        </style>
    </head>
    <body>
        <div class="notification-bar">
            Important! Your password will expire in $DaysRemaining days!
        </div>
        <div class="content">
            <p>Dear $($User.displayName),</p>
            <p>Your account's current password will expire in $DaysRemaining days as per our policy. Please click the link below to reset your password to continue accessing your account:</p>
            <p><a href="$PasswordResetURL">$PasswordResetURL</a></p>
            <p>If you need help, please request assistance from the Service Desk.</p>
            <hr>
            <p>Kata laluan semasa akaun anda akan tamat tempoh dalam $DaysRemaining hari. Sila klik pautan di bawah untuk menetapkan semula kata laluan anda supaya anda boleh terus mengakses akaun anda:</p>
            <p><a href="$PasswordResetURL">$PasswordResetURL</a></p>
            <p>Jika anda memerlukan bantuan, sila minta bantuan daripada Service Desk.</p>
            <div class="footer-note">
                Please do not reply to this automated message.
            </div>
        </div>
        <div class="footer">
            Any unauthorized access, use, or distribution is prohibited. If you are not the user of this system, you are not authorized to enter, duplicate, or disseminate any part of the system.
        </div>
    </body>
    </html>
"@

    # Send Email Notification
    Send-MailMessage -From $SMTPFrom -To $User.mail -Subject "Important! Your password will expire soon!" -Body $Emailbody -BodyAsHtml -SmtpServer $SMTPServer
}

# Process users in batches to reduce memory usage
Get-ADUser -Filter {Enabled -eq $true -and PasswordNeverExpires -eq $false} -Properties msDS-UserPasswordExpiryTimeComputed, mail, displayName | ForEach-Object -Process {
    $pwdExpiryTime = [datetime]::FromFileTime([int64]$_.msDS-UserPasswordExpiryTimeComputed)
    $expire_days = ($pwdExpiryTime - (Get-Date)).Days

    if ($ExpiryDaysNotice -contains $expire_days) {
        Send-PasswordExpiryNotification -User $_ -DaysRemaining $expire_days
    }
}

Write-Host "Process completed successfully!" -ForegroundColor Green
