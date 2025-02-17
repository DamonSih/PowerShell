<#
.SYNOPSIS
  To check the onboarding status and sensor status of the ATP, and email the status of the service to admin

.DESCRIPTION
  This script checks the onboarding status of Advanced Threat Protection and the service status of the Advanced Treat Protection Sensor
  Email admin if service not found/failed to start after 3 attempts

.PARAMETER EmailTo
Recipient email address

.PARAMETER EmailFrom
Sender email address

.PARAMETER SmtpServer
SMTP server address

.EXAMPLE
.\check_atp_status.ps1 -EmailTo "admin@example.com" -EmailFrom "updates@example.com" -SmtpServer "smtp.office365.com"


.NOTES
  Author: Damon Sih Boon Kiat
#>

param (
    [string]$EmailTo,
    [string]$EmailFrom,
    [string]$SmtpServer
)

# Define the registry path and value name of ATP
$registryPath = "HKLM:\SOFTWARE\Microsoft\Windows Advanced Threat Protection\Status"
$valueName = "OnboardingState"

# Define the service name
$serviceName = "Azure Advanced Threat Protection Sensor"

# Function to check service status
function Get-ServiceStatus {
    param (
        [string]$serviceName
    )
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    if ($service) {
        return $service.Status
    } else {
        return "Service not found"
    }
}

# Function to send email
function Send-EmailNotification {
    param (
        [string]$subject,
        [string]$body
    )
    try {
        Send-MailMessage -From $EmailFrom -To $EmailTo -Subject $subject -Body $body -SmtpServer $SmtpServer -UseSsl
        Write-Output "Email notification sent to $EmailTo."
    } catch {
        Write-Output "Failed to send email notification: $_"
    }
}

# Function to start the service
function Start-ServiceWithRetry {
    param (
        [string]$serviceName,
        [int]$maxRetries = 3
    )
    $retryCount = 0
    while ($retryCount -lt $maxRetries) {
        $serviceStatus = Get-ServiceStatus -serviceName $serviceName
        if ($serviceStatus -eq "Running") {
            Write-Output "The service '$serviceName' is now running."
            return $true
        } else {
            Write-Output "Attempt $($retryCount + 1): Starting the service '$serviceName'..."
            Start-Service -Name $serviceName -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 10  # 10 seconds buffer
            $retryCount++
        }
    }
    return $false
}

# Check if registry key exists
if (Test-Path $registryPath) {
    # Try get the registry value
    try {
        $value = Get-ItemProperty -Path $registryPath -Name $valueName -ErrorAction Stop
    
        # Check if the value is 1
        if ($value.OnboardingState -eq 1) {
            Write-Output "The OnboardingState is set to 1." -ForegroundColor Green
        } else {
            Write-Output "The OnboardingState is NOT set to 1. Current value: $($value.OnboardingState)" -ForegroundColor Yellow
        }
    } catch {
        Write-Output "The property '$valueName' does not exist in the registry path '$registryPath'."
    }
} else {
    Write-Output "The registry path $registryPath does not exist." -ForegroundColor Yellow
}

# Check the status of the service
$serviceStatus = Get-ServiceStatus -serviceName $serviceName
Write-Output "The status of the service '$serviceName' is: $serviceStatus"

# If the service is not found, send email
if ($serviceStatus -eq "Service not found") {
    $serverName = $env:COMPUTERNAME 
    $emailSubject = "Service Not Found: $serviceName on $serverName"
    $emailBody = @"
The service '$serviceName' on server '$serverName' was not found. Please check the server and take necessary actions.

Details:
- Service Name: $serviceName
- Server Name: $serverName
- Time: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"@
    Send-EmailNotification -subject $emailSubject -body $emailBody
} elseif ($serviceStatus -ne "Running") {
    $startSuccess = Start-ServiceWithRetry -serviceName $serviceName -maxRetries 3
    if (-not $startSuccess) {
        Write-Output "Failed to start the service '$serviceName' after 3 attempts"
        $serverName =$env:COMPUTERNAME
        $emailSubject = "Service Start Failed: $serviceName on $serverName"
        $emailBody = @"
The service '$serviceName' on server '$serverName' failed to start after 3 attempts. Please check the server and take necessary actions.

Details:
- Service Name: $serviceName
- Server Name: $serverName
- Time: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"@
        Send-EmailNotification -subject $emailSubject -body $emailBody
    }
}
