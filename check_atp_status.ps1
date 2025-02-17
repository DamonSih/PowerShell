<#
.SYNOPSIS
  To check the onboarding status and sensor status of the ATP

.DESCRIPTION
  This script checks the onboarding status of Advanced Threat Protection and the service status of the Advanced Treat Protection Sensor
    
.NOTES
	Author: Damon Sih Boon Kiat

#>

# Define the registry path and value name of ATP
$registryPath = "HKLM\SOFTWARE\Microsoft\ Windows Advanced Threat Protection\Status"
$valueName ="OnboardingState"

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
        return “Service not found" 
}

# Check if registry key exists
if (Test-Path $registryPath) {
    # Get the registry value
    $value = Get-ItemProperty -Path $registryPath -Name $valueName -ErrorAction Stop
    # Check if the value is 1
    if ($value.OnboardingState -eq 1) {
        Write-Output "The OnboardingState is set to 1." -ForegroundColor Green
    } else {
        Write-Output "The OnboardingState is NOT set to 1. Current value: $($value.OnboardingState)" -ForegroundColor Yellow
    ]
} else {
    Write-Output "The registry path $registryPath does not exist." -ForegroundColor Yellow
]

# Check the status of the service
$serviceStatus = Get-ServiceStatus -serviceName $serviceName
Write-Output "The status of the service '$serviceName' is: $serviceStatus"
