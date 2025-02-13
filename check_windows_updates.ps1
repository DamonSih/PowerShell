<#
.SYNOPSIS
  Checks remote computers for pending Windows updates.

.DESCRIPTION
  This script checks one or more remote computers for pending Windows updates.
  It executes parallel operations for efficiency and provides detailed output.
  Features include customizable timeout settings, error handling, and connection throttling.

.PARAMETER Computers
  Specifies the target computers. Enter multiple names separated by commas.

.PARAMETER ConnectionTimeout
  Specifies the connection and operation timeout in seconds (default = 10).

.PARAMETER ThrottleLimit
  Specifies the maximum number of concurrent connections (default = 32).

.EXAMPLE
  PS> .\PS_Remote_Check_WindowsUpdates.ps1 -Computers server1,server2 -ConnectionTimeout 15

    Server    PendingUpdates Status
    ------    -------------- ------
    server1              5 Success
    server2              0 Success
    server3               Connection failed: WinRM not available
    
.NOTES
	Author: Damon Sih Boon Kiat | License: CC0

#>

Param(
    [Parameter(Mandatory=$True, Position=0)]
    [String[]] $Computers,

    [Parameter()]
    [ValidateRange(1, 300)]
    [int] $ConnectionTimeout = 10,

    [Parameter()]
    [int] $ThrottleLimit = 32
)

$scriptBlock = {
    $result = [PSCustomObject]@{
        Server          = $env:COMPUTERNAME
        PendingUpdates  = $null
        Status          = $null
    }

    try {
        $SearchCriteria = "IsInstalled=0 And Type='Software' And IsHidden=0"
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session -ErrorAction Stop
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
        $SearchResult = $UpdateSearcher.Search($SearchCriteria)
        
        $result.PendingUpdates = $SearchResult.Updates.Count
        $result.Status = "Success"
    }
    catch {
        $result.Status = "Execution error: $($_.Exception.Message)"
    }
    
    return $result
}

$sessionOptions = New-PSSessionOption -OpenTimeout ($ConnectionTimeout * 1000) -OperationTimeout ($ConnectionTimeout * 1000)

try {
    $results = Invoke-Command -ComputerName $Computers -ScriptBlock $scriptBlock `
        -SessionOption $sessionOptions -ThrottleLimit $ThrottleLimit `
        -ErrorAction SilentlyContinue -ErrorVariable connectionErrors

    # Process connection errors
    $allResults = @($results)
    if ($connectionErrors) {
        foreach ($errorRecord in $connectionErrors) {
            $computer = $errorRecord.TargetObject
            $exception = $errorRecord.Exception

            # Generate friendly error message
            $statusMsg = switch -Wildcard ($exception.Message) {
                '*WinRM cannot complete the operation*'   { 'Connection failed: WinRM not available' }
                '*access is denied*'                     { 'Access denied' }
                '*The client cannot connect*'             { 'Connection refused' }
                '*could not be resolved*'                 { 'Computer name not found' }
                default                                  { "Connection error: $($exception.Message)" }
            }

            $allResults += [PSCustomObject]@{
                Server          = $computer
                PendingUpdates  = $null
                Status          = $statusMsg
            }
        }
    }

    # Display results
    $allResults | Select-Object Server, PendingUpdates, Status | Sort-Object Server | Format-Table -AutoSize

    # Additional error logging
    if ($connectionErrors) {
        Write-Verbose "Detailed connection errors:" -Verbose
        $connectionErrors | ForEach-Object { Write-Verbose "$_" -Verbose }
    }
}
catch {
    Write-Host "Critical error: $_" -ForegroundColor Red
    exit 1
}
