#requires -Version 5.1
#requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Network-wide system inventory orchestrator
    
.DESCRIPTION
    Queries Active Directory for computer accounts, determines which are active based on
    password age, and runs inventory on each active computer. Tracks progress, monitors
    child processes, and enforces timeouts.
    
.PARAMETER InactiveDays
    Number of days since last password change to consider a computer inactive (default: 90)
    
.PARAMETER OrphanedDays
    Number of days since last password change to consider a computer orphaned (default: 1095/3 years)
    
.PARAMETER MaxConcurrent
    Maximum number of concurrent inventory jobs (default: 10)
    
.PARAMETER Timeout
    Timeout for each inventory job in seconds (default: 180/3 minutes)
    
.PARAMETER OutputPath
    Path to Reports folder (default: .\Reports)
    
.PARAMETER SoftwareXmlPath
    Path to software.xml (default: .\software.xml)
    
.PARAMETER SkipInventory
    Only generate ComputerAccounts.log, don't run inventories
    
.PARAMETER IncludeWin32Product
    Include Win32_Product in software collection (slower)
    
.PARAMETER Credential
    Credentials for remote access
    
.EXAMPLE
    .\NetworkInventory.ps1
    Run with default settings (90 day threshold, 10 concurrent)
    
.EXAMPLE
    .\NetworkInventory.ps1 -InactiveDays 60 -MaxConcurrent 20
    Custom thresholds
    
.EXAMPLE
    .\NetworkInventory.ps1 -SkipInventory
    Only generate AD computer report, don't run inventories
    
.NOTES
    Author: System Admin
    Version: 1.0
    Based on: Inventory.vbs
    Must be run from Domain Controller or computer with AD PowerShell module
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [int]$InactiveDays = 90,
    
    [Parameter(Mandatory=$false)]
    [int]$OrphanedDays = 1095,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxConcurrent = 10,
    
    [Parameter(Mandatory=$false)]
    [int]$Timeout = 180,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\Reports",
    
    [Parameter(Mandatory=$false)]
    [string]$SoftwareXmlPath = ".\software.xml",
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipInventory,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeWin32Product,
    
    [Parameter(Mandatory=$false)]
    [PSCredential]$Credential
)

$ErrorActionPreference = 'Continue'
$scriptStartTime = Get-Date

#region Helper Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Warning','Error','Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch($Level) {
        'Info' { 'Cyan' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        'Success' { 'Green' }
    }
    
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Initialize-Environment {
    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-Log "Created output directory: $OutputPath" -Level Success
    }
    
    # Clean up old report files
    $computerAccountsLog = Join-Path $OutputPath "ComputerAccounts.log"
    $needsUpdateFile = Join-Path $OutputPath "NeedsUpdate.txt"
    
    if (Test-Path $computerAccountsLog) {
        Remove-Item $computerAccountsLog -Force
        Write-Log "Removed existing ComputerAccounts.log" -Level Info
    }
    
    if (Test-Path $needsUpdateFile) {
        Remove-Item $needsUpdateFile -Force
        Write-Log "Removed existing NeedsUpdate.txt" -Level Info
    }
    
    # Create header for NeedsUpdate.txt
    "The following PCs need the listed software updated to the latest version and security patches applied:`n" | 
        Out-File -FilePath $needsUpdateFile -Encoding UTF8 -Force
    
    # Copy/create CSS and icon if they don't exist
    $cssFile = Join-Path $OutputPath "style.css"
    $iconFile = Join-Path $OutputPath "icon.png"
    
    if (-not (Test-Path $cssFile)) {
        # Create default CSS
        @"
table { 
    padding: 0px;
    margin: 0px 0px 10px 0px;
    border-collapse: collapse;
    width: 100%;
}
h1 {
    color: #B03726; 
    font-family: Tahoma, Verdana, Arial, Sans-Serif; 
    font-size: 16pt;
    margin-top: 20px;
}
h2 {
    color: #B03726; 
    font-family: Tahoma, Verdana, Arial, Sans-Serif; 
    font-size: 14pt;
    margin-top: 15px;
}
h3 {
    color: #B03726; 
    font-family: Tahoma, Verdana, Arial, Sans-Serif; 
    font-size: 12pt;
    margin-top: 10px;
}
th {
    color: white; 
    font-family: Tahoma, Verdana, Arial, Sans-Serif;  
    font-style: italic;
    border: solid black 1.5pt; 
    background-color: #B03726;
    padding: 5px;
    text-align: left;
}
td { 
    color: #B03726;
    font-family: Tahoma, Verdana, Arial, Sans-Serif;
    font-size: x-small;
    border: solid black 0.5pt;
    padding: 3px;
}
a:link { color: #B03726; }
a:visited { color: #B03726; }
li { color: #B03726; }
strong { color: #B03726; }
body { 
    background-color: #E6E6E6; 
    font-family: Tahoma, Verdana, Arial, Sans-Serif;
    margin: 20px;
}
b { color: gray; }
p { margin: 5px 0; }
.header {
    border-bottom: 2px solid #B03726;
    padding-bottom: 10px;
    margin-bottom: 20px;
}
.logo {
    max-height: 50px;
    margin-bottom: 10px;
}
.info-section {
    margin: 15px 0;
    padding: 10px;
    background-color: white;
    border-radius: 5px;
}
"@ | Out-File -FilePath $cssFile -Encoding UTF8 -Force
        Write-Log "Created default CSS file" -Level Success
    }
}

#endregion

#region Main Script

try {
    Write-Log "===== Network Inventory Started =====" -Level Success
    Write-Log "Inactive threshold: $InactiveDays days" -Level Info
    Write-Log "Orphaned threshold: $OrphanedDays days" -Level Info
    Write-Log "Max concurrent jobs: $MaxConcurrent" -Level Info
    Write-Log "Job timeout: $Timeout seconds" -Level Info
    
    # Initialize environment
    Initialize-Environment
    
    # Query Active Directory for all computer accounts
    Write-Log "Querying Active Directory for computer accounts..." -Level Info
    
    $cutoffDate = (Get-Date).AddDays(-$InactiveDays)
    $orphanedDate = (Get-Date).AddDays(-$OrphanedDays)
    
    # Get all computers with their password last set date
    $allComputers = Get-ADComputer -Filter * -Properties PasswordLastSet, DistinguishedName | 
        Where-Object { $_.PasswordLastSet }
    
    Write-Log "Found $($allComputers.Count) computer accounts in AD" -Level Info
    
    # Categorize computers
    $activeComputers = @()
    $inactiveComputers = @()
    $orphanedComputers = @()
    
    foreach ($computer in $allComputers) {
        $daysSincePasswordSet = ((Get-Date) - $computer.PasswordLastSet).Days
        
        if ($daysSincePasswordSet -gt $OrphanedDays) {
            $orphanedComputers += [PSCustomObject]@{
                Name = $computer.Name
                DistinguishedName = $computer.DistinguishedName
                PasswordLastSet = $computer.PasswordLastSet
                DaysSince = $daysSincePasswordSet
            }
        }
        elseif ($daysSincePasswordSet -gt $InactiveDays) {
            $inactiveComputers += [PSCustomObject]@{
                Name = $computer.Name
                DistinguishedName = $computer.DistinguishedName
                PasswordLastSet = $computer.PasswordLastSet
                DaysSince = $daysSincePasswordSet
            }
        }
        else {
            $activeComputers += [PSCustomObject]@{
                Name = $computer.Name
                DistinguishedName = $computer.DistinguishedName
                PasswordLastSet = $computer.PasswordLastSet
                DaysSince = $daysSincePasswordSet
            }
        }
    }
    
    # Sort lists
    $activeComputers = $activeComputers | Sort-Object Name
    $inactiveComputers = $inactiveComputers | Sort-Object Name
    $orphanedComputers = $orphanedComputers | Sort-Object Name
    
    # Write ComputerAccounts.log
    $computerAccountsLog = Join-Path $OutputPath "ComputerAccounts.log"
    $logContent = @"
Computer Account Status Report
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Inactive Threshold: $InactiveDays days
Orphaned Threshold: $OrphanedDays days

========================================
ORPHANED COMPUTERS (Password > $OrphanedDays days old)
========================================
"@
    
    foreach ($computer in $orphanedComputers) {
        $logContent += "`nOrphaned: $($computer.DistinguishedName) - password last set: $($computer.PasswordLastSet) ($($computer.DaysSince) days ago)"
    }
    
    $logContent += @"

========================================
INACTIVE COMPUTERS (Password > $InactiveDays days old)
========================================
"@
    
    foreach ($computer in $inactiveComputers) {
        $logContent += "`nInactive: $($computer.DistinguishedName) - password last set: $($computer.PasswordLastSet) ($($computer.DaysSince) days ago)"
    }
    
    $logContent += @"

========================================
ACTIVE COMPUTERS (Password < $InactiveDays days old)
========================================
"@
    
    foreach ($computer in $activeComputers) {
        $logContent += "`nActive: $($computer.DistinguishedName) - password last set: $($computer.PasswordLastSet) ($($computer.DaysSince) days ago)"
    }
    
    $logContent += @"

========================================
SUMMARY
========================================
Finished: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Total computer objects found:   $($allComputers.Count)
Active:                         $($activeComputers.Count)
Inactive:                       $($inactiveComputers.Count)
Orphaned:                       $($orphanedComputers.Count)
----------------------------------------------
"@
    
    $logContent | Out-File -FilePath $computerAccountsLog -Encoding UTF8 -Force
    Write-Log "ComputerAccounts.log saved: $computerAccountsLog" -Level Success
    
    # Display summary
    Write-Log "=============================" -Level Success
    Write-Log "Computer objects found: $($allComputers.Count)" -Level Info
    Write-Log "Active:   $($activeComputers.Count)" -Level Success
    Write-Log "Inactive: $($inactiveComputers.Count)" -Level Warning
    Write-Log "Orphaned: $($orphanedComputers.Count)" -Level Error
    Write-Log "=============================" -Level Success
    
    # Exit if SkipInventory is specified
    if ($SkipInventory) {
        Write-Log "SkipInventory flag set - exiting without running inventories" -Level Info
        exit 0
    }
    
    # Run inventory on active computers
    if ($activeComputers.Count -eq 0) {
        Write-Log "No active computers to inventory" -Level Warning
        exit 0
    }
    
    Write-Log "Starting inventory on $($activeComputers.Count) active computers" -Level Info
    
    # Get path to RunInv.ps1
    $runInvScript = Join-Path $PSScriptRoot "RunInv.ps1"
    if (-not (Test-Path $runInvScript)) {
        Write-Log "RunInv.ps1 not found at: $runInvScript" -Level Error
        exit 1
    }
    
    # Job tracking
    $jobs = @{}
    $completed = 0
    $failed = 0
    $timedOut = 0
    $totalComputers = $activeComputers.Count
    
    # Process computers
    $computerQueue = [System.Collections.Queue]::new()
    $activeComputers | ForEach-Object { $computerQueue.Enqueue($_) }
    
    Write-Log "Processing computers with max $MaxConcurrent concurrent jobs..." -Level Info
    
    while ($computerQueue.Count -gt 0 -or $jobs.Count -gt 0) {
        # Start new jobs if under max concurrent
        while ($jobs.Count -lt $MaxConcurrent -and $computerQueue.Count -gt 0) {
            $computer = $computerQueue.Dequeue()
            $computerName = $computer.Name
            
            Write-Log "Starting inventory for: $computerName" -Level Info
            
            # Delete old HTML report if exists
            $htmlReport = Join-Path $OutputPath "$computerName.html"
            if (Test-Path $htmlReport) {
                Remove-Item $htmlReport -Force -ErrorAction SilentlyContinue
            }
            
            # Build command
            $scriptBlockParams = @{
                ComputerName = $computerName
                OutputPath = $OutputPath
                SoftwareXmlPath = $SoftwareXmlPath
                Timeout = $Timeout
            }
            
            if ($IncludeWin32Product) {
                $scriptBlockParams['IncludeWin32Product'] = $true
            }
            
            if ($Credential) {
                $scriptBlockParams['Credential'] = $Credential
            }
            
            # Start job
            $job = Start-Job -ScriptBlock {
                param($ScriptPath, $Params)
                & $ScriptPath @Params
            } -ArgumentList $runInvScript, $scriptBlockParams
            
            $jobs[$job.Id] = @{
                Job = $job
                ComputerName = $computerName
                StartTime = Get-Date
            }
        }
        
        # Check running jobs
        $jobsToRemove = @()
        
        foreach ($jobId in $jobs.Keys) {
            $jobInfo = $jobs[$jobId]
            $job = $jobInfo.Job
            $computerName = $jobInfo.ComputerName
            $startTime = $jobInfo.StartTime
            $elapsed = ((Get-Date) - $startTime).TotalSeconds
            
            # Check if job completed
            if ($job.State -eq 'Completed') {
                $completed++
                $jobsToRemove += $jobId
                Write-Log "Completed: $computerName ($completed/$totalComputers)" -Level Success
                Receive-Job -Job $job | Out-Null
                Remove-Job -Job $job
            }
            # Check if job failed
            elseif ($job.State -eq 'Failed') {
                $failed++
                $jobsToRemove += $jobId
                Write-Log "Failed: $computerName" -Level Error
                Receive-Job -Job $job | Out-Null
                Remove-Job -Job $job
            }
            # Check for timeout
            elseif ($elapsed -gt $Timeout) {
                $timedOut++
                $jobsToRemove += $jobId
                Write-Log "Timeout: $computerName (exceeded $Timeout seconds)" -Level Warning
                Stop-Job -Job $job
                Remove-Job -Job $job
            }
        }
        
        # Remove completed/failed/timed out jobs from tracking
        foreach ($jobId in $jobsToRemove) {
            $jobs.Remove($jobId)
        }
        
        # Update progress
        $percentComplete = [math]::Round(($completed / $totalComputers) * 100, 1)
        Write-Progress -Activity "Running Network Inventory" `
            -Status "Progress: $completed/$totalComputers completed, $($jobs.Count) running, $failed failed, $timedOut timed out" `
            -PercentComplete $percentComplete
        
        # Sleep briefly before checking again
        if ($jobs.Count -gt 0) {
            Start-Sleep -Milliseconds 500
        }
    }
    
    Write-Progress -Activity "Running Network Inventory" -Completed
    
    # Final summary
    $duration = (Get-Date) - $scriptStartTime
    Write-Log "" -Level Info
    Write-Log "===== Network Inventory Complete =====" -Level Success
    Write-Log "Total Duration: $([math]::Round($duration.TotalMinutes, 2)) minutes" -Level Info
    Write-Log "Computers Processed: $totalComputers" -Level Info
    Write-Log "Completed: $completed" -Level Success
    Write-Log "Failed: $failed" -Level $(if ($failed -gt 0) { 'Error' } else { 'Success' })
    Write-Log "Timed Out: $timedOut" -Level $(if ($timedOut -gt 0) { 'Warning' } else { 'Success' })
    Write-Log "=====================================" -Level Success
    Write-Log "Reports location: $OutputPath" -Level Info
    
}
catch {
    Write-Log "Fatal error: $_" -Level Error
    Write-Log $_.ScriptStackTrace -Level Error
    exit 1
}

#endregion