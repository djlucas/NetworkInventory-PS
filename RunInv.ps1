#Ref: RunInv.ps1
#requires -Version 5.1
<#
.SYNOPSIS
    System inventory script using CIM - collects hardware, software, and configuration data
    
.DESCRIPTION
    Collects comprehensive system information and compares against software.xml requirements.
    Can run directly on local computer or be called remotely via PS Remoting.
    
    Data Sources:
    - CIM/WMI classes for hardware and OS info
    - Mapped Storage (Physical -> Logical)
    - WmiMonitorID (root\wmi) for display info
    - Registry (3 sources): Win32_Product, Uninstall (64-bit), Uninstall (32-bit)
    - Win32_QuickFixEngineering for installed patches
    - Win32_GroupUser for local group memberships (AzureAD/Domain)
    - Win32_UserProfile for profile paths
    - Get-AppxProvisionedPackage for Global UWP apps
    - Get-AppxPackage for User Installed UWP apps
    
.PARAMETER ComputerName
    Target computer name. If omitted, runs against local computer.
    
.PARAMETER Credential
    Credentials for remote access (optional)
    
.PARAMETER OutputPath
    Where to save reports. Default: .\Reports
    
.PARAMETER SoftwareXmlPath
    Path to software.xml. Default: .\software.xml
    
.PARAMETER IncludeProcesses
    Include running processes in report
    
.PARAMETER SkipSoftwareCheck
    Skip software version comparison against XML
    
.PARAMETER Timeout
    Execution timeout in seconds. Default: 180 (3 minutes)
    
.EXAMPLE
    .\RunInv.ps1
    Runs inventory on local computer
    
.EXAMPLE
    .\RunInv.ps1 -ComputerName SERVER01
    Runs inventory on remote computer SERVER01
    
.NOTES
    Author: System Admin
    Version: 5.3
    Based on: SYDI-Server 2.4 and runinv.vbs
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ComputerName,
    
    [Parameter(Mandatory=$false)]
    [PSCredential]$Credential,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\Reports",
    
    [Parameter(Mandatory=$false)]
    [string]$SoftwareXmlPath = ".\software.xml",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeProcesses,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipSoftwareCheck,
    
    [Parameter(Mandatory=$false)]
    [int]$Timeout = 180
)

$ErrorActionPreference = 'Stop'
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

function Get-CimData {
    param(
        [string]$ClassName,
        [string]$Namespace = "root\cimv2",
        [string]$Filter = $null,
        [string[]]$Property = $null,
        [object]$Session
    )
    
    try {
        $params = @{
            ClassName = $ClassName
            Namespace = $Namespace
            ErrorAction = 'Stop'
        }
        
        if ($Session) {
            $params['CimSession'] = $Session
        }
        
        if ($Filter) {
            $params['Filter'] = $Filter
        }

        if ($Property) {
            $params['Property'] = $Property
        }
        
        Get-CimInstance @params
    }
    catch {
        Write-Log "Failed to query $ClassName : $_" -Level Warning
        return $null
    }
}

function Get-InstalledSoftware {
    param(
        [string]$TargetComputer,
        [PSCredential]$Cred
    )
    
    $allSoftware = @()
    
    # ScriptBlock to run on target computer (local or remote)
    $scriptBlock = {
        # Helper to parse raw date strings (YYYYMMDD) into readable format
        function Get-NormalizedDate {
            param($DateString)
            if ([string]::IsNullOrWhiteSpace($DateString)) { return "" }
            
            # Try YYYYMMDD
            if ($DateString -match '^\d{8}$') {
                try {
                    return [DateTime]::ParseExact($DateString, "yyyyMMdd", $null).ToString("MM/dd/yyyy")
                } catch {}
            }
            # Try already formatted
            if ($DateString -as [DateTime]) {
                return ([DateTime]$DateString).ToString("MM/dd/yyyy")
            }
            return $DateString # Return raw if parse fails
        }

        $softwareList = @()
        
        # 1. Registry - 64-bit
        $regPath64 = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
        if (Test-Path $regPath64) {
            Get-ItemProperty -Path $regPath64 -ErrorAction SilentlyContinue | 
                Where-Object { $_.DisplayName } | 
                ForEach-Object {
                    $softwareList += [PSCustomObject]@{
                        Name = $_.DisplayName
                        Version = $_.DisplayVersion
                        Publisher = $_.Publisher
                        InstallDate = Get-NormalizedDate $_.InstallDate
                        Source = 'Registry (64-bit)'
                        RawSource = 1
                    }
                }
        }
        
        # 2. Registry - 32-bit (Wow6432Node)
        $regPath32 = "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        if (Test-Path $regPath32) {
            Get-ItemProperty -Path $regPath32 -ErrorAction SilentlyContinue | 
                Where-Object { $_.DisplayName } | 
                ForEach-Object {
                    $softwareList += [PSCustomObject]@{
                        Name = $_.DisplayName
                        Version = $_.DisplayVersion
                        Publisher = $_.Publisher
                        InstallDate = Get-NormalizedDate $_.InstallDate
                        Source = 'Registry (32-bit)'
                        RawSource = 2
                    }
                }
        }
        
        # 3. Win32_Product (MSI)
        try {
            $msi = Get-CimInstance -ClassName Win32_Product -ErrorAction SilentlyContinue
            $softwareList += $msi | ForEach-Object {
                [PSCustomObject]@{
                    Name = $_.Name
                    Version = $_.Version
                    Publisher = $_.Vendor
                    InstallDate = Get-NormalizedDate $_.InstallDate
                    Source = 'MSI'
                    RawSource = 3
                }
            }
        } catch {}
        
        # INTELLIGENT MERGE & DEDUP
        # Group by Name. If duplicates exist:
        # 1. Prefer entry with Date/Vendor info.
        # 2. If identical, prefer Registry over MSI (usually faster/cleaner).
        
        $merged = $softwareList | Group-Object Name | ForEach-Object {
            $group = $_.Group
            if ($group.Count -eq 1) {
                $group[0]
            } else {
                # Logic: Sort by having Vendor (desc), having Date (desc), then Source (Reg < MSI)
                $best = $group | Sort-Object @{e={![string]::IsNullOrEmpty($_.Publisher)}}, @{e={![string]::IsNullOrEmpty($_.InstallDate)}}, RawSource | Select-Object -Last 1
                
                # Append other sources to source string if merging
                $allSources = ($group.Source | Select-Object -Unique) -join ", "
                $best.Source = $allSources
                $best
            }
        }
        
        return $merged | Sort-Object Name
    }
    
    # Execute locally or remotely
    if ([string]::IsNullOrEmpty($TargetComputer) -or $TargetComputer -eq $env:COMPUTERNAME -or $TargetComputer -eq 'localhost') {
        Write-Log "Collecting merged software list from local machine" -Level Info
        $allSoftware = & $scriptBlock
    }
    else {
        Write-Log "Collecting merged software list from remote: $TargetComputer" -Level Info
        $params = @{
            ComputerName = $TargetComputer
            ScriptBlock = $scriptBlock
            ErrorAction = 'Stop'
        }
        if ($Cred) {
            $params['Credential'] = $Cred
        }
        
        $allSoftware = Invoke-Command @params
    }
    
    return $allSoftware
}

function Load-SoftwareXml {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) {
        Write-Log "Software XML file not found: $Path" -Level Warning
        return $null
    }
    
    try {
        [xml]$xml = Get-Content -Path $Path -ErrorAction Stop
        Write-Log "Loaded software.xml successfully" -Level Success
        return $xml
    }
    catch {
        Write-Log "Failed to load software.xml: $_" -Level Error
        return $null
    }
}

function Compare-SoftwareVersion {
    param(
        [string]$InstalledVersion,
        [string]$ExpectedVersion
    )
    
    if ([string]::IsNullOrWhiteSpace($InstalledVersion)) { return $false }
    if ([string]::IsNullOrWhiteSpace($ExpectedVersion)) { return $false }
    
    try {
        $installed = [version]$InstalledVersion
        $expected = [version]$ExpectedVersion
        return $installed -lt $expected
    }
    catch {
        return $InstalledVersion -ne $ExpectedVersion
    }
}

function Compare-InstalledSoftware {
    param(
        [array]$InstalledSoftware,
        [xml]$SoftwareXml
    )
    
    $issues = @()
    if (-not $SoftwareXml) { return $issues }
    
    $searchRules = $SoftwareXml.check.software.search
    
    foreach ($rule in $searchRules) {
        $ruleName = $rule.name
        $expectedVersion = $rule.expected_version
        $literal = $rule.literal -eq 'True'
        $include = $rule.include
        $exclude = $rule.exclude
        $verspec = $rule.verspec
        $security = $rule.Security -eq 'True'
        
        # Find matching installed software
        $matchingSoftware = $InstalledSoftware | Where-Object {
            $match = $false
            
            if ($literal) {
                # Exact match
                $match = $_.Name -eq $ruleName
            }
            else {
                # Contains match
                $match = $_.Name -like "*$ruleName*"
            }
            
            # Apply Include RegEx if exists
            if ($match -and $include) {
                $match = $_.Name -match $include
            }
            
            # Apply Exclude RegEx if exists
            if ($match -and $exclude) {
                $match = $_.Name -notmatch $exclude
            }
            
            $match
        }
        
        foreach ($software in $matchingSoftware) {
            # Apply Version Spec RegEx if exists
            if ($verspec -and $software.Version -notmatch $verspec) {
                continue
            }
            
            $needsUpdate = $false
            
            if ($expectedVersion -like "Deprecated*" -or $expectedVersion -like "Please *") {
                $needsUpdate = $true
            }
            elseif (Compare-SoftwareVersion -InstalledVersion $software.Version -ExpectedVersion $expectedVersion) {
                $needsUpdate = $true
            }
            
            if ($needsUpdate) {
                $issues += [PSCustomObject]@{
                    Software = $software.Name
                    InstalledVersion = $software.Version
                    ExpectedVersion = $expectedVersion
                    Security = $security
                }
            }
        }
    }
    
    return $issues
}

function Get-OSVersionCode {
    param(
        [string]$OSVersion, 
        [string]$OSCaption,
        [array]$ServerFeatures
    )
    
    $versionParts = $OSVersion -split '\.'
    $major = [int]$versionParts[0]
    $minor = if ($versionParts.Count -gt 1) { [int]$versionParts[1] } else { 0 }
    $build = if ($versionParts.Count -gt 2) { [int]$versionParts[2] } else { 0 }
    
    $isServer = $OSCaption -match 'Server'
    $isRT = $OSCaption -match 'RT'
    
    $isCore = $false
    
    if ($isServer -and $ServerFeatures) {
        $hasGui = $false
        foreach ($feat in $ServerFeatures) {
            if ($feat.ID -eq '478' -or $feat.ID -eq '99') {
                $hasGui = $true
                break
            }
        }
        if (-not $hasGui) { $isCore = $true }
    }
    
    $suffix = "F" 
    if ($isRT) { $suffix = "R" }
    elseif ($isServer -and $isCore) { $suffix = "C" }
    elseif ($isServer) { $suffix = "S" }
    
    if ("$major.$minor" -eq "10.0") {
        switch ($build) {
            10240 { return "10.0$suffix" }
            10586 { return "10.1511$suffix" }
            14393 { return "10.1607$suffix" }
            15063 { return "10.1703$suffix" }
            16299 { return "10.1709$suffix" }
            17134 { return "10.1803$suffix" }
            17763 { return "10.1809$suffix" }
            18362 { return "10.1903$suffix" }
            18363 { return "10.1909$suffix" }
            19041 { return "10.2004$suffix" }
            19042 { return "10.20H2$suffix" }
            19043 { return "10.21H1$suffix" }
            19044 { return "10.21H2$suffix" }
            19045 { return "10.22H2$suffix" }
            20348 { return "2022$suffix" }
            22000 { return "11.0$suffix" }
            22621 { return "11.22H2$suffix" }
            22631 { return "11.23H2$suffix" }
            25398 { return "2022.23H2$suffix" }
            26100 { 
                if ($isServer) { return "2025$suffix" } 
                return "11.24H2$suffix" 
            }
            # Explicit map for 26200
            26200 {
                if ($isServer) { return "2025$suffix" }
                return "11.25H2$suffix"
            }
            { $_ -gt 26100 } {
                if ($isServer) { return "2025$suffix" }
                return "11.25H2$suffix"
            }
            default { return "10$suffix" }
        }
    }
    
    switch ("$major.$minor") {
        "6.0" { return "6.0$suffix" }
        "6.1" { return "6.1$suffix" }
        "6.2" { return "6.2$suffix" }
        "6.3" { return "6.3$suffix" }
    }
    
    return "Unknown"
}

function Compare-InstalledPatches {
    param(
        [array]$InstalledPatches,
        [xml]$SoftwareXml,
        [string]$OSVersionCode
    )
    
    $issues = @()
    if (-not $SoftwareXml) { return $issues }
    
    $updateRules = $SoftwareXml.check.updates.os_update | Where-Object { $_.os -match $OSVersionCode }
    
    foreach ($rule in $updateRules) {
        $kbid = $rule.kbid
        $desc = $rule.desc
        $kbidClean = $kbid.Replace("xx", "")
        
        $installed = $InstalledPatches | Where-Object { $_.HotFixID -like "*$kbidClean*" }
        
        if (-not $installed) {
            $issues += [PSCustomObject]@{
                KBID = $kbidClean
                Description = $desc
            }
        }
    }
    
    return $issues
}

function Write-ToFileWithLock {
    param(
        [string]$FilePath,
        [string]$Content,
        [int]$TimeoutSeconds = 30
    )
    
    $startTime = Get-Date
    $lockAcquired = $false
    $lastError = $null
    
    # Ensure Directory Exists (Absolute)
    $directory = Split-Path -Path $FilePath -Parent
    if (-not (Test-Path $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }
    
    while (-not $lockAcquired -and ((Get-Date) - $startTime).TotalSeconds -lt $TimeoutSeconds) {
        try {
            # Use System.IO.File for precise control, but use absolute path
            $fileStream = [System.IO.File]::Open(
                $FilePath,
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::Read
            )
            
            $lockAcquired = $true
            
            $writer = New-Object System.IO.StreamWriter($fileStream)
            $writer.WriteLine($Content)
            $writer.Flush()
            $writer.Close()
            $fileStream.Close()
            
            Write-Log "Successfully wrote to $FilePath" -Level Success
            return $true
        }
        catch {
            $lastError = $_
            if ($fileStream) {
                $fileStream.Close()
            }
            Start-Sleep -Milliseconds 500
        }
    }
    
    Write-Log "Failed to acquire file lock for $FilePath. Last Error: $($lastError.Exception.Message)" -Level Error
    return $false
}

function Get-ChassisType {
    param([int]$TypeCode)
    
    $types = @{
        1 = "Other"; 2 = "Unknown"; 3 = "Desktop"; 4 = "Low Profile Desktop"
        5 = "Pizza Box"; 6 = "Mini Tower"; 7 = "Tower"; 8 = "Portable"
        9 = "Laptop"; 10 = "Notebook"; 11 = "Hand Held"; 12 = "Docking Station"
        13 = "All in One"; 14 = "Sub Notebook"; 15 = "Space-Saving"
        16 = "Lunch Box"; 17 = "Main System Chassis"; 18 = "Expansion Chassis"
        19 = "SubChassis"; 20 = "Bus Expansion Chassis"; 21 = "Peripheral Chassis"
        22 = "Storage Chassis"; 23 = "Rack Mount Chassis"; 24 = "Sealed-Case PC"
    }
    
    if ($types.ContainsKey($TypeCode)) {
        return $types[$TypeCode]
    }
    return "Unknown"
}

function Get-FormFactor {
    param(
        [int]$FactorCode,
        [int]$MemoryTypeCode,
        [string]$ChassisType
    )
    
    # If BIOS says 0 (Unknown) or 1 (Other), attempt heuristic inference
    if ($FactorCode -le 1) {
        # Heuristic 1: LPDDR (Soldered)
        # 27-30 (LPDDR3/4/4X/5) or 35 (LPDDR5X)
        if ($MemoryTypeCode -in 27, 28, 29, 30, 35) {
            return "Row of Chips (Soldered)"
        }
        
        # Heuristic 2: Chassis based inference
        if ($ChassisType -match "Notebook|Laptop|Portable|SubNotebook|Hand Held") {
            return "SODIMM / Soldered"
        } elseif ($ChassisType -match "Desktop|Tower") {
            return "DIMM"
        }
    }

    $factors = @{
        0 = "Unknown"; 1 = "Other"; 2 = "SIP"; 3 = "DIP"; 4 = "ZIP"; 5 = "SOJ"
        6 = "Proprietary"; 7 = "SIMM"; 8 = "DIMM"; 9 = "TSOP"; 10 = "PGA"
        11 = "RIMM"; 12 = "SODIMM"; 13 = "SRIMM"; 14 = "SMD"; 15 = "SSMP"
        16 = "QFP"; 17 = "TQFP"; 18 = "SOIC"; 19 = "LCC"; 20 = "PLCC"
        21 = "BGA"; 22 = "FPBGA"; 23 = "LGA"
    }
    
    if ($factors.ContainsKey($FactorCode)) { return $factors[$FactorCode] }
    return "Unknown"
}

function Get-MemoryType {
    param([int]$TypeCode)
    $types = @{
        0 = "Unknown"; 1 = "Other"; 2 = "DRAM"; 3 = "Synchronous DRAM"
        4 = "Cache DRAM"; 5 = "EDO"; 6 = "EDRAM"; 7 = "VRAM"; 8 = "SRAM"
        9 = "RAM"; 10 = "ROM"; 11 = "Flash"; 12 = "EEPROM"; 13 = "FEPROM"
        14 = "EPROM"; 15 = "CDRAM"; 16 = "3DRAM"; 17 = "SDRAM"; 18 = "SGRAM"
        19 = "RDRAM"; 20 = "DDR"; 21 = "DDR2"; 22 = "DDR2 FB-DIMM"
        24 = "DDR3"; 26 = "DDR4"
        27 = "LPDDR3"; 28 = "LPDDR4"; 29 = "LPDDR4X"; 30 = "LPDDR5"
        34 = "DDR5"; 35 = "LPDDR5X"
    }
    if ($types.ContainsKey($TypeCode)) { return $types[$TypeCode] }
    return "Unknown"
}

function Get-BiosCharacteristic {
    param([int]$Code)
    switch ($Code) {
        0 { "Reserved" }
        1 { "Reserved" }
        2 { "Unknown" }
        3 { "BIOS Characteristics Not Supported" }
        4 { "ISA is supported" }
        5 { "MCA is supported" }
        6 { "EISA is supported" }
        7 { "PCI is supported" }
        8 { "PC Card (PCMCIA) is supported" }
        9 { "Plug and Play is supported" }
        10 { "APM is supported" }
        11 { "BIOS is Upgradable (Flash)" }
        12 { "BIOS shadowing is allowed" }
        13 { "VL-VESA is supported" }
        14 { "ESCD support is available" }
        15 { "Boot from CD is supported" }
        16 { "Selectable Boot is supported" }
        17 { "BIOS ROM is socketed" }
        18 { "Boot From PC Card (PCMCIA) is supported" }
        19 { "EDD (Enhanced Disk Drive) Specification is supported" }
        20 { "Int 13h - Japanese Floppy for NEC 9800 1.2mb (3.5, 1k Bytes/Sector, 360 RPM) is supported" }
        21 { "Int 13h - Japanese Floppy for Toshiba 1.2mb (3.5, 360 RPM) is supported" }
        22 { "Int 13h - 5.25 / 360 KB Floppy Services are supported" }
        23 { "Int 13h - 5.25 /1.2MB Floppy Services are supported" }
        24 { "Int 13h - 3.5 / 720 KB Floppy Services are supported" }
        25 { "Int 13h - 3.5 / 2.88 MB Floppy Services are supported" }
        26 { "Int 5h, Print Screen Service is supported" }
        27 { "Int 9h, 8042 Keyboard services are supported" }
        28 { "Int 14h, Serial Services are supported" }
        29 { "Int 17h, printer services are supported" }
        30 { "Int 10h, CGA/Mono Video Services are supported" }
        31 { "NEC PC-98" }
        32 { "ACPI supported" }
        33 { "USB Legacy is supported" }
        34 { "AGP is supported" }
        35 { "I2O boot is supported" }
        36 { "LS-120 boot is supported" }
        37 { "ATAPI ZIP Drive boot is supported" }
        38 { "1394 boot is supported" }
        39 { "Smart Battery supported" }
        Default { "Unknown (Undocumented)" }
    }
}

#endregion

#region Main Script

try {
    Write-Log "===== Starting System Inventory =====" -Level Info
    
    # Determine target computer
    if ([string]::IsNullOrEmpty($ComputerName)) {
        $ComputerName = $env:COMPUTERNAME
        Write-Log "No target specified, running on local computer: $ComputerName" -Level Info
    }
    else {
        Write-Log "Target computer: $ComputerName" -Level Info
    }
    
    # Resolve OutputPath to Absolute Path
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-Log "Created output directory: $OutputPath" -Level Success
    }
    $OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    Write-Log "Resolved Output Path: $OutputPath" -Level Info
    
    # Create CIM Session and PS Session if targeting remote computer
    $cimSession = $null
    $psSession = $null  # Added PSSession variable
    $isRemote = $ComputerName -ne $env:COMPUTERNAME -and $ComputerName -ne 'localhost'
    
    if ($isRemote) {
        $cimParams = @{
            ComputerName = $ComputerName
            ErrorAction = 'Stop'
        }
        
        $psParams = @{
            ComputerName = $ComputerName
            ErrorAction = 'Stop'
        }

        if ($Credential) {
            $cimParams['Credential'] = $Credential
            $psParams['Credential'] = $Credential
        }
        
        Write-Log "Establishing CIM and PS sessions to $ComputerName" -Level Info
        $cimSession = New-CimSession @cimParams
        $psSession = New-PSSession @psParams  # Initialize standard PS Remoting session
    }
    
    # Initialize inventory data
    $inventory = @{
        ScanTime = Get-Date -Format "M/d/yyyy h:mm:ss tt"
        ComputerSystem = $null
        OperatingSystem = $null
        BIOS = $null
        SystemProduct = $null
        SystemEnclosure = $null
        TimeZone = $null
        Processor = $null
        ProcessorCount = 0
        PhysicalMemory = @()
        PhysicalDisks = @()
        VideoController = @()
        LogicalDisks = @()
        NetworkAdapters = @()
        Monitors = @()
        Patches = @()
        Software = @()
        ProvisionedApps = @()
        UserApps = @()
        UserProfiles = @()
        Services = @()
        Shares = @()
        Printers = @()
        EventLogs = @()
        StartupCommands = @()
        PageFiles = @()
        Registry = $null
        Processes = @()
        ServerFeatures = @()
        Roles = @()
        SoftwareIssues = @()
        PatchIssues = @()
        LocalGroups = @()
        LocalUsers = @()
        LastLoggedOnUser = "Unknown"
        PowerShellVersion = "Unknown"
    }
    
    # Collect Operating System
    Write-Log "Collecting OS information" -Level Info
    $os = Get-CimData -ClassName "Win32_OperatingSystem" -Session $cimSession | Select-Object -First 1
    if ($os) {
        $inventory.OperatingSystem = [PSCustomObject]@{
            Caption = $os.Caption
            Version = $os.Version
            BuildNumber = $os.BuildNumber
            OSLanguage = switch ($os.OSLanguage) { 1033 { "English" } default { $os.OSLanguage } }
            InstallDate = $os.InstallDate
            WindowsDirectory = $os.WindowsDirectory
            FreePhysicalMemory = $os.FreePhysicalMemory
            FreePhysicalMemoryMB = [math]::Round($os.FreePhysicalMemory / 1024, 0)
        }
    }
    
    # Collect Computer System
    Write-Log "Collecting computer system information" -Level Info
    
    # FIX: Explicitly select properties to avoid "Invalid XML" error caused by large OEMLogoBitmap
    $csProps = @("Name", "Domain", "DomainRole", "TotalPhysicalMemory")
    $cs = Get-CimData -ClassName "Win32_ComputerSystem" -Property $csProps -Session $cimSession | Select-Object -First 1
    
    if ($cs) {
        $domainRole = switch ($cs.DomainRole) {
            0 { "Standalone Workstation" }
            1 { "Member Workstation" }
            2 { "Standalone Server" }
            3 { "Member Server" }
            4 { "Backup Domain Controller" }
            5 { "Primary Domain Controller" }
            default { "Unknown" }
        }
        
        $fqdn = try { "$($cs.Name).$($cs.Domain)" } catch { $cs.Name }
        $domainType = if ($cs.DomainRole -lt 2) { "workgroup" } else { "domain" }

        $inventory.ComputerSystem = [PSCustomObject]@{
            Name = $cs.Name
            FQDN = $fqdn
            Domain = $cs.Domain
            DomainRole = $domainRole
            DomainRoleCode = $cs.DomainRole
            TotalPhysicalMemory = $cs.TotalPhysicalMemory
            TotalPhysicalMemoryMB = [math]::Round($cs.TotalPhysicalMemory / 1MB, 0)
            OSConfigString = "$domainRole in the $($cs.Domain) $domainType"
        }
        
        # Add Roles (DC or Workstation)
        if ($cs.DomainRole -ge 4) { 
            $inventory.Roles += "DC" 
        }
        elseif ($cs.DomainRole -lt 2) {
            # FIX: Explicitly label workstations in the Roles list
            $inventory.Roles += $domainRole
        }
    }
    
    # AD Description Lookup
    $adDescription = "N/A"
    # Add check for null ComputerSystem before accessing properties
    if ($inventory.ComputerSystem -and $inventory.ComputerSystem.DomainRoleCode -ne 0 -and $inventory.ComputerSystem.DomainRoleCode -ne 2) {
        try {
            Write-Log "Querying Active Directory for description" -Level Info
            $searcher = [adsisearcher]"(cn=$($inventory.ComputerSystem.Name))"
            $searcher.PropertiesToLoad.Add("description") | Out-Null
            $res = $searcher.FindOne()
            if ($res -and $res.Properties["description"]) {
                $adDescription = $res.Properties["description"][0]
            }
        } catch {
            Write-Log "Failed to query AD description: $_" -Level Warning
        }
    }
    
    # Collect BIOS
    Write-Log "Collecting BIOS information" -Level Info
    $bios = Get-CimData -ClassName "Win32_BIOS" -Session $cimSession | Select-Object -First 1
    if ($bios) {
        $inventory.BIOS = [PSCustomObject]@{
            SMBIOSBIOSVersion = $bios.SMBIOSBIOSVersion
            SMBIOSMajorVersion = $bios.SMBIOSMajorVersion
            SMBIOSMinorVersion = $bios.SMBIOSMinorVersion
            Version = $bios.Version
            BiosCharacteristics = $bios.BiosCharacteristics
        }
    }
    
    # Collect System Product
    Write-Log "Collecting system product information" -Level Info
    $product = Get-CimData -ClassName "Win32_ComputerSystemProduct" -Session $cimSession | Select-Object -First 1
    if ($product) {
        $inventory.SystemProduct = [PSCustomObject]@{
            Vendor = $product.Vendor
            Name = $product.Name
            IdentifyingNumber = $product.IdentifyingNumber
        }
    }
    
    # Collect System Enclosure
    Write-Log "Collecting system enclosure" -Level Info
    $enclosure = Get-CimData -ClassName "Win32_SystemEnclosure" -Session $cimSession | Select-Object -First 1
    if ($enclosure) {
        $chassisType = if ($enclosure.ChassisTypes) { Get-ChassisType -TypeCode ([int]$enclosure.ChassisTypes[0]) } else { "Unknown" }
        $inventory.SystemEnclosure = [PSCustomObject]@{
            ChassisType = $chassisType
        }
    }
    
    # Collect Processor
    Write-Log "Collecting processor information" -Level Info
    $processors = Get-CimData -ClassName "Win32_Processor" -Session $cimSession
    if ($processors) {
        $processor = $processors | Select-Object -First 1
        $inventory.Processor = [PSCustomObject]@{
            Name = $processor.Name
            Description = $processor.Description
            MaxClockSpeed = $processor.MaxClockSpeed
            L2CacheSize = $processor.L2CacheSize
            ExtClock = $processor.ExtClock
        }
        $inventory.ProcessorCount = ($processors | Measure-Object).Count
    }
    
    # Collect Physical Memory (Enhanced)
    Write-Log "Collecting memory information" -Level Info
    $memory = Get-CimData -ClassName "Win32_PhysicalMemory" -Session $cimSession
    if ($memory) {
        $inventory.PhysicalMemory = $memory | ForEach-Object {
            $typeCode = if ($_.SMBIOSMemoryType -and $_.SMBIOSMemoryType -ne 0) { $_.SMBIOSMemoryType } else { $_.MemoryType }
            $ffCode = $_.FormFactor
            $formFactorStr = Get-FormFactor -FactorCode ([int]$ffCode) -MemoryTypeCode ([int]$typeCode) -ChassisType $inventory.SystemEnclosure.ChassisType
            $serial = $_.SerialNumber
            if ([string]::IsNullOrWhiteSpace($serial) -or $serial -eq "00000000") { $serial = "N/A" }

            [PSCustomObject]@{
                BankLabel = $_.BankLabel
                Capacity = $_.Capacity
                FormFactor = $formFactorStr
                MemoryType = Get-MemoryType -TypeCode ([int]$typeCode)
                Manufacturer = $_.Manufacturer
                PartNumber = $_.PartNumber
                SerialNumber = $serial
            }
        }
    }
    
    # Collect Storage (Physical -> Logical Map)
    Write-Log "Collecting storage information" -Level Info
    $scriptBlock_Storage = {
        $storageData = @()
        $phyDisks = Get-CimInstance Win32_DiskDrive
        $msftDisks = try { Get-CimInstance -Namespace root\microsoft\windows\storage -ClassName MSFT_PhysicalDisk -ErrorAction SilentlyContinue } catch { $null }

        foreach ($disk in $phyDisks) {
            $mediaType = "Unknown"
            $busType = $disk.InterfaceType
            
            if ($msftDisks) {
                $match = $msftDisks | Where-Object { $_.SerialNumber -eq $disk.SerialNumber }
                if (-not $match) { $match = $msftDisks | Where-Object { $_.FriendlyName -eq $disk.Model } }
                
                if ($match) {
                    $mediaType = switch($match.MediaType) { 3 {"HDD"} 4 {"SSD"} 5 {"SCM"} default {"Unspecified"} }
                    $busType = switch($match.BusType) { 7 {"USB"} 11 {"SATA"} 17 {"NVMe"} default {$match.BusType} }
                }
            }

            $partitions = Get-CimAssociatedInstance -InputObject $disk -ResultClassName Win32_DiskPartition
            $logicalDisks = @()
            
            foreach ($part in $partitions) {
                $ld = Get-CimAssociatedInstance -InputObject $part -ResultClassName Win32_LogicalDisk
                if ($ld) {
                    $logicalDisks += $ld | Select-Object DeviceID, VolumeName, Size, FreeSpace, FileSystem
                }
            }
            
            $storageData += [PSCustomObject]@{
                Model = $disk.Model
                DeviceID = $disk.DeviceID
                Interface = $busType
                MediaType = $mediaType
                Size = $disk.Size
                SerialNumber = $disk.SerialNumber
                Health = $disk.Status
                Partitions = $logicalDisks
            }
        }
        return $storageData
    }
    
    if ($isRemote) {
        # FIX: Use PSSession for generic Invoke-Command, not CimSession
        $inventory.PhysicalDisks = Invoke-Command -Session $psSession -ScriptBlock $scriptBlock_Storage
    } else {
        $inventory.PhysicalDisks = & $scriptBlock_Storage
    }
    
    # Collect Video Controller
    Write-Log "Collecting video controller information" -Level Info
    $video = Get-CimData -ClassName "Win32_VideoController" -Session $cimSession
    if ($video) {
        $inventory.VideoController = $video | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                AdapterRAM = $_.AdapterRAM
                AdapterCompatibility = $_.AdapterCompatibility
            }
        }
    }
    
    # Collect Network Adapters (Enhanced)
    Write-Log "Collecting comprehensive network adapter information" -Level Info
    $scriptBlock_Network = {
        $adapters = Get-CimInstance -ClassName Win32_NetworkAdapter
        $configs = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration
        $result = @()
        
        foreach ($nic in $adapters) {
            $conf = $configs | Where-Object { $_.Index -eq $nic.Index }
            
            $type = "Unknown"
            if ($nic.PhysicalAdapter) {
                $type = "Physical"
            } elseif ($nic.PNPDeviceID -match "^PCI") {
                $type = "Physical (Legacy Detect)"
            } elseif ($nic.PNPDeviceID -match "^USB") {
                $type = "Physical (USB)"
            } else {
                $type = "Virtual"
            }
            
            $status = switch ($nic.NetConnectionStatus) {
                0 {"Disconnected"} 1 {"Connecting"} 2 {"Connected"} 3 {"Disconnecting"}
                4 {"Hardware Not Present"} 5 {"Hardware Disabled"} 6 {"Hardware Malfunction"}
                7 {"Media Disconnected"} 8 {"Authenticating"} 9 {"Authentication Succeeded"}
                10 {"Authentication Failed"} 11 {"Invalid Address"} 12 {"Credentials Required"}
                Default {$nic.NetConnectionStatus}
            }
            if ([string]::IsNullOrEmpty($status)) { $status = "Disabled/Not Present" }
            
            $ips = $null
            $gateways = $null
            $dhcp = $false
            
            if ($conf -and $conf.IPEnabled) {
                $ips = $conf.IPAddress
                $gateways = $conf.DefaultIPGateway
                $dhcp = $conf.DHCPEnabled
            }

            $result += [PSCustomObject]@{
                Name = $nic.Name
                Description = $nic.Description
                MACAddress = $nic.MACAddress
                Type = $type
                Status = $status
                IPAddress = $ips
                Gateway = $gateways
                DHCPEnabled = $dhcp
            }
        }
        return $result
    }
    
    if ($isRemote) {
        # FIX: Use PSSession
        $inventory.NetworkAdapters = Invoke-Command -Session $psSession -ScriptBlock $scriptBlock_Network
    } else {
        $inventory.NetworkAdapters = & $scriptBlock_Network
    }
    
    # Collect Monitors (WMI)
    Write-Log "Collecting monitor information" -Level Info
    $scriptBlock_Monitors = {
        function Decode-WmiMonitorString {
            param($IntArray)
            if (-not $IntArray) { return "" }
            $str = ""
            foreach ($char in $IntArray) { if ($char -gt 0) { $str += [char]$char } }
            return $str
        }

        $mons = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -ErrorAction SilentlyContinue
        $conns = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorConnectionParams -ErrorAction SilentlyContinue
        
        $monList = @()
        foreach ($m in $mons) {
            $name = Decode-WmiMonitorString $m.UserFriendlyName
            $serial = Decode-WmiMonitorString $m.SerialNumberID
            $manuf = Decode-WmiMonitorString $m.ManufacturerName
            
            $connType = "Unknown"
            $c = $conns | Where-Object { $_.InstanceName -eq $m.InstanceName }
            if ($c) {
                $connType = switch ($c.VideoOutputTechnology) {
                    0 {"Other"} 1 {"HD15 (VGA)"} 2 {"S-Video"} 3 {"Composite"} 4 {"DVI"} 5 {"HDMI"}
                    6 {"LVDS"} 7 {"D-Japan"} 8 {"SDI"} 9 {"DisplayPort"} 10 {"DisplayPort Embedded"} 11 {"UDI"}
                    12 {"Enterprise Digital"} 13 {"Thunderbolt"} 14 {"Miracast"} 
                    2147483648 {"Internal"}
                    Default {"Unknown ($($c.VideoOutputTechnology))"}
                }
            }
            
            $monList += [PSCustomObject]@{
                Manufacturer = $manuf
                Model = $name
                SerialNumber = $serial
                ConnectionType = $connType
            }
        }
        return $monList
    }

    if ($isRemote) {
        # FIX: Use PSSession
        $inventory.Monitors = Invoke-Command -Session $psSession -ScriptBlock $scriptBlock_Monitors
    } else {
        $inventory.Monitors = & $scriptBlock_Monitors
    }
    
    # Collect TimeZone
    Write-Log "Collecting timezone information" -Level Info
    $tz = Get-CimData -ClassName "Win32_TimeZone" -Session $cimSession | Select-Object -First 1
    if ($tz) {
        $inventory.TimeZone = $tz.Description
    }
    
    # Collect Server Features
    if ($inventory.OperatingSystem.Caption -match "Server") {
        Write-Log "Collecting server features" -Level Info
        $features = Get-CimData -ClassName "Win32_ServerFeature" -Session $cimSession
        if ($features) {
            $inventory.ServerFeatures = $features | ForEach-Object {
                [PSCustomObject]@{
                    ParentID = $_.ParentID
                    ID = $_.ID
                    Name = $_.Name
                }
            }
            
            if ($features | Where-Object { $_.Name -like "*File*" }) { $inventory.Roles += "File" }
            if ($features | Where-Object { $_.Name -like "*Print*" }) { $inventory.Roles += "Print" }
            if ($features | Where-Object { $_.Name -like "*DNS*" }) { $inventory.Roles += "DNS" }
            if ($features | Where-Object { $_.Name -like "*DHCP*" }) { $inventory.Roles += "DHCP" }
        }
    }
    
    # Collect Patches
    Write-Log "Collecting installed patches" -Level Info
    $patches = Get-CimData -ClassName "Win32_QuickFixEngineering" -Session $cimSession
    if ($patches) {
        $inventory.Patches = $patches | Where-Object { $_.HotFixID -notlike "File*" } | ForEach-Object {
            [PSCustomObject]@{
                HotFixID = $_.HotFixID
                Description = $_.Description
                InstalledOn = $_.InstalledOn
            }
        }
    }
    
    # Collect Software (Merged & Normalized)
    Write-Log "Collecting installed software (Registry + MSI)" -Level Info
    $inventory.Software = Get-InstalledSoftware -TargetComputer $ComputerName -Cred $Credential
    
    if ($inventory.Software | Where-Object { $_.Name -like "*SQL Server*" }) {
        $inventory.Roles += "SQL"
    }
    
    # Collect UWP Apps (Conditional)
    Write-Log "Collecting UWP Apps" -Level Info
    $scriptBlock_UWP = {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = [Security.Principal.WindowsPrincipal]$identity
        $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        $provApps = if ($isAdmin) {
            Get-AppxProvisionedPackage -Online | Select-Object DisplayName, Version
        } else {
             @()
        }

        $userApps = if ($isAdmin) {
            Get-AppxPackage -AllUsers | Where-Object { -not $_.IsFramework -and -not $_.IsResourcePackage } | Select-Object Name, Version, Publisher | Sort-Object Name -Unique
        } else {
            Get-AppxPackage | Where-Object { -not $_.IsFramework -and -not $_.IsResourcePackage } | Select-Object Name, Version, Publisher | Sort-Object Name -Unique
        }
        
        return @{ Provisioned = $provApps; User = $userApps }
    }
    
    try {
        $uwpResult = if ($isRemote) {
            # FIX: Use PSSession
            Invoke-Command -Session $psSession -ScriptBlock $scriptBlock_UWP
        } else {
            & $scriptBlock_UWP
        }
        $inventory.ProvisionedApps = $uwpResult.Provisioned
        $inventory.UserApps = $uwpResult.User
    } catch {
        Write-Log "Failed to collect UWP Apps: $_" -Level Warning
    }

    # Collect User Profiles (Instant - No Size)
    Write-Log "Collecting User Profiles" -Level Info
    $scriptBlock_Profiles = {
        $profiles = Get-CimInstance Win32_UserProfile | Where-Object { -not $_.Special }
        $result = @()
        
        foreach ($prof in $profiles) {
            # Resolve SID to User Name
            $userName = try { 
                $objSID = New-Object System.Security.Principal.SecurityIdentifier($prof.SID)
                $objUser = $objSID.Translate([System.Security.Principal.NTAccount])
                $objUser.Value
            } catch { $prof.SID }
            
            $result += [PSCustomObject]@{
                User = $userName
                SID = $prof.SID
                Path = $prof.LocalPath
            }
        }
        return $result
    }
    
    if ($isRemote) {
        # FIX: Use PSSession
        $inventory.UserProfiles = Invoke-Command -Session $psSession -ScriptBlock $scriptBlock_Profiles
    } else {
        $inventory.UserProfiles = & $scriptBlock_Profiles
    }

    # Collect Services
    Write-Log "Collecting services" -Level Info
    $services = Get-CimData -ClassName "Win32_Service" -Session $cimSession
    if ($services) {
        $inventory.Services = $services | ForEach-Object {
            [PSCustomObject]@{
                Caption = $_.Caption
                StartMode = $_.StartMode
                Started = $_.Started
                StartName = $_.StartName
            }
        }
    }
    
    # Collect Shares
    Write-Log "Collecting shares" -Level Info
    $shares = Get-CimData -ClassName "Win32_Share" -Session $cimSession
    if ($shares) {
        $inventory.Shares = $shares | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Path = $_.Path
                Description = $_.Description
            }
        }
    }
    
    # Collect Printers
    Write-Log "Collecting printers" -Level Info
    $printers = Get-CimData -ClassName "Win32_Printer" -Session $cimSession
    if ($printers) {
        $inventory.Printers = $printers | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                PortName = $_.PortName
                DriverName = $_.DriverName
            }
        }
    }
    
    # Collect Event Logs
    Write-Log "Collecting event log configuration" -Level Info
    $eventLogs = Get-CimData -ClassName "Win32_NTEventLogFile" -Session $cimSession
    if ($eventLogs) {
        $inventory.EventLogs = $eventLogs | ForEach-Object {
            $overwritePolicy = switch ($_.OverwritePolicy) {
                "OverwriteAsNeeded" { "Overwrite as needed" }
                "OverwriteOutdated" { "Overwrite events older than x days" }
                "Never" { "Do not overwrite" }
                default { $_.OverwritePolicy }
            }
            [PSCustomObject]@{
                LogfileName = $_.LogfileName
                MaxFileSize = [math]::Round($_.MaxFileSize / 1KB, 0)
                OverwritePolicy = $overwritePolicy
            }
        }
    }
    
    # Collect Startup Commands
    Write-Log "Collecting startup commands" -Level Info
    $startup = Get-CimData -ClassName "Win32_StartupCommand" -Session $cimSession
    if ($startup) {
        $inventory.StartupCommands = $startup | ForEach-Object {
            [PSCustomObject]@{
                User = $_.User
                Name = $_.Name
                Command = $_.Command
            }
        }
    }
    
    # Collect Page Files
    Write-Log "Collecting page file configuration" -Level Info
    $pagefiles = Get-CimData -ClassName "Win32_PageFileSetting" -Session $cimSession
    if ($pagefiles) {
        $inventory.PageFiles = $pagefiles | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                InitialSize = $_.InitialSize
                MaximumSize = $_.MaximumSize
            }
        }
    }
    
    # Collect Registry Info
    Write-Log "Collecting registry information" -Level Info
    $registry = Get-CimData -ClassName "Win32_Registry" -Session $cimSession | Select-Object -First 1
    if ($registry) {
        $inventory.Registry = [PSCustomObject]@{
            CurrentSize = $registry.CurrentSize
            MaximumSize = $registry.MaximumSize
        }
    }
    
    # Collect Extra Registry Values
    $logonKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" -ErrorAction SilentlyContinue
    if ($logonKey) { $inventory.LastLoggedOnUser = $logonKey.LastLoggedOnUser }
    
    $psKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine" -ErrorAction SilentlyContinue
    if ($psKey) { $inventory.PowerShellVersion = $psKey.PowerShellVersion }
    
    # Collect Local Users and Groups (Non-DC)
    # Add check for null ComputerSystem
    if ($inventory.ComputerSystem -and $inventory.ComputerSystem.DomainRoleCode -lt 4) {
        Write-Log "Collecting local users and groups" -Level Info
        $users = Get-CimData -ClassName "Win32_UserAccount" -Filter "LocalAccount=True" -Session $cimSession
        if ($users) {
            $inventory.LocalUsers = $users | ForEach-Object {
                [PSCustomObject]@{
                    Name = $_.Name
                    Description = $_.Description
                }
            }
        }
        
        # Enhanced Local Group Collection (Members)
        $groups = Get-CimData -ClassName "Win32_Group" -Filter "LocalAccount=True" -Session $cimSession
        if ($groups) {
             $groupList = @()
             foreach ($grp in $groups) {
                 $memberNames = @()
                 try {
                     $members = Get-CimAssociatedInstance -InputObject $grp -ResultClassName Win32_Account -ErrorAction SilentlyContinue
                     if ($members) {
                         $memberNames = $members | ForEach-Object { "$($_.Domain)\$($_.Name)" }
                     }
                 } catch {}
                 
                 $groupList += [PSCustomObject]@{
                     Name = $grp.Name
                     Description = $grp.Description
                     Members = $memberNames
                 }
             }
             $inventory.LocalGroups = $groupList
        }
    }
    
    # Collect Processes if requested
    if ($IncludeProcesses) {
        Write-Log "Collecting running processes" -Level Info
        $processes = Get-CimData -ClassName "Win32_Process" -Session $cimSession
        if ($processes) {
            $inventory.Processes = $processes | ForEach-Object {
                [PSCustomObject]@{
                    Caption = $_.Caption
                    ExecutablePath = $_.ExecutablePath
                }
            }
        }
    }
    
    # Software Version Checking
    if (-not $SkipSoftwareCheck) {
        Write-Log "Loading software requirements XML" -Level Info
        $softwareXml = Load-SoftwareXml -Path $SoftwareXmlPath
        
        if ($softwareXml) {
            Write-Log "Comparing installed software against requirements" -Level Info
            $inventory.SoftwareIssues = Compare-InstalledSoftware -InstalledSoftware $inventory.Software -SoftwareXml $softwareXml
            
            if ($inventory.OperatingSystem) {
                $osVersionCode = Get-OSVersionCode -OSVersion $inventory.OperatingSystem.Version -OSCaption $inventory.OperatingSystem.Caption -ServerFeatures $inventory.ServerFeatures
                Write-Log "OS Version Code: $osVersionCode" -Level Info
                
                Write-Log "Checking for missing patches" -Level Info
                $inventory.PatchIssues = Compare-InstalledPatches -InstalledPatches $inventory.Patches -SoftwareXml $softwareXml -OSVersionCode $osVersionCode
            }
        }
    }
    
    # Generate HTML Report
    Write-Log "Generating HTML report" -Level Info
    $htmlFile = Join-Path $OutputPath "$ComputerName.html"
    
    # --- CSS and Logo Management ---
    $destCssPath = Join-Path $OutputPath "style.css"
    
    if (-not (Test-Path $destCssPath)) {
        # Check running directory (PSScriptRoot)
        $sourceCssPath = Join-Path $PSScriptRoot "style.css"
        
        if (Test-Path $sourceCssPath) {
            Write-Log "Found style.css in script directory. Copying to output..." -Level Info
            Copy-Item -Path $sourceCssPath -Destination $destCssPath -Force
            
            # Check for logo asset in CSS
            try {
                $cssContent = Get-Content -Path $sourceCssPath -Raw
                # Regex to find url('filename.ext')
                if ($cssContent -match "url\(['`"]?([^'`")]+)['`"]?\)") {
                    $logoFileName = $matches[1]
                    $sourceLogoPath = Join-Path $PSScriptRoot $logoFileName
                    
                    if (Test-Path $sourceLogoPath) {
                        Write-Log "Found logo asset '$logoFileName' in CSS. Copying..." -Level Info
                        Copy-Item -Path $sourceLogoPath -Destination (Join-Path $OutputPath $logoFileName) -Force
                    }
                }
            }
            catch {
                Write-Log "Error processing CSS assets: $_" -Level Warning
            }
        }
    }

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("<!DOCTYPE html>")
    [void]$sb.AppendLine("<html><head>")
    [void]$sb.AppendLine("<meta charset='UTF-8'>")
    [void]$sb.AppendLine("<title>Documentation for $ComputerName</title>")
    
    $cssPath = Join-Path $OutputPath "style.css"
    if (Test-Path $cssPath) {
        [void]$sb.AppendLine("<link rel='stylesheet' type='text/css' href='style.css'>")
    } else {
        # REWORKED: Windows Default Style / Blue Headers / Light Grey BG
        [void]$sb.AppendLine("<style>")
        [void]$sb.AppendLine("body { background-color: #F0F0F0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: black; }")
        [void]$sb.AppendLine("h1, h2, h3 { color: #0078D7; }")
        [void]$sb.AppendLine("h1 { font-size: 16pt; border-bottom: 2px solid #0078D7; padding-bottom: 5px; }")
        [void]$sb.AppendLine("h2 { font-size: 14pt; }")
        [void]$sb.AppendLine("h3 { font-size: 12pt; }")
        [void]$sb.AppendLine("a { color: #0000EE; text-decoration: none; }")
        [void]$sb.AppendLine("a:visited { color: #551A8B; }")
        [void]$sb.AppendLine("a:hover { text-decoration: underline; }")
        [void]$sb.AppendLine("th { color: white; background-color: #0078D7; font-weight: bold; border: 1px solid #666; text-align: left; padding: 4px; }")
        [void]$sb.AppendLine("table { border-collapse: collapse; width: 100%; background-color: white; }")
        [void]$sb.AppendLine("td { font-size: x-small; color: black; border: 1px solid #DDD; padding: 3px; }")
        [void]$sb.AppendLine("b { color: #444; }")
        [void]$sb.AppendLine("</style>")
    }
    [void]$sb.AppendLine("</head><body>")
    
    # REMOVED: Hardcoded logo block (Handled via CSS only)
    
    [void]$sb.AppendLine("<h1>$ComputerName</h1>")
    
    if ($inventory.ComputerSystem) {
        [void]$sb.AppendLine("<b>NetBIOS:</b> $($inventory.ComputerSystem.Name)<br>")
        [void]$sb.AppendLine("<b>FQDN:</b> $($inventory.ComputerSystem.FQDN)<br>")
        [void]$sb.AppendLine("<b>Roles:</b> $($inventory.Roles -join ', ')<br>")
    }
    if ($inventory.OperatingSystem) {
        [void]$sb.AppendLine("<b>OS:</b> $($inventory.OperatingSystem.Caption)<br>")
    }
    if ($inventory.SystemProduct) {
        [void]$sb.AppendLine("<b>Identifying Number:</b> $($inventory.SystemProduct.IdentifyingNumber)<br>")
    }
    [void]$sb.AppendLine("<b>Scan Time:</b> $($inventory.ScanTime)<br><br>")
    
    [void]$sb.AppendLine("<h1><a name='toc'>Table Of Contents</a></h1>")
    [void]$sb.AppendLine("<ol>")
    
    [void]$sb.AppendLine("<li><a href='#hardware'>Hardware Platform</a></li>")
    [void]$sb.AppendLine("<ol>")
    [void]$sb.AppendLine("<li><a href='#hardware_general'>General Information</a></li>")
    [void]$sb.AppendLine("<li><a href='#hardware_bios'>BIOS Information</a></li>")
    [void]$sb.AppendLine("</ol>")
    
    [void]$sb.AppendLine("<li><a href='#software'>Software Platform</a></li>")
    [void]$sb.AppendLine("<ol>")
    [void]$sb.AppendLine("<li><a href='#software_general'>General Information</a></li>")
    [void]$sb.AppendLine("<li><a href='#software_patches'>Installed Patches</a></li>")
    [void]$sb.AppendLine("<li><a href='#software_installed'>Installed Software</a></li>")
    [void]$sb.AppendLine("<li><a href='#software_uwp'>Universal Apps</a></li>")
    [void]$sb.AppendLine("</ol>")
    
    [void]$sb.AppendLine("<li><a href='#storage'>Storage</a></li>")
    [void]$sb.AppendLine("<ol><li><a href='#storage_general'>General Information</a></li></ol>")
    
    [void]$sb.AppendLine("<li><a href='#network'>Network Configuration</a></li>")
    
    [void]$sb.AppendLine("<li><a href='#miscellaneous'>Miscellaneous Configuration</a></li>")
    [void]$sb.AppendLine("<ol>")
    if ($inventory.EventLogs) { [void]$sb.AppendLine("<li><a href='#miscellaneous_eventlog'>Event Log Files</a></li>") }
    if ($inventory.LocalGroups) { [void]$sb.AppendLine("<li><a href='#miscellaneous_localgroups'>Local Groups</a></li>") }
    if ($inventory.LocalUsers) { [void]$sb.AppendLine("<li><a href='#miscellaneous_localusers'>Local Users</a></li>") }
    if ($inventory.UserProfiles) { [void]$sb.AppendLine("<li><a href='#miscellaneous_userprofiles'>User Profiles</a></li>") }
    if ($inventory.Printers) { [void]$sb.AppendLine("<li><a href='#miscellaneous_printers'>Printers</a></li>") }
    [void]$sb.AppendLine("<li><a href='#miscellaneous_regional'>Regional Settings</a></li>")
    if ($inventory.Processes) { [void]$sb.AppendLine("<li><a href='#miscellaneous_processes'>Currently running processes</a></li>") }
    if ($inventory.Services) { [void]$sb.AppendLine("<li><a href='#miscellaneous_services'>Services</a></li>") }
    if ($inventory.Shares) { [void]$sb.AppendLine("<li><a href='#miscellaneous_shares'>Shares</a></li>") }
    if ($inventory.StartupCommands) { [void]$sb.AppendLine("<li><a href='#miscellaneous_startupcommand'>Startup Command</a></li>") }
    if ($inventory.PageFiles) { [void]$sb.AppendLine("<li><a href='#miscellaneous_virtualmemory'>Virtual Memory</a></li>") }
    if ($inventory.Registry) { [void]$sb.AppendLine("<li><a href='#miscellaneous_registry'>Windows Registry</a></li>") }
    [void]$sb.AppendLine("</ol>")
    [void]$sb.AppendLine("</ol>")
    
    # --- HARDWARE SECTION ---
    [void]$sb.AppendLine("<h1 id='hardware'>Hardware Platform</h1>")
    [void]$sb.AppendLine("<h2 id='hardware_general'>General Information</h2>")
    if ($inventory.SystemProduct) {
        [void]$sb.AppendLine("<b>Manufacturer:</b> $($inventory.SystemProduct.Vendor)<br>")
        [void]$sb.AppendLine("<b>Product Name:</b> $($inventory.SystemProduct.Name)<br>")
        [void]$sb.AppendLine("<b>Identifying Number:</b> $($inventory.SystemProduct.IdentifyingNumber)<br>")
    }
    if ($inventory.SystemEnclosure) {
        [void]$sb.AppendLine("<b>Chassis:</b> $($inventory.SystemEnclosure.ChassisType)<br><br>")
    }
    
    if ($inventory.Processor) {
        [void]$sb.AppendLine("<strong>Processor</strong><br>")
        [void]$sb.AppendLine("<b>Name:</b> $($inventory.Processor.Name)<br>")
        [void]$sb.AppendLine("<b>Description:</b> $($inventory.Processor.Description)<br>")
        [void]$sb.AppendLine("<b>Speed:</b> $($inventory.Processor.MaxClockSpeed) MHz<br>")
        [void]$sb.AppendLine("<b>L2 Cache Size:</b> $($inventory.Processor.L2CacheSize) KB<br>")
        [void]$sb.AppendLine("<b>External Clock:</b> $($inventory.Processor.ExtClock) MHz<br>")
        [void]$sb.AppendLine("The system has $($inventory.ProcessorCount) processor(s)<br><br>")
    }
    
    [void]$sb.AppendLine("<strong>Memory</strong><br>")
    if ($inventory.ComputerSystem) {
        [void]$sb.AppendLine("<b>Total Memory:</b> $($inventory.ComputerSystem.TotalPhysicalMemoryMB) MB<br>")
    }
    if ($inventory.OperatingSystem) {
        [void]$sb.AppendLine("<b>Free Memory:</b> $($inventory.OperatingSystem.FreePhysicalMemoryMB) MB<br>")
    }
    
    [void]$sb.AppendLine("<table><tr><th>Bank</th><th>Capacity</th><th>Form</th><th>Type</th><th>Manufacturer</th><th>Part#</th><th>Serial#</th></tr>")
    foreach ($mem in $inventory.PhysicalMemory) {
        [void]$sb.AppendLine("<tr><td>$($mem.BankLabel)</td><td>$($mem.Capacity / 1MB) MB</td><td>$($mem.FormFactor)</td><td>$($mem.MemoryType)</td><td>$($mem.Manufacturer)</td><td>$($mem.PartNumber)</td><td>$($mem.SerialNumber)</td></tr>")
    }
    [void]$sb.AppendLine("</table>")
    
    if ($inventory.VideoController) {
        [void]$sb.AppendLine("<br><strong>Video Controller</strong><br>")
        [void]$sb.AppendLine("<table><tr><th>Name</th><th>Adapter RAM</th><th>Compatibility</th></tr>")
        foreach ($vid in $inventory.VideoController) {
            [void]$sb.AppendLine("<tr><td>$($vid.Name)</td><td>$($vid.AdapterRAM / 1MB) MB</td><td>$($vid.AdapterCompatibility)</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }
    
    if ($inventory.Monitors) {
        [void]$sb.AppendLine("<br><strong>Monitors</strong><br>")
        [void]$sb.AppendLine("<table><tr><th>Model</th><th>Manufacturer</th><th>Connection</th><th>Serial</th></tr>")
        foreach ($mon in $inventory.Monitors) {
            [void]$sb.AppendLine("<tr><td>$($mon.Model)</td><td>$($mon.Manufacturer)</td><td>$($mon.ConnectionType)</td><td>$($mon.SerialNumber)</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }
    
    # BIOS
    [void]$sb.AppendLine("<h2 id='hardware_bios'>Bios Information</h2>")
    if ($inventory.BIOS) {
        [void]$sb.AppendLine("<b>Bios Version:</b> $($inventory.BIOS.Version)<br>")
        [void]$sb.AppendLine("<b>SMBios Version:</b> $($inventory.BIOS.SMBIOSBIOSVersion) (Major: $($inventory.BIOS.SMBIOSMajorVersion), Minor: $($inventory.BIOS.SMBIOSMinorVersion))<br>")
        
        if ($inventory.BIOS.BiosCharacteristics) {
            [void]$sb.AppendLine("<br><strong>BIOS Characteristics</strong>")
            [void]$sb.AppendLine("<table><tr><th>Code</th><th>Characteristic</th></tr>")
            foreach ($code in $inventory.BIOS.BiosCharacteristics) {
                $desc = Get-BiosCharacteristic -Code $code
                [void]$sb.AppendLine("<tr><td>$code</td><td>$desc</td></tr>")
            }
            [void]$sb.AppendLine("</table>")
        }
    }
    
    # --- SOFTWARE SECTION ---
    [void]$sb.AppendLine("<h1 id='software'>Software Platform</h1>")
    [void]$sb.AppendLine("<h2 id='software_general'>General Information</h2>")
    if ($inventory.OperatingSystem) {
        [void]$sb.AppendLine("<b>OS Name:</b> $($inventory.OperatingSystem.Caption)<br>")
        [void]$sb.AppendLine("<b>Windows Location:</b> $($inventory.OperatingSystem.WindowsDirectory)<br>")
        [void]$sb.AppendLine("<b>Install Date:</b> $($inventory.OperatingSystem.InstallDate)<br>")
        [void]$sb.AppendLine("<b>Operating System Language:</b> $($inventory.OperatingSystem.OSLanguage)<br>")
    }
    if ($inventory.ComputerSystem) {
        [void]$sb.AppendLine("<b>OS Configuration:</b> $($inventory.ComputerSystem.OSConfigString)<br>")
    }
    
    # VBS Parity: Added Last User and PS Version
    [void]$sb.AppendLine("<b>Last Logged on User:</b> $($inventory.LastLoggedOnUser)<br>")
    [void]$sb.AppendLine("<b>PowerShell Version:</b> $($inventory.PowerShellVersion)<br>")
    
    if ($adDescription -ne "N/A") { [void]$sb.AppendLine("<b>Description (AD):</b> $adDescription<br>") }
    
    if ($inventory.ServerFeatures) {
        [void]$sb.AppendLine("<br><strong>Server Features</strong><br>")
        [void]$sb.AppendLine("<table><tr><th>Parent ID</th><th>ID</th><th>Name</th></tr>")
        foreach ($feat in $inventory.ServerFeatures) {
            [void]$sb.AppendLine("<tr><td>$($feat.ParentID)</td><td>$($feat.ID)</td><td>$($feat.Name)</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }
    
    # Patches
    [void]$sb.AppendLine("<h2 id='software_patches'>Installed Patches</h2>")
    [void]$sb.AppendLine("<table><tr><th>Patch ID</th><th>Description</th><th>Install Date</th></tr>")
    foreach ($patch in $inventory.Patches) {
        [void]$sb.AppendLine("<tr><td>$($patch.HotFixID)</td><td>$($patch.Description)</td><td>$($patch.InstalledOn)</td></tr>")
    }
    [void]$sb.AppendLine("</table>")
    
    # Installed Software (Unified Table)
    [void]$sb.AppendLine("<h2 id='software_installed'>Installed Software</h2>")
    [void]$sb.AppendLine("<table><tr><th>Name</th><th>Publisher</th><th>Version</th><th>Install Date</th><th>Source</th></tr>")
    # Filter KB updates here
    $displaySoftware = $inventory.Software | Where-Object { $_.Name -notmatch '^KB\d{6}' }
    foreach ($app in $displaySoftware) {
        [void]$sb.AppendLine("<tr><td>$($app.Name)</td><td>$($app.Publisher)</td><td>$($app.Version)</td><td>$($app.InstallDate)</td><td>$($app.Source)</td></tr>")
    }
    [void]$sb.AppendLine("</table>")
    
    # UWP Apps (Combined Section)
    [void]$sb.AppendLine("<h2 id='software_uwp'>Universal Windows Apps</h2>")
    
    if ($inventory.ProvisionedApps) {
        [void]$sb.AppendLine("<h3>Provisioned (Global) Apps</h3>")
        [void]$sb.AppendLine("<table><tr><th>Display Name</th><th>Version</th></tr>")
        foreach ($uwp in $inventory.ProvisionedApps) {
            [void]$sb.AppendLine("<tr><td>$($uwp.DisplayName)</td><td>$($uwp.Version)</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }
    
    if ($inventory.UserApps) {
        [void]$sb.AppendLine("<h3>Installed Apps (Found on System)</h3>")
        [void]$sb.AppendLine("<table><tr><th>Name</th><th>Publisher</th><th>Version</th></tr>")
        foreach ($uwp in $inventory.UserApps) {
            [void]$sb.AppendLine("<tr><td>$($uwp.Name)</td><td>$($uwp.Publisher)</td><td>$($uwp.Version)</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }
    
    # --- STORAGE SECTION ---
    [void]$sb.AppendLine("<h1 id='storage'>Storage</h1>")
    [void]$sb.AppendLine("<h2 id='storage_general'>General Information</h2>")
    
    foreach ($pdisk in $inventory.PhysicalDisks) {
        $sizeGB = [math]::Round($pdisk.Size / 1GB, 2)
        [void]$sb.AppendLine("<strong>$($pdisk.Model) - $($pdisk.DeviceID)</strong><br>")
        [void]$sb.AppendLine("<b>Interface:</b> $($pdisk.Interface)<br>")
        [void]$sb.AppendLine("<b>Media Type:</b> $($pdisk.MediaType)<br>")
        [void]$sb.AppendLine("<b>Total Disk Size:</b> $sizeGB Gb<br>")
        
        foreach ($partition in $pdisk.Partitions) {
            $partSizeGB = [math]::Round($partition.Size / 1GB, 2)
            $partFreeGB = [math]::Round($partition.FreeSpace / 1GB, 2)
            [void]$sb.AppendLine("$($partition.VolumeName) ($($partition.DeviceID)) $partSizeGB Gb ($partFreeGB Gb Free) $($partition.FileSystem)<br>")
        }
        [void]$sb.AppendLine("<br>")
    }
    
    # --- NETWORK SECTION ---
    [void]$sb.AppendLine("<h1 id='network'>Network Configuration</h1>")
    [void]$sb.AppendLine("<table><tr><th>Name</th><th>Type</th><th>Status</th><th>MAC Address</th><th>IP Address</th></tr>")
    foreach ($nic in $inventory.NetworkAdapters) {
        $ipDisplay = if ($nic.IPAddress) { $nic.IPAddress -join ', ' } else { "" }
        [void]$sb.AppendLine("<tr><td>$($nic.Name)<br><i>$($nic.Description)</i></td><td>$($nic.Type)</td><td>$($nic.Status)</td><td>$($nic.MACAddress)</td><td>$ipDisplay</td></tr>")
    }
    [void]$sb.AppendLine("</table>")
    
    # --- MISC SECTION ---
    [void]$sb.AppendLine("<h1 id='miscellaneous'>Miscellaneous Configuration</h1>")
    
    if ($inventory.EventLogs) {
        [void]$sb.AppendLine("<h2 id='miscellaneous_eventlog'>Event Log Files</h2>")
        foreach ($evt in $inventory.EventLogs) {
            [void]$sb.AppendLine("<strong>$($evt.LogfileName)</strong><br>")
            [void]$sb.AppendLine("<b>File:</b> $($evt.LogfileName)<br>")
            [void]$sb.AppendLine("<b>Maximum Size:</b> $($evt.MaxFileSize) Kb<br>")
            [void]$sb.AppendLine("<b>Overwrite Policy:</b> $($evt.OverwritePolicy)<br><br>")
        }
    }
    
    if ($inventory.LocalGroups) {
        [void]$sb.AppendLine("<h2 id='miscellaneous_localgroups'>Local Groups</h2>")
        [void]$sb.AppendLine("<table><tr><th>Name</th><th>Description</th><th>Members</th></tr>")
        foreach ($grp in $inventory.LocalGroups) {
            $memberStr = $grp.Members -join "<br>"
            [void]$sb.AppendLine("<tr><td>$($grp.Name)</td><td>$($grp.Description)</td><td>$memberStr</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }
    
    if ($inventory.LocalUsers) {
        [void]$sb.AppendLine("<h2 id='miscellaneous_localusers'>Local Users</h2>")
        [void]$sb.AppendLine("<table><tr><th>User</th><th>Description</th></tr>")
        foreach ($usr in $inventory.LocalUsers) {
            [void]$sb.AppendLine("<tr><td>$($usr.Name)</td><td>$($usr.Description)</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }
    
    if ($inventory.UserProfiles) {
        [void]$sb.AppendLine("<h2 id='miscellaneous_userprofiles'>User Profiles</h2>")
        [void]$sb.AppendLine("<table><tr><th>User</th><th>SID</th><th>Path</th></tr>")
        foreach ($prof in $inventory.UserProfiles) {
            [void]$sb.AppendLine("<tr><td>$($prof.User)</td><td>$($prof.SID)</td><td>$($prof.Path)</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }
    
    if ($inventory.Printers) {
        [void]$sb.AppendLine("<h2 id='miscellaneous_printers'>Printers</h2>")
        [void]$sb.AppendLine("<table><tr><th>Name</th><th>Driver</th><th>Port</th></tr>")
        foreach ($prt in $inventory.Printers) {
            [void]$sb.AppendLine("<tr><td>$($prt.Name)</td><td>$($prt.DriverName)</td><td>$($prt.PortName)</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }
    
    [void]$sb.AppendLine("<h2 id='miscellaneous_regional'>Regional Settings</h2>")
    [void]$sb.AppendLine("<b>Time Zone:</b> $($inventory.TimeZone)<br>")
    
    if ($inventory.Processes) {
        [void]$sb.AppendLine("<h2 id='miscellaneous_processes'>Currently running processes</h2>")
        [void]$sb.AppendLine("<table><tr><th>Name</th><th>Executable</th></tr>")
        foreach ($proc in $inventory.Processes) {
            [void]$sb.AppendLine("<tr><td>$($proc.Caption)</td><td>$($proc.ExecutablePath)</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }
    
    if ($inventory.Services) {
        [void]$sb.AppendLine("<h2 id='miscellaneous_services'>Services</h2>")
        [void]$sb.AppendLine("<table><tr><th>Name</th><th>Start Mode</th><th>Started</th></tr>")
        foreach ($svc in $inventory.Services) {
            [void]$sb.AppendLine("<tr><td>$($svc.Caption)</td><td>$($svc.StartMode)</td><td>$($svc.Started)</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }
    
    if ($inventory.Shares) {
        [void]$sb.AppendLine("<h2 id='miscellaneous_shares'>Shares</h2>")
        [void]$sb.AppendLine("<table><tr><th>Name</th><th>Path</th><th>Description</th></tr>")
        foreach ($share in $inventory.Shares) {
            [void]$sb.AppendLine("<tr><td>$($share.Name)</td><td>$($share.Path)</td><td>$($share.Description)</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }
    
    if ($inventory.StartupCommands) {
        [void]$sb.AppendLine("<h2 id='miscellaneous_startupcommand'>Startup Command</h2>")
        [void]$sb.AppendLine("<table><tr><th>User</th><th>Name</th><th>Command</th></tr>")
        foreach ($sc in $inventory.StartupCommands) {
            [void]$sb.AppendLine("<tr><td>$($sc.User)</td><td>$($sc.Name)</td><td>$($sc.Command)</td></tr>")
        }
        [void]$sb.AppendLine("</table>")
    }
    
    [void]$sb.AppendLine("<h2 id='miscellaneous_virtualmemory'>Virtual Memory</h2>")
    foreach ($pf in $inventory.PageFiles) {
        [void]$sb.AppendLine("<strong>Pagefile(s)</strong><br>")
        [void]$sb.AppendLine("<b>Drive:</b> $($pf.Name)<br>")
        [void]$sb.AppendLine("<b>Initial Size:</b> $($pf.InitialSize) MB<br>")
        [void]$sb.AppendLine("<b>Maximum Size:</b> $($pf.MaximumSize) MB<br>")
    }
    
    if ($inventory.Registry) {
        [void]$sb.AppendLine("<h2 id='miscellaneous_registry'>Windows Registry</h2>")
        [void]$sb.AppendLine("<b>Current Registry Size:</b> $($inventory.Registry.CurrentSize) MB<br>")
        [void]$sb.AppendLine("<b>Maximum Registry Size:</b> $($inventory.Registry.MaximumSize) MB<br>")
    }
    
    [void]$sb.AppendLine("</body></html>")
    
    $sb.ToString() | Out-File -FilePath $htmlFile -Encoding UTF8 -Force
    Write-Log "HTML report saved: $htmlFile" -Level Success
    
    # Append to NeedsUpdate.txt if there are issues
    if ($inventory.SoftwareIssues.Count -gt 0 -or $inventory.PatchIssues.Count -gt 0) {
        Write-Log "Appending to NeedsUpdate.txt" -Level Info
        
        # Use AD description in update report
        $updateText = "`n$ComputerName ($adDescription):`n"
        
        foreach ($issue in $inventory.SoftwareIssues) {
            $updateText += "        $($issue.Software) - $($issue.InstalledVersion) -> $($issue.ExpectedVersion)`n"
        }
        
        foreach ($issue in $inventory.PatchIssues) {
            $updateText += "        KB$($issue.KBID): $($issue.Description)`n"
        }
        
        $needsUpdatePath = Join-Path $OutputPath "NeedsUpdate.txt"
        # Write master file
        $success = Write-ToFileWithLock -FilePath $needsUpdatePath -Content $updateText -TimeoutSeconds 30
        
        if (-not $success) {
             Write-Log "Failed to write to NeedsUpdate.txt (Locked/Timeout)" -Level Error
        }
    }
    else {
        Write-Log "No software or patch issues found" -Level Success
    }
    
    # Summary
    $duration = (Get-Date) - $scriptStartTime
    Write-Log "=============================" -Level Success
    Write-Log "Inventory completed successfully" -Level Success
    Write-Log "Computer: $ComputerName" -Level Info
    Write-Log "Software Items: $($inventory.Software.Count)" -Level Info
    Write-Log "Patches: $($inventory.Patches.Count)" -Level Info
    Write-Log "Software Issues: $($inventory.SoftwareIssues.Count)" -Level $(if ($inventory.SoftwareIssues.Count -gt 0) { 'Warning' } else { 'Success' })
    Write-Log "Patch Issues: $($inventory.PatchIssues.Count)" -Level $(if ($inventory.PatchIssues.Count -gt 0) { 'Warning' } else { 'Success' })
    Write-Log "Duration: $([math]::Round($duration.TotalSeconds, 2)) seconds" -Level Info
    Write-Log "=============================" -Level Success
    
}
catch {
    Write-Log "Fatal error: $_" -Level Error
    Write-Log $_.ScriptStackTrace -Level Error
    exit 1
}
finally {
    if ($cimSession) {
        Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
    }
    # FIX: Clean up PSSession
    if ($psSession) {
        Remove-PSSession -Session $psSession -ErrorAction SilentlyContinue
    }
}

#endregion