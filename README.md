# System Inventory PowerShell Scripts

PowerShell replacement for VBScript SYDI-Server inventory system.

## Quick Start

### Run on Single Computer (Local or Remote)
```powershell
# Run on local computer
.\RunInv.ps1

# Run on remote computer
.\RunInv.ps1 -ComputerName SERVER01
```

### Run on All Active Directory Computers
```powershell
# Run from Domain Controller
.\NetworkInventory.ps1
```

## Files

- **NetworkInventory.ps1** - Queries AD and runs inventory on all active computers
- **RunInv.ps1** - Collects inventory from single computer (local or remote)
- **software.xml** - Software version requirements and patch definitions

## Requirements

- PowerShell 5.1+
- Administrator privileges
- Active Directory module (for NetworkInventory.ps1 only)
- WinRM/PowerShell Remoting enabled on target computers

## Enabling PowerShell Remoting via GPO

To enable WinRM/PowerShell Remoting on all computers:

### 1. Create/Edit Group Policy
- Open **Group Policy Management**
- Edit Domain Policy or create new GPO
- Link to appropriate OU containing computer accounts

### 2. Configure WinRM Service
Navigate to:
```
Computer Configuration
└── Policies
    └── Windows Settings
        └── Security Settings
            └── System Services
                └── Windows Remote Management (WS-Management)
```
- Set to **Automatic**
- Click **Start**

### 3. Configure Firewall Rules
Navigate to:
```
Computer Configuration
└── Policies
    └── Windows Settings
        └── Security Settings
            └── Windows Defender Firewall with Advanced Security
                └── Inbound Rules
```

**Create Custom Rule:**
- New Rule > Port
- TCP, Specific local ports: **5985**
- Allow the connection
- Apply to: Domain, Private
- Name: "WinRM for Inventory from DC"
- Edit rule > Scope tab:
  - Remote IP addresses: **10.0.0.1** (Replace with your DC IP)
  - Or for multiple DCs: **10.0.0.1, 10.0.0.2**

### 4. Enable PowerShell Remoting via Startup Script
Navigate to:
```
Computer Configuration
└── Policies
    └── Windows Settings
        └── Scripts (Startup/Shutdown)
            └── Startup
```
Add PowerShell script with content:
```powershell
Enable-PSRemoting -Force -SkipNetworkProfileCheck
```

### 5. Alternative: Use GPO Preferences
Navigate to:
```
Computer Configuration
└── Preferences
    └── Control Panel Settings
        └── Services
```
- New Service: **WinRM**
- Startup: **Automatic**
- Service action: **Start service**

### 6. Apply and Test
```powershell
# Force GPO update on client
gpupdate /force

# Test from DC
Test-WSMan -ComputerName CLIENTPC01
```

## Output

All reports saved to `Reports\` folder:
- **ComputerAccounts.log** - Active/Inactive/Orphaned AD computer accounts
- **NeedsUpdate.txt** - Software and patches needing updates
- **<ComputerName>.html** - Individual computer inventory reports

## Usage Examples

```powershell
# Network inventory with custom settings
.\NetworkInventory.ps1 -InactiveDays 60 -MaxConcurrent 20

# Only generate AD computer report, skip inventory
.\NetworkInventory.ps1 -SkipInventory

# Single computer with all data
.\RunInv.ps1 -ComputerName SERVER01

# Use credentials
$cred = Get-Credential
.\RunInv.ps1 -ComputerName SERVER01 -Credential $cred
```

## Common Parameters

**NetworkInventory.ps1:**
- `-InactiveDays` - Threshold for inactive computers (default: 90)
- `-MaxConcurrent` - Max parallel jobs (default: 10)
- `-Timeout` - Job timeout in seconds (default: 180)
- `-SkipInventory` - Only generate AD report

**RunInv.ps1:**
- `-ComputerName` - Target computer (default: local)
- `-IncludeWin32Product` - Include MSI apps (slower)
- `-IncludeProcesses` - Include running processes
- `-OutputPath` - Custom output location
- `-Credential` - Alternate credentials

## Troubleshooting

**PowerShell Remoting not working:**
```powershell
# Test connectivity
Test-WSMan -ComputerName COMPUTER01

# Enable on target (if GPO not applied)
Invoke-Command -ComputerName COMPUTER01 -ScriptBlock { Enable-PSRemoting -Force }

# Check firewall
Enable-NetFirewallRule -DisplayGroup "Windows Remote Management"
```

**Permission denied:**
- Ensure running as Domain Admin or account with local admin rights on targets
- Check credentials parameter if using alternate account

**Slow execution:**
- Reduce `-MaxConcurrent` parameter
- Skip Win32_Product collection (it's slow)
- Increase `-Timeout` for slower computers

## Documentation

See additional files:
- **INSTALLATION.md** - Detailed setup guide

- **EXAMPLES.md** - Usage examples and workflows
