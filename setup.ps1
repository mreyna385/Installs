#Requires -Version 5.1

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Admin elevation check
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"iwr -useb https://raw.githubusercontent.com/mreyna385/Installs/main/setup.ps1 | iex`"" -Verb RunAs
    exit
}

# Create log directory
if (-not (Test-Path 'C:\WinSetupToolkit')) {
    New-Item -ItemType Directory -Path 'C:\WinSetupToolkit' | Out-Null
}

$global:LogFile = 'C:\WinSetupToolkit\install.log'

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )
    try {
        $ts    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $entry = "[$ts] [$Level] $Message"
        Add-Content -Path $global:LogFile -Value $entry -ErrorAction SilentlyContinue
        switch ($Level) {
            'ERROR'   { Write-Host $entry -ForegroundColor Red }
            'SUCCESS' { Write-Host $entry -ForegroundColor Green }
            'WARN'    { Write-Host $entry -ForegroundColor Yellow }
            default   { Write-Host $entry -ForegroundColor Cyan }
        }
    } catch {
        Write-Host "[LOGGING ERROR] $Message"
    }
}

# Download and dot-source apps.ps1
try {
    Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/mreyna385/Installs/main/apps.ps1' `
        -OutFile "$env:TEMP\WinSetup_apps.ps1" -UseBasicParsing -ErrorAction Stop
    . "$env:TEMP\WinSetup_apps.ps1"
} catch {
    Write-Log "Failed to download apps.ps1: $_" 'ERROR'
    exit 1
}

# Download and dot-source diagnostics.ps1
try {
    Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/mreyna385/Installs/main/diagnostics.ps1' `
        -OutFile "$env:TEMP\WinSetup_diagnostics.ps1" -UseBasicParsing -ErrorAction Stop
    . "$env:TEMP\WinSetup_diagnostics.ps1"
} catch {
    Write-Log "Failed to download diagnostics.ps1: $_" 'ERROR'
    exit 1
}

# Main menu loop
do {
    Clear-Host
    Write-Host '╔══════════════════════════════════════════════╗' -ForegroundColor Cyan
    Write-Host '║      MSP Windows Setup Toolkit v2.0         ║' -ForegroundColor Cyan
    Write-Host '╠══════════════════════════════════════════════╣' -ForegroundColor Cyan
    Write-Host '║  INSTALLS                                   ║' -ForegroundColor Cyan
    Write-Host '║  1) Install Chrome                          ║' -ForegroundColor Cyan
    Write-Host '║  2) Install Firefox                         ║' -ForegroundColor Cyan
    Write-Host '║  3) Install Microsoft 365                   ║' -ForegroundColor Cyan
    Write-Host '║  4) Install Microsoft Teams                 ║' -ForegroundColor Cyan
    Write-Host '║  5) Install Adobe Acrobat Reader            ║' -ForegroundColor Cyan
    Write-Host '║  6) Install Takeoff (On Center)             ║' -ForegroundColor Cyan
    Write-Host '║  7) Install QuickBid (On Center)            ║' -ForegroundColor Cyan
    Write-Host '║  8) Install Bluebeam Revu 21                ║' -ForegroundColor Cyan
    Write-Host '║  9) Remove Preloaded Office/Teams AppX      ║' -ForegroundColor Cyan
    Write-Host '║                                             ║' -ForegroundColor Cyan
    Write-Host '║  UPDATES                                    ║' -ForegroundColor Cyan
    Write-Host '║  W) Windows Update (GUI Selector)           ║' -ForegroundColor Cyan
    Write-Host '║                                             ║' -ForegroundColor Cyan
    Write-Host '║  DIAGNOSTICS & REPORTING                    ║' -ForegroundColor Cyan
    Write-Host '║  D) Full Diagnostic Report                  ║' -ForegroundColor Cyan
    Write-Host '║  E) Scan Event Logs                         ║' -ForegroundColor Cyan
    Write-Host '║  N) Network Diagnostics                     ║' -ForegroundColor Cyan
    Write-Host '║  K) Disk Health (SMART)                     ║' -ForegroundColor Cyan
    Write-Host '║  S) Software Inventory                      ║' -ForegroundColor Cyan
    Write-Host '║  U) Startup Programs Audit                  ║' -ForegroundColor Cyan
    Write-Host '║  V) Driver Update Scan                      ║' -ForegroundColor Cyan
    Write-Host '║                                             ║' -ForegroundColor Cyan
    Write-Host '║  QUICK FIXES                                ║' -ForegroundColor Cyan
    Write-Host '║  F) Flush DNS + Reset Winsock               ║' -ForegroundColor Cyan
    Write-Host '║  P) Reset Print Spooler                     ║' -ForegroundColor Cyan
    Write-Host '║  T) Sync Time (w32tm)                       ║' -ForegroundColor Cyan
    Write-Host '║  G) Force Group Policy Refresh              ║' -ForegroundColor Cyan
    Write-Host '║  R) User Profile Cleanup                    ║' -ForegroundColor Cyan
    Write-Host '║                                             ║' -ForegroundColor Cyan
    Write-Host '║  A) RUN ALL INSTALLS                        ║' -ForegroundColor Cyan
    Write-Host '║  Q) Quit                                    ║' -ForegroundColor Cyan
    Write-Host '╚══════════════════════════════════════════════╝' -ForegroundColor Cyan
    Write-Host ''

    $choice = (Read-Host 'Select an option').ToUpper().Trim()

    switch ($choice) {
        '1' { Install-Chrome }
        '2' { Install-Firefox }
        '3' { Install-M365 }
        '4' { Install-Teams }
        '5' { Install-AdobeReader }
        '6' { Install-Takeoff }
        '7' { Install-QuickBid }
        '8' { Install-Bluebeam }
        '9' { Remove-PreloadedOfficeTeams }
        'W' { Invoke-WindowsUpdate }
        'D' { Invoke-FullDiagnosticReport }
        'E' { Invoke-EventLogScan }
        'N' { Invoke-NetworkDiagnostics }
        'K' { Invoke-DiskHealthCheck }
        'S' { Get-SoftwareInventory }
        'U' { Get-StartupAudit }
        'V' { Invoke-DriverUpdateScan }
        'F' { Invoke-FlushDNS }
        'P' { Invoke-PrintSpoolerReset }
        'T' { Invoke-TimeSyncFix }
        'G' { Invoke-GPUpdate }
        'R' { Invoke-UserProfileCleanup }
        'A' { Invoke-RunAll }
        'Q' { Write-Log 'Exiting MSP Windows Setup Toolkit.' 'INFO' }
        default { Write-Host 'Invalid selection, try again.' -ForegroundColor Yellow }
    }

    if ($choice -ne 'Q') {
        Write-Host ''
        Read-Host 'Press Enter to return to menu'
    }

} while ($choice -ne 'Q')
