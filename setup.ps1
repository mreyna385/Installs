#Requires -Version 5.1
<#
.SYNOPSIS
    Windows Setup Toolkit — entry point.
    Run via: iwr -useb https://raw.githubusercontent.com/USERNAME/REPO/main/setup.ps1 | iex
#>

# ── TLS 1.2 ────────────────────────────────────────────────────────────────────
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ── Admin elevation check ───────────────────────────────────────────────────────
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)) {

    Write-Host "Not running as Administrator — relaunching elevated..." -ForegroundColor Yellow
    $psArgs = "-NoProfile -ExecutionPolicy Bypass -Command `"iwr -useb https://raw.githubusercontent.com/USERNAME/REPO/main/setup.ps1 | iex`""
    Start-Process powershell.exe -ArgumentList $psArgs -Verb RunAs
    exit
}

# ── Log directory ───────────────────────────────────────────────────────────────
$global:LogDir  = 'C:\WinSetupToolkit'
$global:LogFile = "$global:LogDir\install.log"
if (-not (Test-Path $global:LogDir)) { New-Item -ItemType Directory -Path $global:LogDir -Force | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    Add-Content -Path $global:LogFile -Value $line
    switch ($Level) {
        'ERROR'   { Write-Host $line -ForegroundColor Red    }
        'SUCCESS' { Write-Host $line -ForegroundColor Green  }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        default   { Write-Host $line }
    }
}

Write-Log "=== Windows Setup Toolkit started ==="

# ── Download and dot-source apps.ps1 ───────────────────────────────────────────
$appsUrl  = 'https://raw.githubusercontent.com/USERNAME/REPO/main/apps.ps1'
$appsTmp  = "$env:TEMP\WinSetup_apps.ps1"

Write-Log "Downloading apps.ps1 from $appsUrl"
try {
    Invoke-WebRequest -Uri $appsUrl -UseBasicParsing -OutFile $appsTmp -ErrorAction Stop
    . $appsTmp
    Write-Log "apps.ps1 loaded successfully"
} catch {
    Write-Log "FATAL: Could not download apps.ps1 — $_" 'ERROR'
    Read-Host "Press Enter to exit"
    exit 1
}

# ── Menu ────────────────────────────────────────────────────────────────────────
function Show-Menu {
    Clear-Host
    Write-Host ""
    Write-Host "=== Windows Setup Toolkit ===" -ForegroundColor Cyan
    Write-Host "  Log: $global:LogFile" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  1) Install Chrome"
    Write-Host "  2) Install Teams"
    Write-Host "  3) Install Microsoft 365"
    Write-Host "  4) Install Takeoff (On Center)"
    Write-Host "  5) Install QuickBid (On Center)"
    Write-Host "  6) Install Bluebeam Revu 21"
    Write-Host "  7) Remove Preloaded Office/Teams AppX"
    Write-Host "  8) Run Windows Update"
    Write-Host "  9) Run CTT WinUtil"
    Write-Host "  A) RUN ALL"
    Write-Host "  Q) Quit"
    Write-Host ""
}

# ── Main loop ───────────────────────────────────────────────────────────────────
do {
    Show-Menu
    $choice = (Read-Host "Select").Trim().ToUpper()

    switch ($choice) {
        '1' { Install-Chrome }
        '2' { Install-Teams  }
        '3' { Install-M365   }
        '4' { Install-Takeoff }
        '5' { Install-QuickBid }
        '6' { Install-Bluebeam }
        '7' { Remove-PreloadedOfficeTeams }
        '8' { Invoke-WindowsUpdate }
        '9' { Invoke-CttWinUtil }
        'A' { Invoke-RunAll }
        'Q' { Write-Log "User exited toolkit."; break }
        default { Write-Host "  Invalid selection — try again." -ForegroundColor Yellow }
    }

    if ($choice -ne 'Q') {
        Write-Host ""
        Read-Host "Press Enter to return to menu"
    }

} while ($choice -ne 'Q')

Write-Log "=== Windows Setup Toolkit finished ==="
