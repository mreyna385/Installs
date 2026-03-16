#Requires -Version 5.1
<#
.SYNOPSIS
    App definitions and install logic for Windows Setup Toolkit.
    Dot-sourced by setup.ps1 — do not run directly.
#>

# ── Helper: Test winget availability ───────────────────────────────────────────
function Test-WingetAvailable {
    return [bool](Get-Command winget -ErrorAction SilentlyContinue)
}

# ── Helper: Write-Log (stub in case this file is ever tested standalone) ────────
if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    function Write-Log {
        param([string]$Message, [string]$Level = 'INFO')
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Write-Host "[$ts] [$Level] $Message"
    }
}

# ── Helper: Install via winget ──────────────────────────────────────────────────
function Install-ViaWinget {
    param(
        [string]$AppName,
        [string]$WingetId
    )
    Write-Log "Installing $AppName via winget (ID: $WingetId)..."
    try {
        $result = winget install --id $WingetId --silent --accept-package-agreements --accept-source-agreements 2>&1
        if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
            # -1978335189 = APPINSTALLER_ERROR_ALREADY_INSTALLED
            Write-Log "$AppName installed (or already present)." 'SUCCESS'
            return $true
        } else {
            Write-Log "$AppName winget exit code: $LASTEXITCODE. Output: $result" 'WARN'
            return $false
        }
    } catch {
        Write-Log "winget threw an exception for $AppName`: $_" 'ERROR'
        return $false
    }
}

# ── Helper: Install EXE via direct download ─────────────────────────────────────
function Install-EXE {
    param(
        [string]$AppName,
        [string]$Url,
        [string]$SilentArgs = '/S'
    )
    $tmpFile = "$env:TEMP\WinSetup_$($AppName -replace '\s','_').exe"
    Write-Log "Downloading $AppName from $Url..."
    try {
        Invoke-WebRequest -Uri $Url -UseBasicParsing -OutFile $tmpFile -ErrorAction Stop
        Write-Log "Running installer for $AppName with args: $SilentArgs"
        $argList = $SilentArgs -split ' '
        $proc = Start-Process -FilePath $tmpFile -ArgumentList $argList -Wait -PassThru -ErrorAction Stop
        if ($proc.ExitCode -eq 0) {
            Write-Log "$AppName installed successfully." 'SUCCESS'
        } else {
            Write-Log "$AppName installer exited with code $($proc.ExitCode)." 'WARN'
        }
        return ($proc.ExitCode -eq 0)
    } catch {
        Write-Log "Failed to install $AppName`: $_" 'ERROR'
        return $false
    } finally {
        if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    }
}

# ── Helper: Install MSI via direct download ─────────────────────────────────────
function Install-MSI {
    param(
        [string]$AppName,
        [string]$Url,
        [string]$SilentArgs = '/qn'
    )
    $tmpFile = "$env:TEMP\WinSetup_$($AppName -replace '\s','_').msi"
    Write-Log "Downloading $AppName (MSI) from $Url..."
    try {
        Invoke-WebRequest -Uri $Url -UseBasicParsing -OutFile $tmpFile -ErrorAction Stop
        Write-Log "Running msiexec for $AppName with args: /i `"$tmpFile`" $SilentArgs"
        $proc = Start-Process msiexec.exe -ArgumentList "/i `"$tmpFile`" $SilentArgs" -Wait -PassThru -ErrorAction Stop
        if ($proc.ExitCode -eq 0) {
            Write-Log "$AppName installed successfully." 'SUCCESS'
        } else {
            Write-Log "$AppName msiexec exited with code $($proc.ExitCode)." 'WARN'
        }
        return ($proc.ExitCode -eq 0)
    } catch {
        Write-Log "Failed to install $AppName`: $_" 'ERROR'
        return $false
    } finally {
        if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  APP INSTALLS
# ═══════════════════════════════════════════════════════════════════════════════

function Install-Chrome {
    Write-Log "--- Chrome ---"
    try {
        if (Test-WingetAvailable) {
            Install-ViaWinget -AppName 'Google Chrome' -WingetId 'Google.Chrome'
        } else {
            Write-Log "winget unavailable — falling back to direct download for Chrome." 'WARN'
            Install-EXE -AppName 'Google Chrome' `
                        -Url 'https://dl.google.com/chrome/install/latest/chrome_installer.exe' `
                        -SilentArgs '/silent /install'
        }
    } catch {
        Write-Log "Chrome install failed: $_" 'ERROR'
    }
}

function Install-Teams {
    Write-Log "--- Microsoft Teams ---"
    try {
        if (Test-WingetAvailable) {
            Install-ViaWinget -AppName 'Microsoft Teams' -WingetId 'Microsoft.Teams'
        } else {
            Write-Log "winget unavailable — falling back to direct download for Teams." 'WARN'
            Install-EXE -AppName 'Microsoft Teams' `
                        -Url 'https://statics.teams.cdn.office.net/production-windows-x64/enterprise/webview2/lkg/MSTeamsSetup.exe' `
                        -SilentArgs '-s'
        }
    } catch {
        Write-Log "Teams install failed: $_" 'ERROR'
    }
}

function Install-M365 {
    Write-Log "--- Microsoft 365 Apps ---"
    try {
        if (Test-WingetAvailable) {
            Install-ViaWinget -AppName 'Microsoft 365' -WingetId 'Microsoft.Office'
        } else {
            Write-Log "winget unavailable — Microsoft 365 requires winget or the Office Deployment Tool. Skipping direct download." 'WARN'
            Write-Log "Download ODT manually from: https://www.microsoft.com/en-us/download/details.aspx?id=49117" 'WARN'
        }
    } catch {
        Write-Log "Microsoft 365 install failed: $_" 'ERROR'
    }
}

function Install-Takeoff {
    Write-Log "--- Takeoff (On Center) ---"
    try {
        # Not in winget catalog — always use direct download
        Install-EXE -AppName 'Takeoff' `
                    -Url 'https://downloads.oncenter.com/Downloads/OST/400/OST4.0.0.288Setup.exe' `
                    -SilentArgs '/s /v"/qn"'
    } catch {
        Write-Log "Takeoff install failed: $_" 'ERROR'
    }
}

function Install-QuickBid {
    Write-Log "--- QuickBid (On Center) ---"
    try {
        # Not in winget catalog — always use direct download
        Install-EXE -AppName 'QuickBid' `
                    -Url 'https://downloads.oncenter.com/Downloads/QB/499/QB4990516Setup.exe' `
                    -SilentArgs '/s /v"/qn"'
    } catch {
        Write-Log "QuickBid install failed: $_" 'ERROR'
    }
}

function Install-Bluebeam {
    Write-Log "--- Bluebeam Revu 21 ---"
    try {
        Write-Host ""
        Write-Host "  Bluebeam Revu 21 requires a licensed installer URL." -ForegroundColor Yellow
        Write-Host "  Paste your direct download URL (or press Enter to skip):" -ForegroundColor Yellow
        $url = (Read-Host "  Bluebeam installer URL").Trim()

        if ([string]::IsNullOrWhiteSpace($url)) {
            Write-Log "Bluebeam skipped — no URL provided." 'WARN'
            return
        }

        # Determine installer type by extension
        if ($url -match '\.msi(\?|$)') {
            Install-MSI -AppName 'Bluebeam Revu 21' -Url $url -SilentArgs '/qn REBOOT=ReallySuppress'
        } else {
            Install-EXE -AppName 'Bluebeam Revu 21' -Url $url -SilentArgs '/quiet /norestart'
        }
    } catch {
        Write-Log "Bluebeam install failed: $_" 'ERROR'
    }
}

# ── System task: Remove preloaded Office/Teams AppX ────────────────────────────
function Remove-PreloadedOfficeTeams {
    Write-Log "--- Removing preloaded Office/Teams AppX packages ---"
    $packages = @(
        'Microsoft.OfficeHub',
        'Microsoft.MicrosoftOfficeHub',
        'Microsoft.Office.OneNote',
        'MicrosoftTeams'
    )
    foreach ($pkg in $packages) {
        try {
            $found = Get-AppxPackage -Name $pkg -AllUsers -ErrorAction SilentlyContinue
            if ($found) {
                $found | Remove-AppxPackage -AllUsers -ErrorAction Stop
                Write-Log "Removed AppX: $pkg" 'SUCCESS'
            } else {
                Write-Log "AppX not found (already absent): $pkg" 'INFO'
            }
        } catch {
            Write-Log "Failed to remove AppX $pkg`: $_" 'ERROR'
        }
    }

    # Also remove provisioned packages so they don't reinstall for new users
    foreach ($pkg in $packages) {
        try {
            $prov = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName -like "*$pkg*" }
            if ($prov) {
                $prov | Remove-AppxProvisionedPackage -Online -ErrorAction Stop | Out-Null
                Write-Log "Removed provisioned AppX: $pkg" 'SUCCESS'
            }
        } catch {
            Write-Log "Failed to remove provisioned AppX $pkg`: $_" 'WARN'
        }
    }
}

# ── System task: Windows Update ─────────────────────────────────────────────────
function Invoke-WindowsUpdate {
    Write-Log "--- Windows Update ---"
    try {
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-Log "PSWindowsUpdate module not found — installing from PSGallery..."
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop | Out-Null
            Install-Module -Name PSWindowsUpdate -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
            Write-Log "PSWindowsUpdate installed." 'SUCCESS'
        }
        Import-Module PSWindowsUpdate -ErrorAction Stop
        Write-Log "Checking for Windows Updates (this may take a while)..."
        Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot -Verbose
        Write-Log "Windows Update completed." 'SUCCESS'
    } catch {
        Write-Log "Windows Update failed: $_" 'ERROR'
    }
}

# ── System task: CTT WinUtil ────────────────────────────────────────────────────
function Invoke-CttWinUtil {
    Write-Log "--- Chris Titus Tech WinUtil ---"
    try {
        Write-Host ""
        Write-Host "  This will run: iwr christitus.com/win | iex" -ForegroundColor Yellow
        Write-Host "  Review the script at https://github.com/ChrisTitusTech/winutil before proceeding." -ForegroundColor Yellow
        $confirm = (Read-Host "  Proceed? (Y/N)").Trim().ToUpper()
        if ($confirm -ne 'Y') {
            Write-Log "CTT WinUtil skipped by user."
            return
        }
        Write-Log "Launching CTT WinUtil..."
        Invoke-Expression (Invoke-WebRequest -UseBasicParsing 'https://christitus.com/win')
        Write-Log "CTT WinUtil session ended." 'SUCCESS'
    } catch {
        Write-Log "CTT WinUtil failed: $_" 'ERROR'
    }
}

# ── RUN ALL ─────────────────────────────────────────────────────────────────────
function Invoke-RunAll {
    Write-Log "=== RUN ALL started ==="

    $results = [ordered]@{}

    $tasks = [ordered]@{
        'Chrome'                    = { Install-Chrome }
        'Teams'                     = { Install-Teams }
        'Microsoft 365'             = { Install-M365 }
        'Takeoff'                   = { Install-Takeoff }
        'QuickBid'                  = { Install-QuickBid }
        'Bluebeam Revu 21'          = { Install-Bluebeam }
        'Remove Preloaded AppX'     = { Remove-PreloadedOfficeTeams }
        'Windows Update'            = { Invoke-WindowsUpdate }
    }

    foreach ($name in $tasks.Keys) {
        Write-Log ">>> $name"
        try {
            & $tasks[$name]
            $results[$name] = 'OK'
        } catch {
            Write-Log "Unexpected error during $name`: $_" 'ERROR'
            $results[$name] = 'FAILED'
        }
    }

    # ── Summary ────────────────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "══════════════════════════════" -ForegroundColor Cyan
    Write-Host "  RUN ALL — Summary" -ForegroundColor Cyan
    Write-Host "══════════════════════════════" -ForegroundColor Cyan
    foreach ($name in $results.Keys) {
        $status = $results[$name]
        $color  = if ($status -eq 'OK') { 'Green' } else { 'Red' }
        Write-Host ("  {0,-30} {1}" -f $name, $status) -ForegroundColor $color
    }
    Write-Host "══════════════════════════════" -ForegroundColor Cyan
    Write-Log "=== RUN ALL finished ==="
}
