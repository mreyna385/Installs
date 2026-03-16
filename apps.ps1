# ============================================================
#   apps.ps1 - App install library for Windows Setup Toolkit
# ============================================================

$LogPath = "C:\WinSetupToolkit"
$LogFile = Join-Path $LogPath "install.log"

if (!(Test-Path $LogPath)) {
    New-Item -ItemType Directory -Force -Path $LogPath | Out-Null
}

function Write-Log {
    param([string]$msg, [string]$Level = "INFO")
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "[$timestamp] [$Level] $msg"
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
    $color = switch ($Level) {
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        default { "Cyan" }
    }
    Write-Host $entry -ForegroundColor $color
}

function Test-WingetAvailable {
    return ($null -ne (Get-Command winget -ErrorAction SilentlyContinue))
}

function Install-ViaWinget {
    param([string]$Name, [string]$WingetId)
    Write-Log "Installing $Name via winget..."
    try {
        winget install --id $WingetId --silent --accept-package-agreements --accept-source-agreements
        Write-Log "$Name installed successfully."
        return $true
    } catch {
        Write-Log "Winget failed for $Name`: $_" "WARN"
        return $false
    }
}

function Install-EXE {
    param([string]$Name, [string]$URL, [string]$Args)
    Write-Log "Downloading $Name from $URL..."
    $Temp = Join-Path $env:TEMP "$Name.exe"
    try {
        Invoke-WebRequest -Uri $URL -OutFile $Temp -UseBasicParsing
        Write-Log "Installing $Name..."
        if ([string]::IsNullOrWhiteSpace($Args)) {
            Start-Process -FilePath $Temp -Wait
        } else {
            Start-Process -FilePath $Temp -ArgumentList $Args -Wait
        }
        Write-Log "$Name install complete."
    } catch {
        Write-Log "Failed to install $Name`: $_" "ERROR"
    } finally {
        Remove-Item $Temp -Force -ErrorAction SilentlyContinue
    }
}

function Install-MSI {
    param([string]$Name, [string]$URL, [string]$ExtraArgs = "")
    Write-Log "Downloading $Name MSI..."
    $Temp = Join-Path $env:TEMP "$Name.msi"
    try {
        Invoke-WebRequest -Uri $URL -OutFile $Temp -UseBasicParsing
        Write-Log "Installing $Name..."
        $msiArgs = "/i `"$Temp`" /qn /norestart $ExtraArgs".Trim()
        Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs -Wait
        Write-Log "$Name install complete."
    } catch {
        Write-Log "Failed to install $Name`: $_" "ERROR"
    } finally {
        Remove-Item $Temp -Force -ErrorAction SilentlyContinue
    }
}

# ============================================================
#   Individual install functions
# ============================================================

function Install-Chrome {
    Write-Log "=== Chrome ==="
    if (Test-WingetAvailable) {
        Install-ViaWinget -Name "Google Chrome" -WingetId "Google.Chrome"
    } else {
        Install-EXE -Name "Chrome" `
            -URL "https://dl.google.com/chrome/install/latest/chrome_installer.exe" `
            -Args "/silent /install"
    }
}

function Install-Teams {
    Write-Log "=== Microsoft Teams ==="
    if (Test-WingetAvailable) {
        Install-ViaWinget -Name "Microsoft Teams" -WingetId "Microsoft.Teams"
    } else {
        Install-EXE -Name "Teams" `
            -URL "https://statics.teams.cdn.office.net/production-windows-x86/lkg/MSTeamsSetup.exe" `
            -Args "--silent --disableAutoLaunch"
    }
}

function Install-Office365 {
    Write-Log "=== Microsoft 365 ==="
    if (Test-WingetAvailable) {
        Install-ViaWinget -Name "Microsoft 365" -WingetId "Microsoft.Office"
    } else {
        Write-Log "Winget not available. Microsoft 365 requires the Office Deployment Tool." "WARN"
        Write-Log "Download from: https://www.microsoft.com/en-us/download/details.aspx?id=49117" "WARN"
    }
}

function Install-Takeoff {
    Write-Log "=== On Center Takeoff ==="
    Install-EXE -Name "Takeoff" `
        -URL "https://downloads.oncenter.com/Downloads/OST/400/OST4.0.0.288Setup.exe" `
        -Args '/s /v"/qn"'
}

function Install-QuickBid {
    Write-Log "=== On Center QuickBid ==="
    Install-EXE -Name "QuickBid" `
        -URL "https://downloads.oncenter.com/Downloads/QB/499/QB4990516Setup.exe" `
        -Args '/s /v"/qn"'
}

function Install-Bluebeam {
    Write-Log "=== Bluebeam Revu ==="
    Write-Host ""
    Write-Host "Bluebeam requires a direct installer URL (their download page is login-gated)." -ForegroundColor Yellow
    Write-Host "Enter your Bluebeam installer URL, or press Enter to skip: " -ForegroundColor Yellow -NoNewline
    $url = Read-Host
    if ([string]::IsNullOrWhiteSpace($url)) {
        Write-Log "Bluebeam skipped by user." "WARN"
    } else {
        Install-EXE -Name "BluebeamRevu" -URL $url -Args '/s /v"/qn /norestart"'
    }
}

function Remove-PreloadedOffice {
    Write-Log "=== Removing preloaded Office/Teams AppX packages ==="
    $apps = @(
        "Microsoft.OfficeHub",
        "Microsoft.MicrosoftOfficeHub",
        "Microsoft.Office.OneNote",
        "MicrosoftTeams"
    )
    foreach ($app in $apps) {
        try {
            Get-AppxPackage -Name $app -AllUsers -ErrorAction SilentlyContinue |
                Remove-AppxPackage -ErrorAction SilentlyContinue
            Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -eq $app } |
                Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
            Write-Log "Removed: $app"
        } catch {
            Write-Log "Could not remove $app`: $_" "WARN"
        }
    }
    Write-Log "Preloaded Office/Teams removal complete."
}

function Invoke-WindowsUpdate {
    Write-Log "=== Windows Update ==="
    try {
        if (!(Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-Log "Installing PSWindowsUpdate module..."
            Install-Module PSWindowsUpdate -Force -Scope CurrentUser
        }
        Import-Module PSWindowsUpdate
        Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot
        Write-Log "Windows Update complete."
    } catch {
        Write-Log "Windows Update failed: $_" "ERROR"
    }
}

function Invoke-CTTWinUtil {
    Write-Log "=== Chris Titus Windows Utility ==="
    try {
        (Invoke-WebRequest -UseBasicParsing -Uri "https://christitus.com/win").Content | Invoke-Expression
    } catch {
        Write-Log "CTT WinUtil failed: $_" "ERROR"
    }
}

# ============================================================
#   Run All
# ============================================================

function Invoke-RunAll {
    $results = @{}
    Write-Log "=== RUN ALL STARTED ==="

    $tasks = @(
        @{ Name = "Chrome";                Fn = { Install-Chrome } },
        @{ Name = "Teams";                 Fn = { Install-Teams } },
        @{ Name = "Office 365";            Fn = { Install-Office365 } },
        @{ Name = "Takeoff";               Fn = { Install-Takeoff } },
        @{ Name = "QuickBid";              Fn = { Install-QuickBid } },
        @{ Name = "Bluebeam";              Fn = { Install-Bluebeam } },
        @{ Name = "Remove Preloaded Apps"; Fn = { Remove-PreloadedOffice } },
        @{ Name = "Windows Update";        Fn = { Invoke-WindowsUpdate } }
    )

    foreach ($task in $tasks) {
        try {
            & $task.Fn
            $results[$task.Name] = "OK"
        } catch {
            Write-Log "$($task.Name) failed: $_" "ERROR"
            $results[$task.Name] = "FAILED"
        }
    }

    Write-Log "=== RUN ALL FINISHED ==="
    Write-Host ""
    Write-Host "===== SUMMARY =====" -ForegroundColor White
    foreach ($key in $results.Keys) {
        $color = if ($results[$key] -eq "OK") { "Green" } else { "Red" }
        Write-Host "  $key`: $($results[$key])" -ForegroundColor $color
    }
    Write-Host "===================" -ForegroundColor White
}
