# ============================================================
#   Windows Setup Toolkit - GUI Version w/ Logging + Run All
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -------- Logging System --------
$LogPath = "C:\WindowsSetupToolkit"
$LogFile = "$LogPath\install.log"

if (!(Test-Path $LogPath)) { New-Item -ItemType Directory -Force -Path $LogPath }

function Write-Log {
    param([string]$msg)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "$timestamp - $msg"
    Add-Content $LogFile $entry
    $LogBox.AppendText("$entry`r`n")
}

# ---------------------------------------------
# DOWNLOAD + SILENT INSTALL FUNCTIONS
# ---------------------------------------------

function Install-EXE {
    param(
        [string]$Name,
        [string]$URL,
        [string]$Args
    )

    Write-Log "Downloading $Name..."
    $Temp = "$env:TEMP\$Name.exe"
    Invoke-WebRequest $URL -OutFile $Temp -UseBasicParsing

    Write-Log "Installing $Name silently..."
    Start-Process $Temp -ArgumentList $Args -Wait
    Remove-Item $Temp -Force
    Write-Log "$Name installed successfully."
}

function Install-MSI {
    param(
        [string]$Name,
        [string]$URL,
        [string]$Args
    )

    Write-Log "Downloading $Name..."
    $Temp = "$env:TEMP\$Name.msi"
    Invoke-WebRequest $URL -OutFile $Temp -UseBasicParsing

    Write-Log "Installing $Name silently..."
    Start-Process "msiexec.exe" -ArgumentList "/i `"$Temp`" $Args" -Wait
    Remove-Item $Temp -Force

    Write-Log "$Name installed successfully."
}

# ---------------------------------------------
# APP INSTALLS (with correct silent arguments)
# ---------------------------------------------

function Install-Chrome {
    Install-EXE "Chrome" `
        "https://dl.google.com/chrome/install/latest/chrome_installer.exe" `
        "/silent /install"
}

function Install-Takeoff {
    Install-EXE "Takeoff" `
        "https://downloads.oncenter.com/Downloads/OST/400/OST4.0.0.288Setup.exe" `
        "/s /v`"/qn`""
}

function Install-QuickBid {
    Install-EXE "QuickBid" `
        "https://downloads.oncenter.com/Downloads/QB/499/QB4990516Setup.exe" `
        "/s /v`"/qn`""
}

function Install-Bluebeam21 {
    Install-EXE "BluebeamRevu21" `
        "https://bluebeam.com/FullRevuTRIAL" `
        "/s /v`"/qn /norestart`""
}

function Install-Office365 {
    Install-EXE "ODT" `
        "https://download.microsoft.com/download/2/8/E/28E8EC70-BD0C-4CD1-B447-3B0C10CC9F40/officedeploymenttool.exe" `
        "/quiet"
}

function Install-Teams {
    Install-EXE "TeamsWork" `
        "https://statics.teams.cdn.office.net/production-windows-x86/lkg/MSTeamsSetup.exe" `
        "--silent --disableAutoLaunch"
}

function Run-Debloat {
    Write-Log "Running Windows Debloat..."
    iwr -useb https://git.io/debloat | iex
    Write-Log "Debloat complete."
}

function Remove-PreloadedOffice {
    Write-Log "Removing Microsoft preloaded Office apps..."
    
    $Apps = @(
        "Microsoft.Office.Desktop",
        "Microsoft.OfficeHub",
        "Microsoft.MicrosoftOfficeHub",
        "Microsoft.Office.OneNote",
        "MicrosoftTeams"
    )

    foreach ($app in $Apps) {
        Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq $app} |
            Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
    }

    Write-Log "Preloaded Office apps removed."
}

function Run-WindowsUpdate {
    Write-Log "Running Windows Update..."
    Install-Module PSWindowsUpdate -Force
    Import-Module PSWindowsUpdate
    Get-WindowsUpdate -AcceptAll -Install -AutoReboot
    Write-Log "Windows Update completed."
}

# ---------------------------------------------
# RUN ALL IN ORDER
# ---------------------------------------------
function Run-All {
    Write-Log "=== STARTING FULL AUTOMATED INSTALL ==="

    Install-Chrome
    Install-Takeoff
    Install-QuickBid
    Install-Bluebeam21
    Install-Office365
    Install-Teams
    Run-Debloat
    Remove-PreloadedOffice
    Run-WindowsUpdate

    Write-Log "=== FULL INSTALL COMPLETED ==="
}

# ---------------------------------------------
# GUI SETUP
# ---------------------------------------------

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Windows Setup Toolkit"
$Form.Size = New-Object System.Drawing.Size(820,520)
$Form.StartPosition = "CenterScreen"

function Add-Button {
    param($text, $x, $y, $action)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.Size = New-Object System.Drawing.Size(240,35)
    $btn.Location = New-Object System.Drawing.Point($x,$y)
    $btn.Add_Click($action)
    $Form.Controls.Add($btn)
}

Add-Button "Install Chrome" 20 20 { Install-Chrome }
Add-Button "Install QuickBid" 20 65 { Install-QuickBid }
Add-Button "Install Takeoff" 20 110 { Install-Takeoff }
Add-Button "Insta
