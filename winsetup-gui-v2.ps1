# ============================================================
#   Windows Setup Toolkit - GUI Version w/ Logging + Run All
#   v3 - Robust, no null-argument crashes
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------------------------
# GUI SETUP
# ---------------------------------------------

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Windows Setup Toolkit"
$Form.Size = New-Object System.Drawing.Size(900,560)
$Form.StartPosition = "CenterScreen"

$LogBox = New-Object System.Windows.Forms.TextBox
$LogBox.Multiline = $true
$LogBox.ScrollBars = "Vertical"
$LogBox.ReadOnly = $true
$LogBox.WordWrap = $false
$LogBox.Size = New-Object System.Drawing.Size(600,480)
$LogBox.Location = New-Object System.Drawing.Point(270,20)
$Form.Controls.Add($LogBox)

# ---------------------------------------------
# LOGGING
# ---------------------------------------------

$LogPath = "C:\WindowsSetupToolkit"
$LogFile = Join-Path $LogPath "install.log"

if (!(Test-Path $LogPath)) {
    New-Item -ItemType Directory -Force -Path $LogPath | Out-Null
}

function Write-Log {
    param([string]$msg)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "$timestamp - $msg"
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
    if ($LogBox) { $LogBox.AppendText("$entry`r`n") }
}

# ---------------------------------------------
# DOWNLOAD + INSTALL HELPERS
# ---------------------------------------------

function Install-EXE {
    param(
        [string]$Name,
        [string]$URL,
        [string]$Args
    )

    Write-Log "Downloading $Name..."
    $Temp = Join-Path $env:TEMP "$Name.exe"
    Invoke-WebRequest -Uri $URL -OutFile $Temp -UseBasicParsing

    Write-Log "Installing $Name silently..."
    if ([string]::IsNullOrWhiteSpace($Args)) {
        Start-Process -FilePath $Temp -Wait
        Write-Log "WARNING: $Name installed with NO arguments (Args was empty)."
    }
    else {
        Start-Process -FilePath $Temp -ArgumentList $Args -Wait
    }

    Remove-Item $Temp -Force -ErrorAction SilentlyContinue
    Write-Log "$Name install step finished."
}

function Install-MSI {
    param(
        [string]$Name,
        [string]$URL,
        [string]$Args
    )

    Write-Log "Downloading $Name..."
    $Temp = Join-Path $env:TEMP "$Name.msi"
    Invoke-WebRequest -Uri $URL -OutFile $Temp -UseBasicParsing

    Write-Log "Installing $Name silently..."
    if ([string]::IsNullOrWhiteSpace($Args)) {
        $msiArgs = "/i `"{0}`"" -f $Temp
        Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs -Wait
        Write-Log "WARNING: $Name installed with NO extra MSI arguments."
    }
    else {
        $msiArgs = "/i `"{0}`" {1}" -f $Temp, $Args
        Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs -Wait
    }

    Remove-Item $Temp -Force -ErrorAction SilentlyContinue
    Write-Log "$Name install step finished."
}

# ---------------------------------------------
# APP INSTALL FUNCTIONS
# ---------------------------------------------

function Install-Chrome {
    Install-EXE -Name "Chrome" `
        -URL "https://dl.google.com/chrome/install/latest/chrome_installer.exe" `
        -Args "/silent /install"
}

function Install-Takeoff {
    Install-EXE -Name "Takeoff" `
        -URL "https://downloads.oncenter.com/Downloads/OST/400/OST4.0.0.288Setup.exe" `
        -Args '/s /v"/qn"'
}

function Install-QuickBid {
    Install-EXE -Name "QuickBid" `
        -URL "https://downloads.oncenter.com/Downloads/QB/499/QB4990516Setup.exe" `
        -Args '/s /v"/qn"'
}

function Install-Bluebeam21 {
    Install-EXE -Name "BluebeamRevu21" `
        -URL "https://bluebeam.com/FullRevuTRIAL" `
        -Args '/s /v"/qn /norestart"'
}

function Install-Office365 {
    Install-EXE -Name "ODT" `
        -URL "https://download.microsoft.com/download/2/8/E/28E8EC70-BD0C-4CD1-B447-3B0C10CC9F40/officedeploymenttool.exe" `
        -Args "/quiet"
}

function Install-Teams {
    Install-EXE -Name "TeamsWork" `
        -URL "https://statics.teams.cdn.office.net/production-windows-x86/lkg/MSTeamsSetup.exe" `
        -Args "--silent --disableAutoLaunch"
}

function Run-Debloat {
    Write-Log "Running Windows Debloat..."
    (Invoke-WebRequest -UseBasicParsing -Uri "https://git.io/debloat").Content | Invoke-Expression
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
        Get-AppxProvisionedPackage -Online |
            Where-Object { $_.DisplayName -eq $app } |
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

function Run-CTTWinUtil {
    Write-Log "Launching Chris Titus Windows Utility..."
    (Invoke-WebRequest -UseBasicParsing -Uri "https://christitus.com/win").Content | Invoke-Expression
    Write-Log "CTT Windows Utility execution finished."
}

# ---------------------------------------------
# RUN ALL
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
# BUTTON HELPER
# ---------------------------------------------

function Add-Button {
    param(
        [string]$text,
        [int]$x,
        [int]$y,
        [scriptblock]$action
    )

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.Size = New-Object System.Drawing.Size(240,35)
    $btn.Location = New-Object System.Drawing.Point($x,$y)
    $btn.Add_Click($action)
    $Form.Controls.Add($btn)
}

# ---------------------------------------------
# BUTTONS
# ---------------------------------------------

Add-Button "Install Chrome"            20  20 { Install-Chrome }
Add-Button "Install QuickBid"          20  65 { Install-QuickBid }
Add-Button "Install Takeoff"           20 110 { Install-Takeoff }
Add-Button "Install Bluebeam 21"       20 155 { Install-Bluebeam21 }
Add-Button "Install Office 365"        20 200 { Install-Office365 }
Add-Button "Install Teams"             20 245 { Install-Teams }
Add-Button "Run Debloat"               20 290 { Run-Debloat }
Add-Button "Remove Preloaded Office"   20 335 { Remove-PreloadedOffice }
Add-Button "Run Windows Update"        20 380 { Run-WindowsUpdate }
Add-Button "Run CTT Windows Utility"   20 425 { Run-CTTWinUtil }
Add-Button "RUN ALL"                   20 470 { Run-All }

Write-Log "GUI loaded."
[void]$Form.ShowDialog()
