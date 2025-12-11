# ============================================================
#   Windows Setup Toolkit - Checkbox GUI + Silent Install
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# ---------------------------------------------
# LOGGING
# ---------------------------------------------

$global:LogBox = $null

$LogPath = "C:\WindowsSetupToolkit"
$LogFile = Join-Path $LogPath "install.log"

if (!(Test-Path $LogPath)) {
    New-Item -ItemType Directory -Force -Path $LogPath | Out-Null
}

function Write-Log {
    param([string]$Message)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry     = "$timestamp - $Message"

    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue

    if ($global:LogBox -ne $null) {
        $global:LogBox.AppendText("$entry`r`n")
    }

    Write-Host $entry
}

# ---------------------------------------------
# DOWNLOAD + INSTALL HELPERS
# ---------------------------------------------

function Install-EXE {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$URL,
        [Parameter()][string]$Args
    )

    Write-Log "Downloading $Name from $URL ..."
    $Temp = Join-Path $env:TEMP "$Name.exe"
    Invoke-WebRequest -Uri $URL -OutFile $Temp -UseBasicParsing

    Write-Log "Installing $Name silently..."
    if ([string]::IsNullOrWhiteSpace($Args)) {
        Start-Process -FilePath $Temp -Wait
        Write-Log "WARNING: $Name installed without explicit silent arguments."
    } else {
        Start-Process -FilePath $Temp -ArgumentList $Args -Wait
    }

    Remove-Item $Temp -Force -ErrorAction SilentlyContinue
    Write-Log "$Name installation step finished."
}

function Install-MSI {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$URL,
        [Parameter()][string]$Args
    )

    Write-Log "Downloading $Name from $URL ..."
    $Temp = Join-Path $env:TEMP "$Name.msi"
    Invoke-WebRequest -Uri $URL -OutFile $Temp -UseBasicParsing

    Write-Log "Installing $Name silently via msiexec..."
    if ([string]::IsNullOrWhiteSpace($Args)) {
        $msiArgs = "/i `"$Temp`" /qn /norestart"
    } else {
        $msiArgs = "/i `"$Temp`" $Args"
    }

    Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs -Wait

    Remove-Item $Temp -Force -ErrorAction SilentlyContinue
    Write-Log "$Name installation step finished."
}

# ---------------------------------------------
# APP INSTALL FUNCTIONS (SILENT)
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

# ---------------------------------------------
# EXTRA TASKS
# ---------------------------------------------

function Run-Debloat {
    Write-Log "Running Windows Debloat script (Chris Titus)..."
    (Invoke-WebRequest -UseBasicParsing -Uri "https://git.io/debloat").Content | Invoke-Expression
    Write-Log "Debloat script finished."
}

function Remove-PreloadedOffice {
    Write-Log "Removing preloaded Microsoft Office UWP apps..."

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

    Write-Log "Preloaded Office apps removed (where present)."
}

function Run-WindowsUpdate {
    Write-Log "Running Windows Update via PSWindowsUpdate..."
    Install-Module PSWindowsUpdate -Force
    Import-Module PSWindowsUpdate
    Get-WindowsUpdate -AcceptAll -Install -AutoReboot
    Write-Log "Windows Update command issued (system may reboot automatically)."
}

# ---------------------------------------------
# GUI SETUP
# ---------------------------------------------

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Windows Setup Toolkit"
$Form.Size = New-Object System.Drawing.Size(900,560)
$Form.StartPosition = "CenterScreen"

# Log box on the right
$global:LogBox = New-Object System.Windows.Forms.TextBox
$global:LogBox.Multiline  = $true
$global:LogBox.ScrollBars = "Vertical"
$global:LogBox.ReadOnly   = $true
$global:LogBox.WordWrap   = $false
$global:LogBox.Size       = New-Object System.Drawing.Size(580,480)
$global:LogBox.Location   = New-Object System.Drawing.Point(290,20)
$Form.Controls.Add($global:LogBox)

# Helper to create checkboxes
function New-CheckBox {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y
    )

    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $Text
    $cb.AutoSize = $true
    $cb.Location = New-Object System.Drawing.Point($X,$Y)
    $Form.Controls.Add($cb)
    return $cb
}

# Checkboxes (left side)
$chkChrome       = New-CheckBox "Install Chrome"              20  20
$chkQuickBid     = New-CheckBox "Install QuickBid"            20  50
$chkTakeoff      = New-CheckBox "Install Takeoff"             20  80
$chkBluebeam     = New-CheckBox "Install Bluebeam 21"         20 110
$chkOffice365    = New-CheckBox "Install Office 365 (ODT)"    20 140
$chkTeams        = New-CheckBox "Install Teams"               20 170
$chkDebloat      = New-CheckBox "Run Debloat (may show UI)"   20 210
$chkRmOfficeApps = New-CheckBox "Remove preloaded Office apps"20 240
$chkWinUpdate    = New-CheckBox "Run Windows Update"          20 270

# Install button
$btnInstall = New-Object System.Windows.Forms.Button
$btnInstall.Text = "Install Selected"
$btnInstall.Size = New-Object System.Drawing.Size(240,40)
$btnInstall.Location = New-Object System.Drawing.Point(20,320)

$btnInstall.Add_Click({
    Write-Log "===== Starting selected actions ====="

    if ($chkChrome.Checked)       { Install-Chrome }
    if ($chkQuickBid.Checked)     { Install-QuickBid }
    if ($chkTakeoff.Checked)      { Install-Takeoff }
    if ($chkBluebeam.Checked)     { Install-Bluebeam21 }
    if ($chkOffice365.Checked)    { Install-Office365 }
    if ($chkTeams.Checked)        { Install-Teams }
    if ($chkDebloat.Checked)      { Run-Debloat }
    if ($chkRmOfficeApps.Checked) { Remove-PreloadedOffice }
    if ($chkWinUpdate.Checked)    { Run-WindowsUpdate }

    Write-Log "===== All selected actions finished ====="
    [System.Windows.Forms.MessageBox]::Show("Finished running selected items.","Windows Setup Toolkit") | Out-Null
})

$Form.Controls.Add($btnInstall)

Write-Log "GUI loaded. Check what you want, then click 'Install Selected'."

[void]$Form.ShowDialog()
