# ============================================================
#   Windows Setup Toolkit - GUI Version w/ Logging
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

# ------- Download + Install Helpers -------
function Install-EXE {
    param($Name, $URL, $Args = "/silent /verysilent /norestart")

    Write-Log "Downloading $Name..."
    $Temp = "$env:TEMP\$Name.exe"
    Invoke-WebRequest $URL -OutFile $Temp

    Write-Log "Installing $Name..."
    Start-Process $Temp -ArgumentList $Args -Wait
    Remove-Item $Temp -Force
    Write-Log "$Name installed successfully."
}

function Install-MSI {
    param($Name, $URL, $Args = "/quiet /norestart")

    Write-Log "Downloading $Name..."
    $Temp = "$env:TEMP\$Name.msi"
    Invoke-WebRequest $URL -OutFile $Temp

    Write-Log "Installing $Name..."
    Start-Process "msiexec.exe" -ArgumentList "/i `"$Temp`" $Args" -Wait
    Remove-Item $Temp -Force
    Write-Log "$Name installed successfully."
}

# ------- App Install Routines -------
function Install-Chrome {
    Install-EXE "Chrome" "https://dl.google.com/chrome/install/latest/chrome_installer.exe"
}

function Install-Takeoff {
    Install-EXE "OnScreenTakeoff" "https://downloads.oncenter.com/Downloads/OST/400/OST4.0.0.288Setup.exe"
}

function Install-QuickBid {
    Install-EXE "QuickBid" "https://downloads.oncenter.com/Downloads/QB/499/QB4990516Setup.exe"
}

function Install-Bluebeam21 {
    # Bluebeam provides a universal link that redirects to the correct installer
    Install-EXE "BluebeamRevu21" "https://bluebeam.com/FullRevuTRIAL"
}

function Install-Office365 {
    Install-EXE "ODT" "https://download.microsoft.com/download/2/8/E/28E8EC70-BD0C-4CD1-B447-3B0C10CC9F40/officedeploymenttool.exe"
    Write-Log "Office Deployment Tool installed. Configure configuration.xml for deployment."
}

function Install-Teams {
    Install-EXE "TeamsWork" "https://statics.teams.cdn.office.net/production-windows-x86/lkg/MSTeamsSetup.exe"
}

function Run-Debloat {
    Write-Log "Running Windows Debloat..."
    iwr -useb https://git.io/debloat | iex
    Write-Log "Debloat complete."
}

function Remove-PreloadedOffice {
    Write-Log "Removing bundled Microsoft Office Store apps..."

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

# ============================================================
#                        GUI Setup
# ============================================================

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Windows Setup Toolkit"
$Form.Size = New-Object System.Drawing.Size(650,500)
$Form.StartPosition = "CenterScreen"

# Buttons
function Add-Button {
    param($text, $x, $y, $action)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text
    $btn.Size = New-Object System.Drawing.Size(200,35)
    $btn.Location = New-Object System.Drawing.Point($x,$y)
    $btn.Add_Click($action)
    $Form.Controls.Add($btn)
}

Add-Button "Install Chrome" 20 20 { Install-Chrome }
Add-Button "Install On-Screen Quick Bid" 20 65 { Install-QuickBid }
Add-Button "Install On-Screen Takeoff" 20 110 { Install-Takeoff }
Add-Button "Install Bluebeam Revu 21" 20 155 { Install-Bluebeam21 }
Add-Button "Install Office 365" 20 200 { Install-Office365 }
Add-Button "Install Teams (Work)" 20 245 { Install-Teams }
Add-Button "Run Debloat" 20 290 { Run-Debloat }
Add-Button "Remove Preloaded Office" 20 335 { Remove-PreloadedOffice }
Add-Button "Run Windows Updates" 20 380 { Run-WindowsUpdate }

# Log Box
$LogBox = New-Object System.Windows.Forms.TextBox
$LogBox.Multiline = $true
$LogBox.ScrollBars = "Vertical"
$LogBox.ReadOnly = $true
$LogBox.Size = New-Object System.Drawing.Size(360,390)
$LogBox.Location = New-Object System.Drawing.Point(260,20)
$Form.Controls.Add($LogBox)

# Start GUI
Write-Log "Toolkit started."
$Form.ShowDialog()
