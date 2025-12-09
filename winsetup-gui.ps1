# ============================================================
#   Windows Setup Toolkit - GUI Version w/ Logging + Run All
#   (Fixed string error + safer logging)
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------------------------
# GUI SETUP (create UI early so logging can use it)
# ---------------------------------------------

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Windows Setup Toolkit"
$Form.Size = New-Object System.Drawing.Size(900,560)
$Form.StartPosition = "CenterScreen"

# Log box
$LogBox = New-Object System.Windows.Forms.TextBox
$LogBox.Multiline = $true
$LogBox.ScrollBars = "Vertical"
$LogBox.ReadOnly = $true
$LogBox.WordWrap = $false
$LogBox.Size = New-Object System.Drawing.Size(600,480)
$LogBox.Location = New-Object System.Drawing.Point(270,20)
$Form.Controls.Add($LogBox)

# ---------------------------------------------
# Logging System
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

    try {
        Add-Content -Path $LogFile -Value $entry
    } catch {
        # If file is locked or path issues, still try to show in UI
    }

    if ($LogBox) {
        $LogBox.AppendText("$entry`r`n")
    }
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

    try {
        Write-Log "Downloading $Name..."
        $Temp = Join-Path $env:TEMP "$Name.exe"
        Invoke-WebRequest -Uri $URL -OutFile $Temp -UseBasicParsin
