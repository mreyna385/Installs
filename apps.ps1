if (-not $global:LogFile) { $global:LogFile = 'C:\WinSetupToolkit\install.log' }

if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    function Write-Log {
        param([string]$Message, [string]$Level = 'INFO')
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Write-Host "[$ts] [$Level] $Message"
    }
}

# ─────────────────────────────────────────────
# HELPER: Test if winget is available
# ─────────────────────────────────────────────
function Test-WingetAvailable {
    try {
        return [bool](Get-Command winget -ErrorAction SilentlyContinue)
    } catch {
        Write-Log "Test-WingetAvailable failed: $_" 'ERROR'
        return $false
    }
}

# ─────────────────────────────────────────────
# HELPER: Install via winget
# ─────────────────────────────────────────────
function Install-ViaWinget {
    param(
        [string]$AppName,
        [string]$WingetId
    )
    try {
        Write-Log "Installing $AppName via winget (ID: $WingetId)..."
        winget install --id $WingetId --silent --accept-package-agreements --accept-source-agreements
        if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
            Write-Log "$AppName installed successfully via winget." 'SUCCESS'
            return $true
        } else {
            Write-Log "$AppName winget install returned exit code $LASTEXITCODE." 'WARN'
            return $false
        }
    } catch {
        Write-Log "Install-ViaWinget failed for $($AppName): $($_)" 'ERROR'
        return $false
    }
}

# ─────────────────────────────────────────────
# HELPER: Install via direct EXE download
# ─────────────────────────────────────────────
function Install-EXE {
    param(
        [string]$AppName,
        [string]$Url,
        [string]$SilentArgs = '/S'
    )
    $tmpFile = "$env:TEMP\WinSetup_$($AppName -replace '\s','_').exe"
    try {
        Write-Log "Downloading $AppName from $Url..."
        Invoke-WebRequest -Uri $Url -OutFile $tmpFile -UseBasicParsing -ErrorAction Stop
        Write-Log "Installing $AppName..."
        $proc = Start-Process -FilePath $tmpFile -ArgumentList $SilentArgs -Wait -PassThru -ErrorAction Stop
        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
            Write-Log "$AppName installed successfully." 'SUCCESS'
            return $true
        } else {
            Write-Log "$AppName installer exited with code $($proc.ExitCode)." 'WARN'
            return $false
        }
    } catch {
        Write-Log "Install-EXE failed for $($AppName): $($_)" 'ERROR'
        return $false
    } finally {
        if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    }
}

# ─────────────────────────────────────────────
# HELPER: Install via direct MSI download
# ─────────────────────────────────────────────
function Install-MSI {
    param(
        [string]$AppName,
        [string]$Url,
        [string]$SilentArgs = '/qn /norestart'
    )
    $tmpFile = "$env:TEMP\WinSetup_$($AppName -replace '\s','_').msi"
    try {
        Write-Log "Downloading $AppName MSI from $Url..."
        Invoke-WebRequest -Uri $Url -OutFile $tmpFile -UseBasicParsing -ErrorAction Stop
        Write-Log "Installing $AppName via msiexec..."
        $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList "/i `"$tmpFile`" $SilentArgs" -Wait -PassThru -ErrorAction Stop
        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
            Write-Log "$AppName installed successfully." 'SUCCESS'
            return $true
        } else {
            Write-Log "$AppName MSI installer exited with code $($proc.ExitCode)." 'WARN'
            return $false
        }
    } catch {
        Write-Log "Install-MSI failed for $($AppName): $($_)" 'ERROR'
        return $false
    } finally {
        if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    }
}

# ─────────────────────────────────────────────
# APP: Google Chrome
# ─────────────────────────────────────────────
function Install-Chrome {
    try {
        Write-Log 'Starting Google Chrome installation...'
        if (Test-WingetAvailable) {
            $result = Install-ViaWinget -AppName 'Google Chrome' -WingetId 'Google.Chrome'
            if ($result) { return $true }
            Write-Log 'Winget failed for Chrome. Falling back to direct download...' 'WARN'
        }
        return Install-EXE -AppName 'Google Chrome' `
            -Url 'https://dl.google.com/chrome/install/latest/chrome_installer.exe' `
            -SilentArgs '/silent /install'
    } catch {
        Write-Log "Install-Chrome failed: $_" 'ERROR'
        return $false
    }
}

# ─────────────────────────────────────────────
# APP: Mozilla Firefox
# ─────────────────────────────────────────────
function Install-Firefox {
    try {
        Write-Log 'Starting Mozilla Firefox installation...'
        if (Test-WingetAvailable) {
            $result = Install-ViaWinget -AppName 'Mozilla Firefox' -WingetId 'Mozilla.Firefox'
            if ($result) { return $true }
            Write-Log 'Winget failed for Firefox. Falling back to direct download...' 'WARN'
        }
        return Install-EXE -AppName 'Mozilla Firefox' `
            -Url 'https://download.mozilla.org/?product=firefox-latest&os=win64&lang=en-US' `
            -SilentArgs '/S'
    } catch {
        Write-Log "Install-Firefox failed: $_" 'ERROR'
        return $false
    }
}

# ─────────────────────────────────────────────
# APP: Microsoft 365
# ─────────────────────────────────────────────
function Install-M365 {
    try {
        Write-Log 'Starting Microsoft 365 installation...'
        if (Test-WingetAvailable) {
            $result = Install-ViaWinget -AppName 'Microsoft 365' -WingetId 'Microsoft.Office'
            if ($result) { return $true }
        }
        Write-Log 'Microsoft 365 requires the Office Deployment Tool for manual installation.' 'WARN'
        Write-Log 'Download ODT from: https://www.microsoft.com/en-us/download/details.aspx?id=49117' 'WARN'
        return $false
    } catch {
        Write-Log "Install-M365 failed: $_" 'ERROR'
        return $false
    }
}

# ─────────────────────────────────────────────
# APP: Microsoft Teams
# ─────────────────────────────────────────────
function Install-Teams {
    try {
        Write-Log 'Starting Microsoft Teams installation...'
        if (Test-WingetAvailable) {
            $result = Install-ViaWinget -AppName 'Microsoft Teams' -WingetId 'Microsoft.Teams'
            if ($result) { return $true }
            Write-Log 'Winget failed for Teams. Falling back to direct download...' 'WARN'
        }
        return Install-EXE -AppName 'Microsoft Teams' `
            -Url 'https://statics.teams.cdn.office.net/production-windows-x64/enterprise/webview2/lkg/MSTeamsSetup.exe' `
            -SilentArgs '-s'
    } catch {
        Write-Log "Install-Teams failed: $_" 'ERROR'
        return $false
    }
}

# ─────────────────────────────────────────────
# APP: Adobe Acrobat Reader
# ─────────────────────────────────────────────
function Install-AdobeReader {
    try {
        Write-Log 'Starting Adobe Acrobat Reader installation...'
        if (Test-WingetAvailable) {
            $result = Install-ViaWinget -AppName 'Adobe Acrobat Reader' -WingetId 'Adobe.Acrobat.Reader.64-bit'
            if ($result) { return $true }
        }
        Write-Log 'Adobe Acrobat Reader requires a dynamic download URL. Please install manually from https://get.adobe.com/reader/' 'WARN'
        return $false
    } catch {
        Write-Log "Install-AdobeReader failed: $_" 'ERROR'
        return $false
    }
}

# ─────────────────────────────────────────────
# APP: On Center Takeoff
# ─────────────────────────────────────────────
function Install-Takeoff {
    try {
        Write-Log 'Starting On Center Takeoff installation...'
        return Install-EXE -AppName 'On Center Takeoff' `
            -Url 'https://downloads.oncenter.com/Downloads/OST/400/OST4.0.0.288Setup.exe' `
            -SilentArgs '/s /v"/qn"'
    } catch {
        Write-Log "Install-Takeoff failed: $_" 'ERROR'
        return $false
    }
}

# ─────────────────────────────────────────────
# APP: On Center QuickBid
# ─────────────────────────────────────────────
function Install-QuickBid {
    try {
        Write-Log 'Starting On Center QuickBid installation...'
        return Install-EXE -AppName 'On Center QuickBid' `
            -Url 'https://downloads.oncenter.com/Downloads/QB/499/QB4990516Setup.exe' `
            -SilentArgs '/s /v"/qn"'
    } catch {
        Write-Log "Install-QuickBid failed: $_" 'ERROR'
        return $false
    }
}

# ─────────────────────────────────────────────
# APP: Bluebeam Revu 21 (licensed — prompt for URL)
# ─────────────────────────────────────────────
function Install-Bluebeam {
    try {
        Write-Log 'Bluebeam Revu 21 requires a licensed installer URL.'
        $installerUrl = Read-Host 'Enter the Bluebeam installer URL (leave blank to cancel)'
        if ([string]::IsNullOrWhiteSpace($installerUrl)) {
            Write-Log 'Bluebeam installation cancelled — no URL provided.' 'WARN'
            return $false
        }
        if ($installerUrl -match '\.msi$') {
            return Install-MSI -AppName 'Bluebeam Revu 21' -Url $installerUrl
        } else {
            return Install-EXE -AppName 'Bluebeam Revu 21' -Url $installerUrl
        }
    } catch {
        Write-Log "Install-Bluebeam failed: $_" 'ERROR'
        return $false
    }
}

# ─────────────────────────────────────────────
# APP: Remove Preloaded Office/Teams AppX
# ─────────────────────────────────────────────
function Remove-PreloadedOfficeTeams {
    try {
        Write-Log 'Removing preloaded Office and Teams AppX packages...'
        $packages = @(
            'Microsoft.OfficeHub',
            'Microsoft.MicrosoftOfficeHub',
            'Microsoft.Office.OneNote',
            'MicrosoftTeams'
        )
        foreach ($pkg in $packages) {
            try {
                $appx = Get-AppxPackage -Name $pkg -AllUsers -ErrorAction SilentlyContinue
                if ($appx) {
                    $appx | Remove-AppxPackage -ErrorAction SilentlyContinue
                    Write-Log "Removed AppX package: $pkg" 'SUCCESS'
                } else {
                    Write-Log "AppX package not found (may already be removed): $pkg" 'INFO'
                }
            } catch {
                Write-Log "Failed to remove AppX package $($pkg): $($_)" 'WARN'
            }
            try {
                $provisioned = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName -eq $pkg }
                if ($provisioned) {
                    $provisioned | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
                    Write-Log "Removed provisioned package: $pkg" 'SUCCESS'
                } else {
                    Write-Log "Provisioned package not found: $pkg" 'INFO'
                }
            } catch {
                Write-Log "Failed to remove provisioned package $($pkg): $($_)" 'WARN'
            }
        }
        Write-Log 'Preloaded Office/Teams AppX removal complete.' 'SUCCESS'
    } catch {
        Write-Log "Remove-PreloadedOfficeTeams failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# UPDATES: Windows Update (WinForms GUI)
# ─────────────────────────────────────────────
function Invoke-WindowsUpdate {
    try {
        Add-Type -AssemblyName System.Windows.Forms

        # Step 1: Ensure PSWindowsUpdate module is present
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-Log 'PSWindowsUpdate module not found. Installing...'
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop
            Install-Module PSWindowsUpdate -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
        }

        Import-Module PSWindowsUpdate -ErrorAction Stop
        Write-Log 'Scanning for updates, please wait...'
        $updates = Get-WindowsUpdate -AcceptAll

        if (-not $updates -or $updates.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                'No updates available. System is up to date.',
                'Windows Update',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }

        # Build WinForms window
        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'Windows Updates Available'
        $form.Size = New-Object System.Drawing.Size(700, 500)
        $form.StartPosition = 'CenterScreen'
        $form.FormBorderStyle = 'FixedDialog'
        $form.MaximizeBox = $false

        $label = New-Object System.Windows.Forms.Label
        $label.Text = "$($updates.Count) update(s) available. Select updates to install:"
        $label.Location = New-Object System.Drawing.Point(10, 10)
        $label.Size = New-Object System.Drawing.Size(660, 20)
        $form.Controls.Add($label)

        $clb = New-Object System.Windows.Forms.CheckedListBox
        $clb.Location = New-Object System.Drawing.Point(10, 35)
        $clb.Size = New-Object System.Drawing.Size(660, 360)
        foreach ($u in $updates) {
            $idx = $clb.Items.Add("KB$($u.KB) -- $($u.Title)")
            $clb.SetItemChecked($idx, $true)
        }
        $form.Controls.Add($clb)

        $btnInstall = New-Object System.Windows.Forms.Button
        $btnInstall.Text = 'Install Selected'
        $btnInstall.BackColor = [System.Drawing.Color]::Green
        $btnInstall.ForeColor = [System.Drawing.Color]::White
        $btnInstall.Location = New-Object System.Drawing.Point(478, 415)
        $btnInstall.Size = New-Object System.Drawing.Size(120, 30)
        $form.Controls.Add($btnInstall)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = 'Cancel'
        $btnCancel.Location = New-Object System.Drawing.Point(607, 415)
        $btnCancel.Size = New-Object System.Drawing.Size(65, 30)
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Controls.Add($btnCancel)
        $form.CancelButton = $btnCancel

        $btnInstall.Add_Click({
            $selectedKBs = @()
            for ($i = 0; $i -lt $clb.Items.Count; $i++) {
                if ($clb.GetItemChecked($i)) {
                    $item = $clb.Items[$i]
                    if ($item -match '^KB(\d+)') {
                        $selectedKBs += $Matches[1]
                    }
                }
            }
            if ($selectedKBs.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    'No updates selected.',
                    'Windows Update',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }
            $form.Hide()
            Get-WindowsUpdate -KBArticleID $selectedKBs -Install -AcceptAll -IgnoreReboot
            [System.Windows.Forms.MessageBox]::Show(
                "Installed $($selectedKBs.Count) update(s). Reboot may be required.",
                'Windows Update Complete',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            $form.Close()
        })

        $form.ShowDialog() | Out-Null
    } catch {
        Write-Log "Invoke-WindowsUpdate failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# RUN ALL INSTALLS
# ─────────────────────────────────────────────
function Invoke-RunAll {
    try {
        Write-Log 'Starting RUN ALL INSTALLS sequence...'

        $results = [ordered]@{
            'Google Chrome'            = $false
            'Mozilla Firefox'          = $false
            'Microsoft 365'            = $false
            'Microsoft Teams'          = $false
            'Adobe Acrobat Reader'     = $false
            'On Center Takeoff'        = $false
            'On Center QuickBid'       = $false
            'Bluebeam Revu 21'         = $false
            'Remove Office/Teams AppX' = $false
        }

        try { $results['Google Chrome']            = Install-Chrome }          catch { Write-Log "Chrome failed in RunAll: $_" 'ERROR' }
        try { $results['Mozilla Firefox']          = Install-Firefox }         catch { Write-Log "Firefox failed in RunAll: $_" 'ERROR' }
        try { $results['Microsoft 365']            = Install-M365 }            catch { Write-Log "M365 failed in RunAll: $_" 'ERROR' }
        try { $results['Microsoft Teams']          = Install-Teams }           catch { Write-Log "Teams failed in RunAll: $_" 'ERROR' }
        try { $results['Adobe Acrobat Reader']     = Install-AdobeReader }     catch { Write-Log "Adobe failed in RunAll: $_" 'ERROR' }
        try { $results['On Center Takeoff']        = Install-Takeoff }         catch { Write-Log "Takeoff failed in RunAll: $_" 'ERROR' }
        try { $results['On Center QuickBid']       = Install-QuickBid }        catch { Write-Log "QuickBid failed in RunAll: $_" 'ERROR' }
        try { $results['Bluebeam Revu 21']         = Install-Bluebeam }        catch { Write-Log "Bluebeam failed in RunAll: $_" 'ERROR' }
        try {
            Remove-PreloadedOfficeTeams
            $results['Remove Office/Teams AppX'] = $true
        } catch { Write-Log "Remove AppX failed in RunAll: $_" 'ERROR' }

        Write-Host ''
        Write-Host ('=' * 52) -ForegroundColor Cyan
        Write-Host '  RUN ALL INSTALLS — SUMMARY' -ForegroundColor Cyan
        Write-Host ('=' * 52) -ForegroundColor Cyan
        foreach ($key in $results.Keys) {
            if ($results[$key]) {
                Write-Host "  [OK]     $key" -ForegroundColor Green
            } else {
                Write-Host "  [FAILED] $key" -ForegroundColor Red
            }
        }
        Write-Host ('=' * 52) -ForegroundColor Cyan
        Write-Log 'RUN ALL INSTALLS sequence complete.' 'SUCCESS'
    } catch {
        Write-Log "Invoke-RunAll failed: $_" 'ERROR'
    }
}
