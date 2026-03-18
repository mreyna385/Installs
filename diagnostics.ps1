if (-not $global:LogFile) { $global:LogFile = 'C:\WinSetupToolkit\install.log' }
$global:DiagLog = 'C:\WinSetupToolkit\diagnostics.log'

if (-not (Test-Path 'C:\WinSetupToolkit')) {
    New-Item -ItemType Directory -Path 'C:\WinSetupToolkit' | Out-Null
}

if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    function Write-Log {
        param([string]$Message, [string]$Level = 'INFO')
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Write-Host "[$ts] [$Level] $Message"
    }
}

function Write-DiagLog {
    param([string]$Message)
    try {
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Add-Content -Path $global:DiagLog -Value "[$ts] $Message" -ErrorAction SilentlyContinue
    } catch { }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# SHARED UI: Fix Menu (WinForms)
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Show-FixMenu {
    param(
        [string]$Title,
        [string]$Summary,
        [array]$FixOptions
    )
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        if (-not $FixOptions -or $FixOptions.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                'No fixes available for issues found.',
                $Title,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }

        $form               = New-Object System.Windows.Forms.Form
        $form.Text          = $Title
        $form.Size          = New-Object System.Drawing.Size(700, 500)
        $form.StartPosition = 'CenterScreen'
        $form.FormBorderStyle = 'FixedDialog'
        $form.MaximizeBox   = $false
        try { $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#1E1E1E') } catch { }

        $lblSummary           = New-Object System.Windows.Forms.Label
        $lblSummary.Text      = $Summary
        $lblSummary.Location  = New-Object System.Drawing.Point(10, 10)
        $lblSummary.Size      = New-Object System.Drawing.Size(665, 50)
        $lblSummary.Font      = New-Object System.Drawing.Font('Segoe UI', 10)
        $lblSummary.ForeColor = [System.Drawing.Color]::White
        $lblSummary.BackColor = [System.Drawing.Color]::Transparent
        $form.Controls.Add($lblSummary)

        $clb            = New-Object System.Windows.Forms.CheckedListBox
        $clb.Location   = New-Object System.Drawing.Point(10, 70)
        $clb.Size       = New-Object System.Drawing.Size(665, 340)
        $clb.DrawMode   = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
        $clb.ItemHeight = 22
        try {
            $clb.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
            $clb.ForeColor = [System.Drawing.Color]::White
        } catch { }

        foreach ($fix in $FixOptions) {
            $text = "[$($fix.Severity)] $($fix.Label)"
            $idx  = $clb.Items.Add($text)
            $clb.SetItemChecked($idx, ($fix.Severity -eq 'CRITICAL' -or $fix.Severity -eq 'WARN'))
        }

        $clb.Add_DrawItem({
            param($drawSender, $e)
            if ($e.Index -lt 0) { return }
            try {
                $isSelected = [bool]($e.State -band [System.Windows.Forms.DrawItemState]::Selected)
                $bgColor    = if ($isSelected) {
                    [System.Drawing.Color]::FromArgb(60, 60, 60)
                } else {
                    [System.Drawing.Color]::FromArgb(30, 30, 30)
                }
                $bgBrush = New-Object System.Drawing.SolidBrush($bgColor)
                $e.Graphics.FillRectangle($bgBrush, $e.Bounds)
                $bgBrush.Dispose()

                $itemText  = $drawSender.Items[$e.Index]
                $textColor = [System.Drawing.Color]::White
                if      ($itemText -match '^\[CRITICAL\]') { $textColor = [System.Drawing.Color]::Red    }
                elseif  ($itemText -match '^\[WARN\]')     { $textColor = [System.Drawing.Color]::Yellow }
                elseif  ($itemText -match '^\[INFO\]')     { $textColor = [System.Drawing.Color]::Cyan   }

                $checkState = if ($drawSender.GetItemChecked($e.Index)) {
                    [System.Windows.Forms.CheckBoxState]::CheckedNormal
                } else {
                    [System.Windows.Forms.CheckBoxState]::UncheckedNormal
                }
                $checkSize  = [System.Windows.Forms.CheckBoxRenderer]::GetGlyphSize($e.Graphics, $checkState)
                $checkPoint = New-Object System.Drawing.Point(
                    ($e.Bounds.X + 2),
                    ($e.Bounds.Y + [int](($e.Bounds.Height - $checkSize.Height) / 2))
                )
                [System.Windows.Forms.CheckBoxRenderer]::DrawCheckBox($e.Graphics, $checkPoint, $checkState)

                $textX    = $e.Bounds.X + $checkSize.Width + 6
                $textW    = $e.Bounds.Right - $textX
                $textRect = New-Object System.Drawing.Rectangle($textX, $e.Bounds.Y, $textW, $e.Bounds.Height)
                $flags    = [System.Windows.Forms.TextFormatFlags]::Left -bor
                            [System.Windows.Forms.TextFormatFlags]::VerticalCenter
                [System.Windows.Forms.TextRenderer]::DrawText(
                    $e.Graphics, $itemText, $e.Font, $textRect, $textColor, $flags
                )
                $e.DrawFocusRectangle()
            } catch { }
        })

        $form.Controls.Add($clb)

        $btnApply           = New-Object System.Windows.Forms.Button
        $btnApply.Text      = 'Apply Selected Fixes'
        $btnApply.BackColor = [System.Drawing.Color]::Green
        $btnApply.ForeColor = [System.Drawing.Color]::White
        $btnApply.Location  = New-Object System.Drawing.Point(448, 420)
        $btnApply.Size      = New-Object System.Drawing.Size(155, 30)
        $form.Controls.Add($btnApply)

        $btnSkip              = New-Object System.Windows.Forms.Button
        $btnSkip.Text         = 'Skip All Fixes'
        $btnSkip.BackColor    = [System.Drawing.Color]::Gray
        $btnSkip.ForeColor    = [System.Drawing.Color]::White
        $btnSkip.Location     = New-Object System.Drawing.Point(310, 420)
        $btnSkip.Size         = New-Object System.Drawing.Size(128, 30)
        $btnSkip.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Controls.Add($btnSkip)
        $form.CancelButton = $btnSkip

        # Store references in Tag so the click handler can reach them via closure
        $form.Tag = [PSCustomObject]@{
            FixOptions = $FixOptions
            Clb        = $clb
            Title      = $Title
        }

        $btnApply.Add_Click({
            $tag             = $form.Tag
            $selectedIndices = @()
            for ($i = 0; $i -lt $tag.Clb.Items.Count; $i++) {
                if ($tag.Clb.GetItemChecked($i)) { $selectedIndices += $i }
            }
            if ($selectedIndices.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    'No fixes selected.',
                    $tag.Title,
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }
            $form.Hide()
            foreach ($selIdx in $selectedIndices) {
                try {
                    & $tag.FixOptions[$selIdx].Action
                    Write-Log "Fix applied: $($tag.FixOptions[$selIdx].Label)" 'SUCCESS'
                } catch {
                    Write-Log "Fix failed '$($tag.FixOptions[$selIdx].Label)': $($_)" 'ERROR'
                }
            }
            [System.Windows.Forms.MessageBox]::Show(
                'Fixes applied. Check console for details.',
                'Fixes Complete',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            $form.Close()
        })

        $form.ShowDialog() | Out-Null
    } catch {
        Write-Log "Show-FixMenu failed: $($_)" 'ERROR'
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# DIAGNOSTIC: Event Log Scan
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Invoke-EventLogScan {
    try {
        Write-Log 'Scanning event logs for errors in the last 24 hours...'
        Write-DiagLog '=== EVENT LOG SCAN ==='
        $cutoff = (Get-Date).AddHours(-24)

        $tips = @{
            41   = 'Unexpected shutdown. Check PSU, run memory test (mdsched.exe).'
            6008 = 'Unexpected shutdown. Check for BSOD in minidump: C:\Windows\Minidump'
            7034 = 'Service crashed unexpectedly. Check service dependencies in services.msc'
            7031 = 'Service terminated. Review Event Log for related service errors.'
            1000 = 'App crash. Note faulting module, check vendor KB or reinstall app.'
            4625 = 'Failed logon. Check for brute force. Review Security log for source IP.'
        }

        $allErrors = @()
        foreach ($logName in @('System', 'Application', 'Security')) {
            try {
                $events = Get-WinEvent -FilterHashtable @{
                    LogName   = $logName
                    Level     = 1, 2
                    StartTime = $cutoff
                } -ErrorAction SilentlyContinue
                if ($events) { $allErrors += $events }
            } catch {
                Write-Log "Could not scan log '$($logName)': $($_)" 'WARN'
            }
        }

        if ($allErrors.Count -eq 0) {
            Write-Host 'No critical errors found in last 24 hours.' -ForegroundColor Green
            Write-DiagLog 'No critical errors found in last 24 hours.'
            return
        }

        Write-Log "Found $($allErrors.Count) error(s) in the last 24 hours." 'WARN'
        Write-DiagLog "Found $($allErrors.Count) error(s) in the last 24 hours."

        $grouped = $allErrors | Group-Object -Property ProviderName | Sort-Object Count -Descending | Select-Object -First 10
        Write-Host ''
        Write-Host 'Top Error Sources (last 24 hours):' -ForegroundColor Yellow
        Write-DiagLog 'Top Error Sources:'
        foreach ($g in $grouped) {
            $line = "  [$($g.Count)x] $($g.Name)"
            Write-Host $line -ForegroundColor Yellow
            Write-DiagLog $line
        }

        Write-Host ''
        Write-Host 'Recent Errors:' -ForegroundColor Yellow
        Write-DiagLog 'Recent Errors:'
        foreach ($evt in ($allErrors | Select-Object -First 25)) {
            $rawMsg = if ($evt.Message) { $evt.Message } else { '(no message)' }
            $msg    = $rawMsg.Substring(0, [Math]::Min(200, $rawMsg.Length))
            $line   = "[$($evt.TimeCreated)] ID:$($evt.Id) Source:$($evt.ProviderName) -- $msg"
            Write-Host $line -ForegroundColor Red
            Write-DiagLog $line
            if ($tips.ContainsKey([int]$evt.Id)) {
                $tip = "  TIP: $($tips[[int]$evt.Id])"
                Write-Host $tip -ForegroundColor Cyan
                Write-DiagLog $tip
            }
        }

        # â”€â”€ Build fix options â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        $fixOptions  = @()
        $foundIds    = $allErrors | ForEach-Object { [int]$_.Id } | Select-Object -Unique

        # EventID 41 â€” Kernel-Power unexpected shutdown
        if ($foundIds -contains 41) {
            $fixOptions += @{
                Label    = 'Run Memory Diagnostic (mdsched.exe)'
                Severity = 'CRITICAL'
                Action   = { Start-Process mdsched.exe }
            }
        }

        # EventIDs 7034 / 7031 â€” service crashes
        $serviceEvents = $allErrors | Where-Object { $_.Id -eq 7034 -or $_.Id -eq 7031 }
        if ($serviceEvents) {
            $serviceNames = $serviceEvents | Select-Object -ExpandProperty ProviderName -Unique
            foreach ($svcName in $serviceNames) {
                $capturedSvc = $svcName
                $fixOptions += @{
                    Label    = "Restart service: $($capturedSvc)"
                    Severity = 'WARN'
                    Action   = {
                        try {
                            Restart-Service -Name $capturedSvc -Force -ErrorAction Stop
                            Write-Log "Service $($capturedSvc) restarted." 'SUCCESS'
                        } catch {
                            Write-Log "Could not restart $($capturedSvc): $($_)" 'ERROR'
                        }
                    }.GetNewClosure()
                }
            }
        }

        # EventID 4625 â€” failed logons
        if ($foundIds -contains 4625) {
            $fixOptions += @{
                Label    = 'Show last 20 failed logon attempts with details'
                Severity = 'WARN'
                Action   = {
                    Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4625 } `
                        -MaxEvents 20 -ErrorAction SilentlyContinue |
                        ForEach-Object {
                            $msgShort = $_.Message.Substring(0, [Math]::Min(300, $_.Message.Length))
                            Write-Host "Time: $($_.TimeCreated) | Message: $($msgShort)" -ForegroundColor Yellow
                        }
                }
            }
        }

        # EventID 1000 â€” application errors
        if ($foundIds -contains 1000) {
            $fixOptions += @{
                Label    = 'Open Application Event Log for detailed review'
                Severity = 'INFO'
                Action   = { Start-Process eventvwr.msc -ArgumentList '/c:Application' }
            }
        }

        Show-FixMenu `
            -Title      'Event Log Scan -- Fix Options' `
            -Summary    "$($allErrors.Count) error(s) found in the last 24 hours. Select fixes to apply:" `
            -FixOptions $fixOptions
    } catch {
        Write-Log "Invoke-EventLogScan failed: $($_)" 'ERROR'
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# DIAGNOSTIC: Network Diagnostics
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Invoke-NetworkDiagnostics {
    try {
        Write-Log 'Running network diagnostics...'
        Write-DiagLog '=== NETWORK DIAGNOSTICS ==='

        # Active adapters
        $adapters = Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' }
        Write-Host 'Active Network Adapters:' -ForegroundColor Cyan
        Write-DiagLog 'Active Network Adapters:'
        foreach ($a in $adapters) {
            $ipInfo = Get-NetIPAddress -InterfaceIndex $a.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
            $ipv4   = if ($ipInfo) { $ipInfo.IPAddress } else { 'N/A' }
            $line   = "  $($a.Name) | Status: $($a.Status) | IPv4: $ipv4 | MAC: $($a.MacAddress) | Speed: $($a.LinkSpeed)"
            Write-Host $line -ForegroundColor White
            Write-DiagLog $line
        }

        # Default gateway
        $gw     = (Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
            Sort-Object RouteMetric | Select-Object -First 1).NextHop
        $gwLine = "Default Gateway: $gw"
        Write-Host $gwLine -ForegroundColor White
        Write-DiagLog $gwLine

        # DNS servers
        $dns     = (Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue).ServerAddresses |
            Select-Object -Unique
        $dnsLine = "DNS Servers: $($dns -join ', ')"
        Write-Host $dnsLine -ForegroundColor White
        Write-DiagLog $dnsLine

        $pingResults = @{}

        # Ping gateway
        if ($gw -and $gw -ne '0.0.0.0') {
            try {
                $gwPing = Test-Connection -ComputerName $gw -Count 2 -ErrorAction Stop
                $gwAvg  = [Math]::Round(($gwPing | Measure-Object -Property ResponseTime -Average).Average, 0)
                $pingResults[$gw] = $gwAvg
                $pline = "  Ping Gateway ($($gw)): $($gwAvg)ms"
                Write-Host $pline -ForegroundColor Green
                Write-DiagLog $pline
            } catch {
                $pingResults[$gw] = $null
                $pline = "  Ping Gateway ($($gw)): FAILED"
                Write-Host $pline -ForegroundColor Red
                Write-DiagLog $pline
            }
        }

        # Ping standard targets
        foreach ($target in @('8.8.8.8', '1.1.1.1', 'google.com')) {
            try {
                $ping = Test-Connection -ComputerName $target -Count 2 -ErrorAction Stop
                $avg  = [Math]::Round(($ping | Measure-Object -Property ResponseTime -Average).Average, 0)
                $pingResults[$target] = $avg
                $pline = "  Ping $($target): $($avg)ms"
                Write-Host $pline -ForegroundColor Green
                Write-DiagLog $pline
            } catch {
                $pingResults[$target] = $null
                $pline = "  Ping $($target): FAILED"
                Write-Host $pline -ForegroundColor Red
                Write-DiagLog $pline
            }
        }

        # Public IP
        try {
            $pubIp  = (Invoke-WebRequest -Uri 'https://api.ipify.org' -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop).Content.Trim()
            $ipLine = "Public IP: $pubIp"
            Write-Host $ipLine -ForegroundColor White
            Write-DiagLog $ipLine
        } catch {
            $ipLine = 'Public IP: Could not retrieve'
            Write-Host $ipLine -ForegroundColor Yellow
            Write-DiagLog $ipLine
        }

        # Conditional advice
        Write-Host ''
        Write-Host 'Network Advice:' -ForegroundColor Cyan
        Write-DiagLog 'Network Advice:'
        $gwOk       = $gw -and $null -ne $pingResults[$gw]
        $dnsFailing = $null -eq $pingResults['google.com'] -and $null -ne $pingResults['8.8.8.8']
        $anyFailed  = ($null -eq $pingResults['8.8.8.8']) -or
                      ($null -eq $pingResults['1.1.1.1']) -or
                      ($null -eq $pingResults['google.com'])

        if ($dnsFailing) {
            $advice = 'DNS resolution failing -- check DNS server settings'
            Write-Host "  WARN: $advice" -ForegroundColor Yellow
            Write-DiagLog "  WARN: $advice"
        } elseif ($null -eq $pingResults['8.8.8.8'] -and $gwOk) {
            $advice = 'Internet routing issue -- check modem/ISP'
            Write-Host "  WARN: $advice" -ForegroundColor Yellow
            Write-DiagLog "  WARN: $advice"
        } elseif (-not $gwOk) {
            $advice = 'Local network issue -- check cable/WiFi and router'
            Write-Host "  WARN: $advice" -ForegroundColor Yellow
            Write-DiagLog "  WARN: $advice"
        } else {
            Write-Host '  Network appears healthy.' -ForegroundColor Green
            Write-DiagLog '  Network appears healthy.'
        }

        Write-Log 'Network diagnostics complete.' 'SUCCESS'

        # â”€â”€ Build fix options â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        $fixOptions = @()

        $fixOptions += @{
            Label    = 'Flush DNS + Reset Winsock'
            Severity = 'INFO'
            Action   = { Invoke-FlushDNS }
        }

        if ($dnsFailing) {
            $fixOptions += @{
                Label    = 'Set DNS to Google (8.8.8.8 / 8.8.4.4) on all adapters'
                Severity = 'WARN'
                Action   = {
                    Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | ForEach-Object {
                        try {
                            Set-DnsClientServerAddress -InterfaceIndex $_.ifIndex `
                                -ServerAddresses ('8.8.8.8', '8.8.4.4') -ErrorAction Stop
                            Write-Log "DNS set to Google on adapter: $($_.Name)" 'SUCCESS'
                        } catch {
                            Write-Log "Failed to set DNS on $($_.Name): $($_)" 'ERROR'
                        }
                    }
                }
            }
        }

        if ($anyFailed) {
            $fixOptions += @{
                Label    = 'Open Network Troubleshooter'
                Severity = 'INFO'
                Action   = { Start-Process msdt.exe -ArgumentList '/id NetworkDiagnosticsNetworkAdapter' }
            }
        }

        Show-FixMenu `
            -Title      'Network Diagnostics -- Fix Options' `
            -Summary    'Network scan complete. Select fixes to apply:' `
            -FixOptions $fixOptions
    } catch {
        Write-Log "Invoke-NetworkDiagnostics failed: $($_)" 'ERROR'
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# DIAGNOSTIC: Disk Health Check
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Invoke-DiskHealthCheck {
    try {
        Write-Log 'Running disk health check...'
        Write-DiagLog '=== DISK HEALTH CHECK ==='

        $unhealthyDisks = @()
        $lowSpaceDrives = @()

        Write-Host 'Physical Disks:' -ForegroundColor Cyan
        Write-DiagLog 'Physical Disks:'
        $physDisks = Get-PhysicalDisk -ErrorAction Stop
        foreach ($disk in $physDisks) {
            $sizeGB = [Math]::Round($disk.Size / 1GB, 1)
            $line   = "  $($disk.FriendlyName) | Type: $($disk.MediaType) | Status: $($disk.OperationalStatus) | Health: $($disk.HealthStatus) | Size: $($sizeGB)GB"
            if ($disk.HealthStatus -ne 'Healthy') {
                Write-Host "  CRITICAL: $line" -ForegroundColor Red
                Write-DiagLog "  CRITICAL: $line"
                $unhealthyDisks += $disk
            } else {
                Write-Host $line -ForegroundColor White
                Write-DiagLog $line
            }
        }

        Write-Host ''
        Write-Host 'Logical Drives:' -ForegroundColor Cyan
        Write-DiagLog 'Logical Drives:'
        $drives = Get-PSDrive -PSProvider FileSystem -ErrorAction Stop
        foreach ($drive in $drives) {
            if ($null -eq $drive.Used -or $null -eq $drive.Free) { continue }
            $totalBytes = $drive.Used + $drive.Free
            if ($totalBytes -le 0) { continue }
            $totalGB = [Math]::Round($totalBytes / 1GB, 1)
            $freeGB  = [Math]::Round($drive.Free / 1GB, 1)
            $pctFree = [Math]::Round($drive.Free / $totalBytes * 100, 1)
            $line    = "  $($drive.Name):\ | Total: $($totalGB)GB | Free: $($freeGB)GB | Free%: $($pctFree)%"
            if ($pctFree -lt 5) {
                Write-Host "  CRITICAL: $line" -ForegroundColor Red
                Write-DiagLog "  CRITICAL: $line"
                $lowSpaceDrives += $drive
            } elseif ($pctFree -lt 15) {
                Write-Host "  WARN: $line" -ForegroundColor Yellow
                Write-DiagLog "  WARN: $line"
                $lowSpaceDrives += $drive
            } else {
                Write-Host $line -ForegroundColor White
                Write-DiagLog $line
            }
        }

        # Pending reboot check
        $pendingReboot = $false
        foreach ($regPath in @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
        )) {
            if (Test-Path $regPath -ErrorAction SilentlyContinue) { $pendingReboot = $true }
        }
        try {
            $pfro = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' `
                -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
            if ($pfro) { $pendingReboot = $true }
        } catch { }

        Write-Log 'Disk health check complete.' 'SUCCESS'

        # â”€â”€ Build fix options â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        $fixOptions = @()

        foreach ($d in $lowSpaceDrives) {
            $capturedLetter = $d.Name
            $fixOptions += @{
                Label    = "Run Disk Cleanup on $($capturedLetter):"
                Severity = 'WARN'
                Action   = { Start-Process cleanmgr.exe -ArgumentList "/d $($capturedLetter)" -Wait }.GetNewClosure()
            }
        }

        $fixOptions += @{
            Label    = 'Clear all Temp files (C:\Windows\Temp + user Temp)'
            Severity = 'INFO'
            Action   = {
                $tempPaths = @($env:TEMP, 'C:\Windows\Temp')
                foreach ($tp in $tempPaths) {
                    Get-ChildItem -Path $tp -Recurse -Force -ErrorAction SilentlyContinue |
                        Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                }
                Write-Log 'Temp files cleared.' 'SUCCESS'
            }
        }

        if ($unhealthyDisks.Count -gt 0) {
            $fixOptions += @{
                Label    = 'Open Disk Management for review'
                Severity = 'CRITICAL'
                Action   = { Start-Process diskmgmt.msc }
            }
        }

        if ($pendingReboot) {
            $fixOptions += @{
                Label    = 'Reboot computer now'
                Severity = 'WARN'
                Action   = {
                    Add-Type -AssemblyName System.Windows.Forms
                    $confirm = [System.Windows.Forms.MessageBox]::Show(
                        'Reboot now?', 'Confirm Reboot',
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    )
                    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
                        Restart-Computer -Force
                    }
                }
            }
        }

        Show-FixMenu `
            -Title      'Disk Health -- Fix Options' `
            -Summary    'Disk scan complete. Select fixes to apply:' `
            -FixOptions $fixOptions
    } catch {
        Write-Log "Invoke-DiskHealthCheck failed: $($_)" 'ERROR'
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# DIAGNOSTIC: Software Inventory
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Get-SoftwareInventory {
    try {
        Write-Log 'Collecting software inventory...'
        Write-DiagLog '=== SOFTWARE INVENTORY ==='

        $regPaths = @(
            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )

        $software = foreach ($path in $regPaths) {
            try {
                Get-ItemProperty -Path $path -ErrorAction SilentlyContinue |
                    Where-Object { -not [string]::IsNullOrEmpty($_.DisplayName) } |
                    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
            } catch { }
        }

        $sorted = $software | Sort-Object DisplayName -Unique

        Write-Host ''
        Write-Host "Installed Software ($($sorted.Count) entries):" -ForegroundColor Cyan
        Write-DiagLog "Installed Software ($($sorted.Count) entries):"
        foreach ($app in $sorted) {
            $line = "  $($app.DisplayName) | Version: $($app.DisplayVersion) | Publisher: $($app.Publisher) | Installed: $($app.InstallDate)"
            Write-Host $line -ForegroundColor White
            Write-DiagLog $line
        }

        Write-Host "Total: $($sorted.Count) applications" -ForegroundColor Cyan
        Write-DiagLog "Total: $($sorted.Count) applications"
        Write-Log 'Software inventory complete.' 'SUCCESS'
    } catch {
        Write-Log "Get-SoftwareInventory failed: $($_)" 'ERROR'
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# DIAGNOSTIC: Startup Programs Audit
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Get-StartupAudit {
    try {
        Write-Log 'Auditing startup programs...'
        Write-DiagLog '=== STARTUP PROGRAMS AUDIT ==='

        $startupItems = @()

        try {
            $hklmRun = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -ErrorAction SilentlyContinue
            if ($hklmRun) {
                $hklmRun.PSObject.Properties |
                    Where-Object { $_.Name -notmatch '^PS' } |
                    ForEach-Object {
                        $startupItems += [PSCustomObject]@{ Name = $_.Name; Command = $_.Value; Location = 'HKLM\Run' }
                    }
            }
        } catch { }

        try {
            $hkcuRun = Get-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -ErrorAction SilentlyContinue
            if ($hkcuRun) {
                $hkcuRun.PSObject.Properties |
                    Where-Object { $_.Name -notmatch '^PS' } |
                    ForEach-Object {
                        $startupItems += [PSCustomObject]@{ Name = $_.Name; Command = $_.Value; Location = 'HKCU\Run' }
                    }
            }
        } catch { }

        try {
            $cimStartup = Get-CimInstance Win32_StartupCommand -ErrorAction SilentlyContinue
            foreach ($s in $cimStartup) {
                $startupItems += [PSCustomObject]@{ Name = $s.Name; Command = $s.Command; Location = $s.Location }
            }
        } catch { }

        Write-Host ''
        Write-Host "Startup Programs ($($startupItems.Count) entries):" -ForegroundColor Cyan
        Write-DiagLog "Startup Programs ($($startupItems.Count) entries):"

        $suspiciousItems = @()
        foreach ($item in $startupItems) {
            $suspicious = $false
            $reasons    = @()

            $exePath = $item.Command -replace '^"([^"]+)".*$', '$1'
            $exePath = $exePath -replace '^\s*(\S+\.exe).*$', '$1'
            if ($exePath -match '\.exe' -and -not (Test-Path $exePath -ErrorAction SilentlyContinue)) {
                $suspicious = $true
                $reasons   += 'Executable not found on disk'
            }
            if ($item.Command -match '%TEMP%|AppData\\Local\\Temp') {
                $suspicious = $true
                $reasons   += 'Path in temp directory'
            }
            if ($item.Command -match '[a-f0-9]{8,}\.exe') {
                $suspicious = $true
                $reasons   += 'Random-looking filename'
            }

            $line = "  [$($item.Location)] $($item.Name): $($item.Command)"
            if ($suspicious) {
                $item | Add-Member -NotePropertyName 'Reasons' -NotePropertyValue $reasons -Force
                $suspiciousItems += $item
                Write-Host "  SUSPICIOUS: $line -- Reason: $($reasons -join '; ')" -ForegroundColor Yellow
                Write-DiagLog "  SUSPICIOUS: $line -- Reason: $($reasons -join '; ')"
            } else {
                Write-Host $line -ForegroundColor White
                Write-DiagLog $line
            }
        }

        Write-Log 'Startup audit complete.' 'SUCCESS'

        # â”€â”€ Build fix options â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        $fixOptions = @()

        foreach ($susItem in $suspiciousItems) {
            $capturedName    = $susItem.Name
            $capturedLoc     = $susItem.Location
            $capturedRegPath = if ($capturedLoc -eq 'HKLM\Run') {
                'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
            } elseif ($capturedLoc -eq 'HKCU\Run') {
                'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
            } else {
                $null
            }

            if ($capturedRegPath) {
                $fixOptions += @{
                    Label    = "Disable suspicious startup entry: $($capturedName)"
                    Severity = 'WARN'
                    Action   = {
                        try {
                            Remove-ItemProperty -Path $capturedRegPath -Name $capturedName -ErrorAction Stop
                            Write-Log "Removed startup entry: $($capturedName)" 'SUCCESS'
                        } catch {
                            Write-Log "Could not remove $($capturedName): $($_)" 'ERROR'
                        }
                    }.GetNewClosure()
                }
            }
        }

        $fixOptions += @{
            Label    = 'Open Startup tab in Task Manager'
            Severity = 'INFO'
            Action   = { Start-Process taskmgr.exe }
        }

        Show-FixMenu `
            -Title      'Startup Audit -- Fix Options' `
            -Summary    "$($suspiciousItems.Count) suspicious startup item(s) found. Select fixes to apply:" `
            -FixOptions $fixOptions
    } catch {
        Write-Log "Get-StartupAudit failed: $($_)" 'ERROR'
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# DIAGNOSTIC: System Stats
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Get-SystemStats {
    try {
        Write-Log 'Collecting system statistics...'
        Write-DiagLog '=== SYSTEM STATS ==='

        # CPU
        $cpu     = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1
        $cpuLoad = [int]$cpu.LoadPercentage
        $cpuLine = "CPU: $($cpu.Name) | Cores: $($cpu.NumberOfCores) | Load: $($cpuLoad)%"
        Write-Host $cpuLine -ForegroundColor White
        Write-DiagLog $cpuLine

        # RAM â€” TotalVisibleMemorySize and FreePhysicalMemory are in KB
        $os      = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $totalGB = [Math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
        $freeGB  = [Math]::Round($os.FreePhysicalMemory / 1MB, 2)
        $usedGB  = [Math]::Round($totalGB - $freeGB, 2)
        $ramPct  = if ($totalGB -gt 0) { [Math]::Round($usedGB / $totalGB * 100, 1) } else { 0 }
        $ramLine = "RAM: Total: $($totalGB)GB | Used: $($usedGB)GB | Free: $($freeGB)GB | Usage: $($ramPct)%"
        Write-Host $ramLine -ForegroundColor White
        Write-DiagLog $ramLine

        # Uptime
        $uptime     = (Get-Date) - $os.LastBootUpTime
        $uptimeLine = "Uptime: $($uptime.Days) days, $($uptime.Hours) hours, $($uptime.Minutes) minutes"
        Write-Host $uptimeLine -ForegroundColor White
        Write-DiagLog $uptimeLine

        # OS
        $osLine = "OS: $($os.Caption) | Version: $($os.Version) | Build: $($os.BuildNumber) | Arch: $($os.OSArchitecture)"
        Write-Host $osLine -ForegroundColor White
        Write-DiagLog $osLine

        # Host / User
        $hostLine = "Hostname: $($env:COMPUTERNAME) | User: $($env:USERNAME)"
        Write-Host $hostLine -ForegroundColor White
        Write-DiagLog $hostLine

        # Pending reboot
        $pendingReboot = $false
        foreach ($regPath in @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
        )) {
            if (Test-Path $regPath -ErrorAction SilentlyContinue) { $pendingReboot = $true }
        }
        try {
            $pfro = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' `
                -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
            if ($pfro) { $pendingReboot = $true }
        } catch { }

        if ($pendingReboot) {
            Write-Host 'WARNING: PENDING REBOOT DETECTED' -ForegroundColor Red
            Write-DiagLog 'WARNING: PENDING REBOOT DETECTED'
        }

        Write-Log 'System stats collection complete.' 'SUCCESS'

        # â”€â”€ Build fix options (only if issues exist) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        $fixOptions = @()

        if ($pendingReboot) {
            $fixOptions += @{
                Label    = 'Reboot computer now'
                Severity = 'CRITICAL'
                Action   = { Restart-Computer -Force }
            }
        }

        if ($cpuLoad -gt 85) {
            $fixOptions += @{
                Label    = 'Open Task Manager (CPU view)'
                Severity = 'WARN'
                Action   = { Start-Process taskmgr.exe }
            }
        }

        if ($ramPct -gt 85) {
            $fixOptions += @{
                Label    = 'Open Task Manager (Memory view)'
                Severity = 'WARN'
                Action   = { Start-Process taskmgr.exe }
            }
        }

        if ($fixOptions.Count -gt 0) {
            Show-FixMenu `
                -Title      'System Stats -- Fix Options' `
                -Summary    'System issues detected. Select fixes to apply:' `
                -FixOptions $fixOptions
        }
    } catch {
        Write-Log "Get-SystemStats failed: $($_)" 'ERROR'
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# DIAGNOSTIC: Driver Update Scan
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Invoke-DriverUpdateScan {
    try {
        Write-Log 'Scanning for outdated drivers...'
        Write-DiagLog '=== DRIVER UPDATE SCAN ==='

        $cutoffDate     = (Get-Date).AddYears(-3)
        $drivers        = Get-CimInstance Win32_PnPSignedDriver -ErrorAction Stop |
            Where-Object {
                -not [string]::IsNullOrEmpty($_.DeviceName) -and
                $_.Manufacturer -notmatch 'Microsoft' -and
                $_.DeviceClass -notmatch 'System|Computer'
            } |
            Sort-Object DriverDate

        $outdatedCount   = 0
        $outdatedDrivers = @()

        foreach ($drv in $drivers) {
            $drvDate = $drv.DriverDate
            $isOld   = $drvDate -and $drvDate -lt $cutoffDate
            $dateStr = if ($drvDate) { $drvDate.ToString('yyyy-MM-dd') } else { 'Unknown' }
            $line    = "  $($drv.DeviceName) | Version: $($drv.DriverVersion) | Date: $dateStr | Manufacturer: $($drv.Manufacturer)"
            if ($isOld) {
                $outdatedCount++
                $outdatedDrivers += $drv
                Write-Host "  OUTDATED: $line" -ForegroundColor Yellow
                Write-DiagLog "  OUTDATED: $line"
            } else {
                Write-Host $line -ForegroundColor White
                Write-DiagLog $line
            }
        }

        $summaryColor = if ($outdatedCount -gt 0) { 'Yellow' } else { 'Green' }
        $summary      = "Outdated drivers (older than 3 years): $outdatedCount"
        Write-Host $summary -ForegroundColor $summaryColor
        Write-DiagLog $summary

        $note = 'To update drivers, use Device Manager or visit manufacturer website.'
        Write-Host $note -ForegroundColor Cyan
        Write-DiagLog $note

        Write-Log 'Driver update scan complete.' 'SUCCESS'

        # â”€â”€ Build fix options â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        $fixOptions         = @()
        $addedWindowsUpdate = $false

        foreach ($drv in ($outdatedDrivers | Select-Object -First 10)) {
            $capturedDevName = $drv.DeviceName
            $capturedMfr     = $drv.Manufacturer

            # Option 1: Open Device Manager
            $fixOptions += @{
                Label    = "Open Device Manager to update: $($capturedDevName)"
                Severity = 'INFO'
                Action   = { Start-Process devmgmt.msc }.GetNewClosure()
            }

            # Option 2: Windows Update â€” add only once
            if (-not $addedWindowsUpdate) {
                $fixOptions += @{
                    Label    = 'Run Windows Update (includes driver updates)'
                    Severity = 'INFO'
                    Action   = { Invoke-WindowsUpdate }
                }
                $addedWindowsUpdate = $true
            }

            # Option 3: Google search for driver
            $capturedSearchUrl = 'https://www.google.com/search?q=' +
                (($capturedMfr + ' ' + $capturedDevName + ' driver download') -replace ' ', '+')
            $capturedDevLabel  = $capturedDevName
            $fixOptions += @{
                Label    = "Search driver download: $($capturedDevLabel)"
                Severity = 'INFO'
                Action   = { Start-Process $capturedSearchUrl }.GetNewClosure()
            }
        }

        Show-FixMenu `
            -Title      'Driver Update Options' `
            -Summary    "$outdatedCount outdated driver(s) found. Select fixes to apply:" `
            -FixOptions $fixOptions
    } catch {
        Write-Log "Invoke-DriverUpdateScan failed: $($_)" 'ERROR'
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# QUICK FIX: Flush DNS + Reset Winsock
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Invoke-FlushDNS {
    try {
        Write-Log 'Flushing DNS cache...'
        cmd /c 'ipconfig /flushdns' | ForEach-Object { Write-Host $_ }
        Write-Log 'DNS cache flushed.' 'SUCCESS'

        Write-Log 'Resetting Winsock...'
        cmd /c 'netsh winsock reset' | ForEach-Object { Write-Host $_ }
        Write-Log 'Winsock reset complete.' 'SUCCESS'

        Write-Log 'Resetting IP stack...'
        cmd /c 'netsh int ip reset' | ForEach-Object { Write-Host $_ }
        Write-Log 'IP stack reset complete.' 'SUCCESS'

        Write-Log 'DNS flushed and Winsock reset complete. A reboot is recommended.' 'SUCCESS'

        $reboot = Read-Host 'Reboot now? (Y/N)'
        if ($reboot -eq 'Y' -or $reboot -eq 'y') {
            Restart-Computer -Force
        }
    } catch {
        Write-Log "Invoke-FlushDNS failed: $($_)" 'ERROR'
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# QUICK FIX: Reset Print Spooler
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Invoke-PrintSpoolerReset {
    try {
        Write-Log 'Stopping Print Spooler service...'
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        Write-Log 'Print Spooler stopped.'

        Write-Log 'Clearing print queue...'
        $spoolPath = 'C:\Windows\System32\spool\PRINTERS'
        Get-ChildItem -Path $spoolPath -ErrorAction SilentlyContinue |
            Remove-Item -Force -ErrorAction SilentlyContinue

        Write-Log 'Starting Print Spooler service...'
        Start-Service -Name Spooler -ErrorAction Stop
        Write-Log 'Print Spooler restarted. Queue cleared.' 'SUCCESS'
    } catch {
        Write-Log "Invoke-PrintSpoolerReset failed: $($_)" 'ERROR'
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# QUICK FIX: Sync Time (w32tm)
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Invoke-TimeSyncFix {
    try {
        Write-Log 'Fixing Windows Time service...'
        cmd /c 'net stop w32tm'       | ForEach-Object { Write-Host $_ }
        cmd /c 'w32tm /unregister'    | ForEach-Object { Write-Host $_ }
        cmd /c 'w32tm /register'      | ForEach-Object { Write-Host $_ }
        cmd /c 'net start w32tm'      | ForEach-Object { Write-Host $_ }
        cmd /c 'w32tm /resync /force' | ForEach-Object { Write-Host $_ }
        Write-Host ''
        Write-Log 'Time sync status:'
        cmd /c 'w32tm /query /status' | ForEach-Object { Write-Host $_ }
        Write-Log 'Time sync complete.' 'SUCCESS'
    } catch {
        Write-Log "Invoke-TimeSyncFix failed: $($_)" 'ERROR'
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# QUICK FIX: Force Group Policy Refresh
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Invoke-GPUpdate {
    try {
        Write-Log 'Running gpupdate /force -- this may take up to 2 minutes...'
        $result = cmd /c 'gpupdate /force' 2>&1
        foreach ($line in $result) { Write-Host $line }
        Write-Log 'Group Policy refresh complete.' 'SUCCESS'
    } catch {
        Write-Log "Invoke-GPUpdate failed: $($_)" 'ERROR'
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# QUICK FIX: User Profile Cleanup (WinForms GUI)
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Invoke-UserProfileCleanup {
    try {
        Add-Type -AssemblyName System.Windows.Forms

        Write-Log 'Loading user profiles...'
        $profiles = Get-CimInstance Win32_UserProfile -ErrorAction Stop |
            Where-Object { -not $_.Special -and -not $_.Loaded }

        if (-not $profiles -or $profiles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                'No inactive user profiles found.',
                'Profile Cleanup',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }

        # Gather profile details
        $profileData = foreach ($p in $profiles) {
            $username = Split-Path $p.LocalPath -Leaf
            $lastUsed = if ($p.LastUseTime) { $p.LastUseTime.ToString('yyyy-MM-dd') } else { 'Unknown' }
            $sizeMB   = 0
            try {
                $sizeMB = [Math]::Round(
                    (Get-ChildItem -Path $p.LocalPath -Recurse -Force -ErrorAction SilentlyContinue |
                        Measure-Object -Property Length -Sum).Sum / 1MB, 1
                )
            } catch { }
            [PSCustomObject]@{
                Username = $username
                LastUsed = $lastUsed
                SizeMB   = $sizeMB
                Path     = $p.LocalPath
                Profile  = $p
            }
        }

        # Build WinForms window
        $form = New-Object System.Windows.Forms.Form
        $form.Text            = 'User Profile Cleanup'
        $form.Size            = New-Object System.Drawing.Size(750, 520)
        $form.StartPosition   = 'CenterScreen'
        $form.FormBorderStyle = 'FixedDialog'
        $form.MaximizeBox     = $false

        $label          = New-Object System.Windows.Forms.Label
        $label.Text     = 'The following profiles are not currently in use. Select profiles to remove:'
        $label.Location = New-Object System.Drawing.Point(10, 10)
        $label.Size     = New-Object System.Drawing.Size(720, 20)
        $form.Controls.Add($label)

        $grid                              = New-Object System.Windows.Forms.DataGridView
        $grid.Location                     = New-Object System.Drawing.Point(10, 35)
        $grid.Size                         = New-Object System.Drawing.Size(720, 200)
        $grid.ReadOnly                     = $true
        $grid.AllowUserToAddRows           = $false
        $grid.SelectionMode                = 'FullRowSelect'
        $grid.ColumnHeadersHeightSizeMode  = 'AutoSize'
        $null = $grid.Columns.Add('Username', 'Username')
        $null = $grid.Columns.Add('LastUsed', 'Last Used')
        $null = $grid.Columns.Add('SizeMB',   'Size (MB)')
        $null = $grid.Columns.Add('Path',      'Path')
        foreach ($pd in $profileData) {
            $null = $grid.Rows.Add($pd.Username, $pd.LastUsed, $pd.SizeMB, $pd.Path)
        }
        $form.Controls.Add($grid)

        $clbLabel          = New-Object System.Windows.Forms.Label
        $clbLabel.Text     = 'Check profiles to delete (unchecked by default for safety):'
        $clbLabel.Location = New-Object System.Drawing.Point(10, 245)
        $clbLabel.Size     = New-Object System.Drawing.Size(720, 20)
        $form.Controls.Add($clbLabel)

        $clb          = New-Object System.Windows.Forms.CheckedListBox
        $clb.Location = New-Object System.Drawing.Point(10, 270)
        $clb.Size     = New-Object System.Drawing.Size(720, 150)
        foreach ($pd in $profileData) {
            $idx = $clb.Items.Add($pd.Path)
            $clb.SetItemChecked($idx, $false)
        }
        $form.Controls.Add($clb)

        $btnDelete           = New-Object System.Windows.Forms.Button
        $btnDelete.Text      = 'Delete Selected Profiles'
        $btnDelete.BackColor = [System.Drawing.Color]::Red
        $btnDelete.ForeColor = [System.Drawing.Color]::White
        $btnDelete.Location  = New-Object System.Drawing.Point(530, 435)
        $btnDelete.Size      = New-Object System.Drawing.Size(165, 30)
        $form.Controls.Add($btnDelete)

        $btnCancel              = New-Object System.Windows.Forms.Button
        $btnCancel.Text         = 'Cancel'
        $btnCancel.Location     = New-Object System.Drawing.Point(455, 435)
        $btnCancel.Size         = New-Object System.Drawing.Size(65, 30)
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Controls.Add($btnCancel)
        $form.CancelButton = $btnCancel

        $btnDelete.Add_Click({
            $selectedPaths = @()
            for ($i = 0; $i -lt $clb.Items.Count; $i++) {
                if ($clb.GetItemChecked($i)) { $selectedPaths += $clb.Items[$i] }
            }

            if ($selectedPaths.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    'No profiles selected.',
                    'Profile Cleanup',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }

            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Are you sure? This cannot be undone. $($selectedPaths.Count) profile(s) will be permanently deleted.",
                'Confirm Deletion',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

            $deletedCount = 0
            foreach ($path in $selectedPaths) {
                try {
                    $profileObj = $profileData | Where-Object { $_.Path -eq $path } | Select-Object -First 1
                    if ($profileObj) {
                        Remove-CimInstance -InputObject $profileObj.Profile -ErrorAction Stop
                        Write-Log "Deleted user profile: $path" 'SUCCESS'
                        $deletedCount++
                    }
                } catch {
                    Write-Log "Failed to delete profile $($path): $($_)" 'ERROR'
                }
            }

            [System.Windows.Forms.MessageBox]::Show(
                "Deleted $deletedCount profile(s).",
                'Profile Cleanup Complete',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            $form.Close()
        })

        $form.ShowDialog() | Out-Null
    } catch {
        Write-Log "Invoke-UserProfileCleanup failed: $($_)" 'ERROR'
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# FULL DIAGNOSTIC REPORT
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function Invoke-FullDiagnosticReport {
    try {
        Write-Log 'Starting full diagnostic report...'
        Write-DiagLog "=== FULL DIAGNOSTIC REPORT -- $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  SYSTEM STATISTICS' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Get-SystemStats } catch { Write-Log "Get-SystemStats failed in report: $($_)" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  EVENT LOG SCAN' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Invoke-EventLogScan } catch { Write-Log "Invoke-EventLogScan failed in report: $($_)" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  NETWORK DIAGNOSTICS' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Invoke-NetworkDiagnostics } catch { Write-Log "Invoke-NetworkDiagnostics failed in report: $($_)" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  DISK HEALTH CHECK' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Invoke-DiskHealthCheck } catch { Write-Log "Invoke-DiskHealthCheck failed in report: $($_)" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  SOFTWARE INVENTORY' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Get-SoftwareInventory } catch { Write-Log "Get-SoftwareInventory failed in report: $($_)" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  STARTUP PROGRAMS AUDIT' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Get-StartupAudit } catch { Write-Log "Get-StartupAudit failed in report: $($_)" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  DRIVER UPDATE SCAN' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Invoke-DriverUpdateScan } catch { Write-Log "Invoke-DriverUpdateScan failed in report: $($_)" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Log 'Full diagnostic report saved to C:\WinSetupToolkit\diagnostics.log' 'SUCCESS'
    } catch {
        Write-Log "Invoke-FullDiagnosticReport failed: $($_)" 'ERROR'
    }
}

