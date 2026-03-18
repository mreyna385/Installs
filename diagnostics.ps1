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

# ─────────────────────────────────────────────
# DIAGNOSTIC: Event Log Scan
# ─────────────────────────────────────────────
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
                Write-Log "Could not scan log '$logName': $_" 'WARN'
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
            $msg = $rawMsg.Substring(0, [Math]::Min(200, $rawMsg.Length))
            $line = "[$($evt.TimeCreated)] ID:$($evt.Id) Source:$($evt.ProviderName) -- $msg"
            Write-Host $line -ForegroundColor Red
            Write-DiagLog $line
            if ($tips.ContainsKey([int]$evt.Id)) {
                $tip = "  TIP: $($tips[[int]$evt.Id])"
                Write-Host $tip -ForegroundColor Cyan
                Write-DiagLog $tip
            }
        }
    } catch {
        Write-Log "Invoke-EventLogScan failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# DIAGNOSTIC: Network Diagnostics
# ─────────────────────────────────────────────
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
            $ipv4 = if ($ipInfo) { $ipInfo.IPAddress } else { 'N/A' }
            $line = "  $($a.Name) | Status: $($a.Status) | IPv4: $ipv4 | MAC: $($a.MacAddress) | Speed: $($a.LinkSpeed)"
            Write-Host $line -ForegroundColor White
            Write-DiagLog $line
        }

        # Default gateway
        $gw = (Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
            Sort-Object RouteMetric | Select-Object -First 1).NextHop
        $gwLine = "Default Gateway: $gw"
        Write-Host $gwLine -ForegroundColor White
        Write-DiagLog $gwLine

        # DNS servers
        $dns = (Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue).ServerAddresses |
            Select-Object -Unique
        $dnsLine = "DNS Servers: $($dns -join ', ')"
        Write-Host $dnsLine -ForegroundColor White
        Write-DiagLog $dnsLine

        $pingResults = @{}

        # Ping gateway
        if ($gw -and $gw -ne '0.0.0.0') {
            try {
                $gwPing = Test-Connection -ComputerName $gw -Count 2 -ErrorAction Stop
                $gwAvg = [Math]::Round(($gwPing | Measure-Object -Property ResponseTime -Average).Average, 0)
                $pingResults[$gw] = $gwAvg
                $pline = "  Ping Gateway ($gw): ${gwAvg}ms"
                Write-Host $pline -ForegroundColor Green
                Write-DiagLog $pline
            } catch {
                $pingResults[$gw] = $null
                $pline = "  Ping Gateway ($gw): FAILED"
                Write-Host $pline -ForegroundColor Red
                Write-DiagLog $pline
            }
        }

        # Ping targets
        foreach ($target in @('8.8.8.8', '1.1.1.1', 'google.com')) {
            try {
                $ping = Test-Connection -ComputerName $target -Count 2 -ErrorAction Stop
                $avg = [Math]::Round(($ping | Measure-Object -Property ResponseTime -Average).Average, 0)
                $pingResults[$target] = $avg
                $pline = "  Ping ${target}: ${avg}ms"
                Write-Host $pline -ForegroundColor Green
                Write-DiagLog $pline
            } catch {
                $pingResults[$target] = $null
                $pline = "  Ping ${target}: FAILED"
                Write-Host $pline -ForegroundColor Red
                Write-DiagLog $pline
            }
        }

        # Public IP
        try {
            $pubIp = (Invoke-WebRequest -Uri 'https://api.ipify.org' -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop).Content.Trim()
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
        $gwOk = $gw -and $null -ne $pingResults[$gw]
        if ($null -eq $pingResults['google.com'] -and $null -ne $pingResults['8.8.8.8']) {
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
    } catch {
        Write-Log "Invoke-NetworkDiagnostics failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# DIAGNOSTIC: Disk Health Check
# ─────────────────────────────────────────────
function Invoke-DiskHealthCheck {
    try {
        Write-Log 'Running disk health check...'
        Write-DiagLog '=== DISK HEALTH CHECK ==='

        Write-Host 'Physical Disks:' -ForegroundColor Cyan
        Write-DiagLog 'Physical Disks:'
        $physDisks = Get-PhysicalDisk -ErrorAction Stop
        foreach ($disk in $physDisks) {
            $sizeGB = [Math]::Round($disk.Size / 1GB, 1)
            $line = "  $($disk.FriendlyName) | Type: $($disk.MediaType) | Status: $($disk.OperationalStatus) | Health: $($disk.HealthStatus) | Size: ${sizeGB}GB"
            if ($disk.HealthStatus -ne 'Healthy') {
                Write-Host "  CRITICAL: $line" -ForegroundColor Red
                Write-DiagLog "  CRITICAL: $line"
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
            $totalGB  = [Math]::Round($totalBytes / 1GB, 1)
            $freeGB   = [Math]::Round($drive.Free / 1GB, 1)
            $pctFree  = [Math]::Round($drive.Free / $totalBytes * 100, 1)
            $line = "  $($drive.Name):\ | Total: ${totalGB}GB | Free: ${freeGB}GB | Free%: ${pctFree}%"
            if ($pctFree -lt 5) {
                Write-Host "  CRITICAL: $line" -ForegroundColor Red
                Write-DiagLog "  CRITICAL: $line"
            } elseif ($pctFree -lt 15) {
                Write-Host "  WARN: $line" -ForegroundColor Yellow
                Write-DiagLog "  WARN: $line"
            } else {
                Write-Host $line -ForegroundColor White
                Write-DiagLog $line
            }
        }

        Write-Log 'Disk health check complete.' 'SUCCESS'
    } catch {
        Write-Log "Invoke-DiskHealthCheck failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# DIAGNOSTIC: Software Inventory
# ─────────────────────────────────────────────
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
        Write-Log "Get-SoftwareInventory failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# DIAGNOSTIC: Startup Programs Audit
# ─────────────────────────────────────────────
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

        foreach ($item in $startupItems) {
            $suspicious = $false
            $reasons    = @()

            # Extract bare executable path
            $exePath = $item.Command -replace '^"([^"]+)".*$', '$1'
            $exePath = $exePath -replace '^\s*(\S+\.exe).*$', '$1'
            if ($exePath -match '\.exe' -and -not (Test-Path $exePath -ErrorAction SilentlyContinue)) {
                $suspicious = $true
                $reasons += 'Executable not found on disk'
            }

            if ($item.Command -match '%TEMP%|AppData\\Local\\Temp') {
                $suspicious = $true
                $reasons += 'Path in temp directory'
            }

            if ($item.Command -match '[a-f0-9]{8,}\.exe') {
                $suspicious = $true
                $reasons += 'Random-looking filename'
            }

            $line = "  [$($item.Location)] $($item.Name): $($item.Command)"
            if ($suspicious) {
                Write-Host "  SUSPICIOUS: $line -- Reason: $($reasons -join '; ')" -ForegroundColor Yellow
                Write-DiagLog "  SUSPICIOUS: $line -- Reason: $($reasons -join '; ')"
            } else {
                Write-Host $line -ForegroundColor White
                Write-DiagLog $line
            }
        }

        Write-Log 'Startup audit complete.' 'SUCCESS'
    } catch {
        Write-Log "Get-StartupAudit failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# DIAGNOSTIC: System Stats
# ─────────────────────────────────────────────
function Get-SystemStats {
    try {
        Write-Log 'Collecting system statistics...'
        Write-DiagLog '=== SYSTEM STATS ==='

        # CPU
        $cpu = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1
        $cpuLine = "CPU: $($cpu.Name) | Cores: $($cpu.NumberOfCores) | Load: $($cpu.LoadPercentage)%"
        Write-Host $cpuLine -ForegroundColor White
        Write-DiagLog $cpuLine

        # RAM — TotalVisibleMemorySize and FreePhysicalMemory are in KB
        $os      = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $totalGB = [Math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
        $freeGB  = [Math]::Round($os.FreePhysicalMemory / 1MB, 2)
        $usedGB  = [Math]::Round($totalGB - $freeGB, 2)
        $ramPct  = if ($totalGB -gt 0) { [Math]::Round($usedGB / $totalGB * 100, 1) } else { 0 }
        $ramLine = "RAM: Total: ${totalGB}GB | Used: ${usedGB}GB | Free: ${freeGB}GB | Usage: ${ramPct}%"
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
        foreach ($path in @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
        )) {
            if (Test-Path $path -ErrorAction SilentlyContinue) { $pendingReboot = $true }
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
    } catch {
        Write-Log "Get-SystemStats failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# DIAGNOSTIC: Driver Update Scan
# ─────────────────────────────────────────────
function Invoke-DriverUpdateScan {
    try {
        Write-Log 'Scanning for outdated drivers...'
        Write-DiagLog '=== DRIVER UPDATE SCAN ==='

        $cutoffDate = (Get-Date).AddYears(-3)
        $drivers = Get-CimInstance Win32_PnPSignedDriver -ErrorAction Stop |
            Where-Object {
                -not [string]::IsNullOrEmpty($_.DeviceName) -and
                $_.Manufacturer -notmatch 'Microsoft' -and
                $_.DeviceClass -notmatch 'System|Computer'
            } |
            Sort-Object DriverDate

        $outdatedCount = 0
        foreach ($drv in $drivers) {
            $drvDate = $drv.DriverDate
            $isOld   = $drvDate -and $drvDate -lt $cutoffDate
            $dateStr = if ($drvDate) { $drvDate.ToString('yyyy-MM-dd') } else { 'Unknown' }
            $line = "  $($drv.DeviceName) | Version: $($drv.DriverVersion) | Date: $dateStr | Manufacturer: $($drv.Manufacturer)"
            if ($isOld) {
                $outdatedCount++
                Write-Host "  OUTDATED: $line" -ForegroundColor Yellow
                Write-DiagLog "  OUTDATED: $line"
            } else {
                Write-Host $line -ForegroundColor White
                Write-DiagLog $line
            }
        }

        $summary = "Outdated drivers (older than 3 years): $outdatedCount"
        $summaryColor = if ($outdatedCount -gt 0) { 'Yellow' } else { 'Green' }
        Write-Host $summary -ForegroundColor $summaryColor
        Write-DiagLog $summary

        $note = 'To update drivers, use Device Manager or visit manufacturer website.'
        Write-Host $note -ForegroundColor Cyan
        Write-DiagLog $note

        Write-Log 'Driver update scan complete.' 'SUCCESS'
    } catch {
        Write-Log "Invoke-DriverUpdateScan failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# QUICK FIX: Flush DNS + Reset Winsock
# ─────────────────────────────────────────────
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
        Write-Log "Invoke-FlushDNS failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# QUICK FIX: Reset Print Spooler
# ─────────────────────────────────────────────
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
        Write-Log "Invoke-PrintSpoolerReset failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# QUICK FIX: Sync Time (w32tm)
# ─────────────────────────────────────────────
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
        Write-Log "Invoke-TimeSyncFix failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# QUICK FIX: Force Group Policy Refresh
# ─────────────────────────────────────────────
function Invoke-GPUpdate {
    try {
        Write-Log 'Running gpupdate /force -- this may take up to 2 minutes...'
        $result = cmd /c 'gpupdate /force' 2>&1
        foreach ($line in $result) { Write-Host $line }
        Write-Log 'Group Policy refresh complete.' 'SUCCESS'
    } catch {
        Write-Log "Invoke-GPUpdate failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# QUICK FIX: User Profile Cleanup (WinForms GUI)
# ─────────────────────────────────────────────
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
        $form.Text = 'User Profile Cleanup'
        $form.Size = New-Object System.Drawing.Size(750, 520)
        $form.StartPosition = 'CenterScreen'
        $form.FormBorderStyle = 'FixedDialog'
        $form.MaximizeBox = $false

        $label = New-Object System.Windows.Forms.Label
        $label.Text = 'The following profiles are not currently in use. Select profiles to remove:'
        $label.Location = New-Object System.Drawing.Point(10, 10)
        $label.Size = New-Object System.Drawing.Size(720, 20)
        $form.Controls.Add($label)

        $grid = New-Object System.Windows.Forms.DataGridView
        $grid.Location = New-Object System.Drawing.Point(10, 35)
        $grid.Size = New-Object System.Drawing.Size(720, 200)
        $grid.ReadOnly = $true
        $grid.AllowUserToAddRows = $false
        $grid.SelectionMode = 'FullRowSelect'
        $grid.ColumnHeadersHeightSizeMode = 'AutoSize'
        $null = $grid.Columns.Add('Username', 'Username')
        $null = $grid.Columns.Add('LastUsed', 'Last Used')
        $null = $grid.Columns.Add('SizeMB',   'Size (MB)')
        $null = $grid.Columns.Add('Path',      'Path')
        foreach ($pd in $profileData) {
            $null = $grid.Rows.Add($pd.Username, $pd.LastUsed, $pd.SizeMB, $pd.Path)
        }
        $form.Controls.Add($grid)

        $clbLabel = New-Object System.Windows.Forms.Label
        $clbLabel.Text = 'Check profiles to delete (unchecked by default for safety):'
        $clbLabel.Location = New-Object System.Drawing.Point(10, 245)
        $clbLabel.Size = New-Object System.Drawing.Size(720, 20)
        $form.Controls.Add($clbLabel)

        $clb = New-Object System.Windows.Forms.CheckedListBox
        $clb.Location = New-Object System.Drawing.Point(10, 270)
        $clb.Size = New-Object System.Drawing.Size(720, 150)
        foreach ($pd in $profileData) {
            $idx = $clb.Items.Add($pd.Path)
            $clb.SetItemChecked($idx, $false)
        }
        $form.Controls.Add($clb)

        $btnDelete = New-Object System.Windows.Forms.Button
        $btnDelete.Text = 'Delete Selected Profiles'
        $btnDelete.BackColor = [System.Drawing.Color]::Red
        $btnDelete.ForeColor = [System.Drawing.Color]::White
        $btnDelete.Location = New-Object System.Drawing.Point(530, 435)
        $btnDelete.Size = New-Object System.Drawing.Size(165, 30)
        $form.Controls.Add($btnDelete)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = 'Cancel'
        $btnCancel.Location = New-Object System.Drawing.Point(455, 435)
        $btnCancel.Size = New-Object System.Drawing.Size(65, 30)
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Controls.Add($btnCancel)
        $form.CancelButton = $btnCancel

        $btnDelete.Add_Click({
            $selectedPaths = @()
            for ($i = 0; $i -lt $clb.Items.Count; $i++) {
                if ($clb.GetItemChecked($i)) {
                    $selectedPaths += $clb.Items[$i]
                }
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
                    Write-Log "Failed to delete profile $path: $_" 'ERROR'
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
        Write-Log "Invoke-UserProfileCleanup failed: $_" 'ERROR'
    }
}

# ─────────────────────────────────────────────
# FULL DIAGNOSTIC REPORT
# ─────────────────────────────────────────────
function Invoke-FullDiagnosticReport {
    try {
        Write-Log 'Starting full diagnostic report...'
        Write-DiagLog "=== FULL DIAGNOSTIC REPORT -- $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  SYSTEM STATISTICS' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Get-SystemStats } catch { Write-Log "Get-SystemStats failed in report: $_" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  EVENT LOG SCAN' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Invoke-EventLogScan } catch { Write-Log "Invoke-EventLogScan failed in report: $_" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  NETWORK DIAGNOSTICS' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Invoke-NetworkDiagnostics } catch { Write-Log "Invoke-NetworkDiagnostics failed in report: $_" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  DISK HEALTH CHECK' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Invoke-DiskHealthCheck } catch { Write-Log "Invoke-DiskHealthCheck failed in report: $_" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  SOFTWARE INVENTORY' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Get-SoftwareInventory } catch { Write-Log "Get-SoftwareInventory failed in report: $_" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  STARTUP PROGRAMS AUDIT' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Get-StartupAudit } catch { Write-Log "Get-StartupAudit failed in report: $_" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Host '  DRIVER UPDATE SCAN' -ForegroundColor Cyan
        Write-Host ('=' * 60) -ForegroundColor Cyan
        try { Invoke-DriverUpdateScan } catch { Write-Log "Invoke-DriverUpdateScan failed in report: $_" 'ERROR' }

        Write-Host ('=' * 60) -ForegroundColor Cyan
        Write-Log 'Full diagnostic report saved to C:\WinSetupToolkit\diagnostics.log' 'SUCCESS'
    } catch {
        Write-Log "Invoke-FullDiagnosticReport failed: $_" 'ERROR'
    }
}
