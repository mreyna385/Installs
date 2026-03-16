# Windows Setup Toolkit

A PowerShell-based Windows setup system designed to run from a single command.

## Quick Start

```powershell
iwr -useb https://raw.githubusercontent.com/USERNAME/REPO/main/setup.ps1 | iex
```

> Replace `USERNAME` and `REPO` with your GitHub username and repository name before hosting.

---

## What it does

| Menu item | Method |
|---|---|
| Google Chrome | winget `Google.Chrome` |
| Microsoft Teams | winget `Microsoft.Teams` |
| Microsoft 365 | winget `Microsoft.Office` |
| Takeoff (On Center) | Direct download from oncenter.com |
| QuickBid (On Center) | Direct download from oncenter.com |
| Bluebeam Revu 21 | User-supplied URL (prompted at runtime) |
| Remove preloaded Office/Teams AppX | `Remove-AppxPackage` (all users + provisioned) |
| Windows Update | PSWindowsUpdate module |
| CTT WinUtil | `iwr christitus.com/win \| iex` (confirmation required) |

- winget is preferred for any app in its catalog; direct download is the automatic fallback.
- Every install is wrapped in try/catch; Run All continues on failure and prints a summary.
- Log: `C:\WinSetupToolkit\install.log`

---

## Setup: host on GitHub

1. Create a **public** repository on GitHub.
2. Push both `setup.ps1` and `apps.ps1` to the `main` branch.
3. In **both files**, replace every occurrence of:
   ```
   USERNAME/REPO
   ```
   with your actual `githubusername/reponame`.
4. The raw URLs will then be:
   ```
   https://raw.githubusercontent.com/USERNAME/REPO/main/setup.ps1
   https://raw.githubusercontent.com/USERNAME/REPO/main/apps.ps1
   ```

---

## Local testing (without hosting)

```powershell
# Run as Administrator
Set-ExecutionPolicy Bypass -Scope Process -Force
. .\apps.ps1
# Then call any function directly, e.g.:
Install-Chrome
```

---

## Files

```
setup.ps1                          Entry point — elevation, TLS, menu loop
apps.ps1                           All app definitions and helper functions
.github/workflows/validate.ps1.yml PSScriptAnalyzer CI check
```

---

## Requirements

- Windows 10 1809+ or Windows 11
- PowerShell 5.1+
- Internet access
- Administrator rights (script self-elevates if needed)
