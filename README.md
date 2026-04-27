# PowerShell Admin Scripts

A collection of universal PowerShell scripts for Windows Server administration. All scripts use popup dialogs for input — no hardcoded customer-specific data.

## Quick Launch

Run the following command in an **elevated PowerShell prompt** on any Windows Server:

```powershell
irm https://raw.githubusercontent.com/FauconnierFrederik/Scripts/main/Install.ps1 | iex
```

This downloads all scripts to `C:\Scripts\batch\` and launches the menu automatically.

## Included Scripts

| Script | Description |
|--------|-------------|
| **Launcher** | Central menu to select and run one or multiple scripts |
| **Setup-DomainController** | Promotes a Windows Server to the first DC in a new AD forest |
| **Setup-RDPV4** | Creates a self-signed certificate, configures RDP-Tcp and generates a signed RDP file |
| **Enable-Archive** | Enables Exchange Online archive for a mailbox and applies a 3-year retention policy |
| **Create-RebootTasks** | Creates weekly and one-time reboot scheduled tasks under the SYSTEM account |
| **Install-CheckPSCMService** | Installs a scheduled task that monitors and auto-restarts a Windows service every hour |
| **MFATypeReport** | Exports a CSV report of MFA registration status for all Microsoft 365 users |
| **Launch-WinUtil** | Launches Chris Titus Tech's WinUtil — software install, Windows tweaks and repairs (requires internet) |
| **Clear-DiskSpace** | Safely frees disk space by removing temp files, logs and caches. Option to schedule weekly. |
| **Get-ServerHealth** | Checks CPU, RAM, disks, services and event log. Exports color-coded HTML report. |

## Requirements

- Windows Server 2016 or later
- PowerShell 5.1+
- Run as Administrator
