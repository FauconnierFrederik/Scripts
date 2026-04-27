# ============================================================
#  Script Installer
#  Gebruik:
#    irm https://raw.githubusercontent.com/FauconnierFrederik/Scripts/main/Install.ps1 | iex
#
#  Downloadt alle scripts naar C:\Scripts\batch\ en start de launcher.
# ============================================================

$RepoBase  = "https://raw.githubusercontent.com/FauconnierFrederik/Scripts/main"
$LocalBase = "C:\Scripts\batch"

$Scripts = @(
    "Launcher.ps1"
    "MFATypeReport.ps1"
    "Setup-RDPV4.ps1"
    "Enable-Archive.ps1"
    "Create-RebootTasks.ps1"
    "Install-CheckPSCMService.ps1"
    "Setup-DomainController.ps1"
    "Launch-WinUtil.ps1"
    "Clear-DiskSpace.ps1"
    "Get-ServerHealth.ps1"
)

Write-Host ""
Write-Host "  Script Installer" -ForegroundColor Cyan
Write-Host "  $('─' * 40)" -ForegroundColor DarkGray
Write-Host ""

if (-not (Test-Path $LocalBase)) {
    New-Item -ItemType Directory -Path $LocalBase -Force | Out-Null
    Write-Host "  Map aangemaakt: $LocalBase" -ForegroundColor DarkGray
}

foreach ($script in $Scripts) {
    $url  = "$RepoBase/$script"
    $dest = "$LocalBase\$script"
    try {
        $inhoud = Invoke-RestMethod -Uri $url
        [System.IO.File]::WriteAllText($dest, $inhoud, [System.Text.Encoding]::UTF8)
        Write-Host "  [OK] $script" -ForegroundColor Green
    } catch {
        Write-Host "  [!!] $script - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "  Alle scripts gedownload naar $LocalBase" -ForegroundColor Cyan
Write-Host "  Launcher starten..." -ForegroundColor DarkGray
Write-Host ""

& "$LocalBase\Launcher.ps1"
