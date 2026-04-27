# ============================================================
#  Install-CheckPSCMService.ps1
#  Schrijft het check-script naar C:\_uvion_batch en
#  registreert de Scheduled Task 'UViON - Check PSCM Service'.
#  Eenmalig uitvoeren als Administrator.
# ============================================================

#Requires -RunAsAdministrator

# ── Configuratie ────────────────────────────────────────────
$TaskName   = "UViON - Check PSCM Service"
$ScriptDir  = "C:\_uvion_batch"
$ScriptPath = "$ScriptDir\Watch-CagService.ps1"
$Description = "Controleert elk uur of de CagService actief is en herstart deze indien nodig."

# ============================================================
#  STAP 1 – Schrijf het check-script naar schijf
# ============================================================

if (-not (Test-Path $ScriptDir)) {
    New-Item -ItemType Directory -Path $ScriptDir -Force | Out-Null
    Write-Host "Map aangemaakt: $ScriptDir" -ForegroundColor Cyan
}

$CheckScriptContent = @'
# ============================================================
#  Watch-CagService.ps1
#  Controleert of de CagService draait en start deze indien nodig.
#  Wordt automatisch uitgevoerd via een Scheduled Task (elk uur).
# ============================================================

$ServiceName  = "CagService"
$LogDir       = "C:\Logs\CagService"
$LogFile      = "$LogDir\CagService-Watch.log"
$MaxLogSizeMB = 5

# --- Zorg dat de logmap bestaat ---
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

# --- Log-rotatie: hernoem het logbestand als het te groot wordt ---
if (Test-Path $LogFile) {
    $SizeMB = (Get-Item $LogFile).Length / 1MB
    if ($SizeMB -gt $MaxLogSizeMB) {
        $Stamp = Get-Date -Format "yyyyMMdd-HHmmss"
        Rename-Item -Path $LogFile -NewName "CagService-Watch-$Stamp.log"
    }
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Line = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $Line
    Write-Output $Line
}

Write-Log "--- Check gestart ---"

# --- Controleer of de service bestaat ---
$Service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue

if ($null -eq $Service) {
    Write-Log "Service '$ServiceName' werd NIET gevonden op dit systeem." "ERROR"
    exit 1
}

Write-Log "Huidige status van '$ServiceName': $($Service.Status)"

# --- Start de service als deze niet draait ---
if ($Service.Status -ne "Running") {
    Write-Log "Service is niet actief. Poging tot starten..." "WARN"
    try {
        Start-Service -Name $ServiceName -ErrorAction Stop
        Start-Sleep -Seconds 5

        $Service.Refresh()
        if ($Service.Status -eq "Running") {
            Write-Log "Service '$ServiceName' succesvol gestart." "INFO"
        } else {
            Write-Log "Service '$ServiceName' kon niet worden gestart. Status: $($Service.Status)" "ERROR"
            exit 2
        }
    } catch {
        Write-Log "Fout bij starten van service: $_" "ERROR"
        exit 3
    }
} else {
    Write-Log "Service '$ServiceName' draait normaal. Geen actie vereist."
}

Write-Log "--- Check afgerond ---"
exit 0
'@

Set-Content -Path $ScriptPath -Value $CheckScriptContent -Encoding UTF8
Write-Host "Check-script weggeschreven naar: $ScriptPath" -ForegroundColor Cyan

# ============================================================
#  STAP 2 – Registreer de Scheduled Task
# ============================================================

# --- Verwijder bestaande taak met dezelfde naam (bij herinstallatie) ---
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Write-Host "Bestaande taak '$TaskName' wordt verwijderd..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# --- Actie: PowerShell uitvoeren met het check-script ---
$Action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NonInteractive -NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""

# --- Trigger: elk uur, startend op het volgende volle uur ---
$Trigger = New-ScheduledTaskTrigger `
    -RepetitionInterval (New-TimeSpan -Hours 1) `
    -Once `
    -At (Get-Date -Minute 0 -Second 0).AddHours(1)

# --- Instellingen ---
$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable:$false

# --- Uitvoeren als SYSTEM met hoogste rechten ---
$Principal = New-ScheduledTaskPrincipal `
    -UserId    "NT AUTHORITY\SYSTEM" `
    -LogonType ServiceAccount `
    -RunLevel  Highest

# --- Taak registreren ---
Register-ScheduledTask `
    -TaskName    $TaskName `
    -Action      $Action `
    -Trigger     $Trigger `
    -Settings    $Settings `
    -Principal   $Principal `
    -Description $Description `
    -Force

Write-Host ""
Write-Host " Scheduled Task '$TaskName' succesvol aangemaakt." -ForegroundColor Green
Write-Host "  Script  : $ScriptPath"                           -ForegroundColor Cyan
Write-Host "  Interval: elk uur"                               -ForegroundColor Cyan
Write-Host "  Account : NT AUTHORITY\SYSTEM"                   -ForegroundColor Cyan
Write-Host ""
Write-Host "Logs worden weggeschreven naar: C:\Logs\CagService\CagService-Watch.log" -ForegroundColor Cyan
