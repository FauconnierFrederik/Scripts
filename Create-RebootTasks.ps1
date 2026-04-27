#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Maakt geplande taken aan voor wekelijkse en eenmalige reboots op Windows Server.

.DESCRIPTION
    Dit script maakt twee geplande taken aan:
      1. WeeklyReboot   - Elke week op een gekozen dag en tijdstip
      2. OnceReboot     - Eenmalige reboot op een opgegeven datum en tijdstip

.NOTES
    Moet uitgevoerd worden als Administrator.
    De taken worden aangemaakt onder het SYSTEM-account.
#>

# -----------------------------------------------
#  CONFIGURATIE - via popup invoer
# -----------------------------------------------

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

$WeeklyDay = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Dag voor de wekelijkse reboot:`n`nKies uit: Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday",
    "Reboot taken - Wekelijkse dag",
    "Sunday"
)
if ([string]::IsNullOrWhiteSpace($WeeklyDay)) { [System.Windows.Forms.MessageBox]::Show("Geen dag ingegeven. Script gestopt.", "Fout", "OK", "Error"); exit }

$WeeklyTime = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Tijdstip voor de wekelijkse reboot (HH:mm, 24-uurs formaat):`n`nVoorbeelden: 02:00  |  03:30",
    "Reboot taken - Wekelijks tijdstip",
    "02:00"
)
if ([string]::IsNullOrWhiteSpace($WeeklyTime)) { [System.Windows.Forms.MessageBox]::Show("Geen tijdstip ingegeven. Script gestopt.", "Fout", "OK", "Error"); exit }

$OnceDate = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Datum voor de eenmalige reboot (yyyy-MM-dd):`n`nVoorbeelden: 2025-12-01  |  2026-03-15",
    "Reboot taken - Eenmalige datum",
    (Get-Date).AddDays(7).ToString("yyyy-MM-dd")
)
if ([string]::IsNullOrWhiteSpace($OnceDate)) { [System.Windows.Forms.MessageBox]::Show("Geen datum ingegeven. Script gestopt.", "Fout", "OK", "Error"); exit }

$OnceTime = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Tijdstip voor de eenmalige reboot (HH:mm, 24-uurs formaat):`n`nVoorbeelden: 03:00  |  22:00",
    "Reboot taken - Eenmalig tijdstip",
    "03:00"
)
if ([string]::IsNullOrWhiteSpace($OnceTime)) { [System.Windows.Forms.MessageBox]::Show("Geen tijdstip ingegeven. Script gestopt.", "Fout", "OK", "Error"); exit }

# Taaknamen (pas aan indien gewenst)
$WeeklyTaskName = "WeeklyServerReboot"
$OnceTaskName   = "OnceServerReboot"

# -----------------------------------------------
#  HULPFUNCTIES
# -----------------------------------------------

function Write-Step    { param([string]$Msg) Write-Host "`n[*] $Msg" -ForegroundColor Cyan }
function Write-Success { param([string]$Msg) Write-Host "    [OK] $Msg" -ForegroundColor Green }
function Write-Warn    { param([string]$Msg) Write-Host "    [!]  $Msg" -ForegroundColor Yellow }

# -----------------------------------------------
#  VALIDATIE
# -----------------------------------------------

Write-Step "Invoer valideren..."

$ValidDays = @("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
if ($WeeklyDay -notin $ValidDays) {
    Write-Error "Ongeldige dag: '$WeeklyDay'. Kies uit: $($ValidDays -join ', ')"
    exit 1
}

try   { [datetime]::ParseExact($WeeklyTime, "HH:mm", $null) | Out-Null }
catch { Write-Error "Ongeldig tijdstip voor wekelijks: '$WeeklyTime'. Verwacht: HH:mm"; exit 1 }

try   { $OnceParsed = [datetime]::ParseExact("$OnceDate $OnceTime", "yyyy-MM-dd HH:mm", $null) }
catch { Write-Error "Ongeldige datum/tijd voor eenmalig: '$OnceDate $OnceTime'"; exit 1 }

if ($OnceParsed -lt (Get-Date)) {
    Write-Warn "Eenmalige rebootdatum ligt in het verleden. Taak wordt aangemaakt maar niet uitgevoerd."
}

Write-Success "Invoer is geldig."

# -----------------------------------------------
#  GEMEENSCHAPPELIJKE TAAKINSTELLINGEN
# -----------------------------------------------

$Action = New-ScheduledTaskAction `
    -Execute  "shutdown.exe" `
    -Argument "/r /f /t 60 /c `"Gepland herstart door taakplanner`""

$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
    -MultipleInstances  IgnoreNew `
    -StartWhenAvailable

$Principal = New-ScheduledTaskPrincipal `
    -UserId    "SYSTEM" `
    -LogonType ServiceAccount `
    -RunLevel  Highest

# -----------------------------------------------
#  TAAK 1 - WEKELIJKSE REBOOT
# -----------------------------------------------

Write-Step "Wekelijkse reboot-taak aanmaken ($WeeklyDay om $WeeklyTime)..."

if (Get-ScheduledTask -TaskName $WeeklyTaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $WeeklyTaskName -Confirm:$false
    Write-Warn "Bestaande taak '$WeeklyTaskName' verwijderd."
}

$WeeklyTrigger = New-ScheduledTaskTrigger `
    -Weekly `
    -DaysOfWeek $WeeklyDay `
    -At $WeeklyTime

$WeeklyDesc = "Wekelijkse server reboot elke $WeeklyDay om $WeeklyTime"

$WeeklyTask = New-ScheduledTask `
    -Action      $Action `
    -Trigger     $WeeklyTrigger `
    -Settings    $Settings `
    -Principal   $Principal `
    -Description $WeeklyDesc

Register-ScheduledTask `
    -TaskName    $WeeklyTaskName `
    -InputObject $WeeklyTask `
    -Force | Out-Null

Write-Success "Taak '$WeeklyTaskName' aangemaakt."

# -----------------------------------------------
#  TAAK 2 - EENMALIGE REBOOT
# -----------------------------------------------

Write-Step "Eenmalige reboot-taak aanmaken ($OnceDate om $OnceTime)..."

if (Get-ScheduledTask -TaskName $OnceTaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $OnceTaskName -Confirm:$false
    Write-Warn "Bestaande taak '$OnceTaskName' verwijderd."
}

$OnceTrigger = New-ScheduledTaskTrigger `
    -Once `
    -At $OnceParsed

$OnceDesc = "Eenmalige server reboot op $OnceDate om $OnceTime"

$OnceTask = New-ScheduledTask `
    -Action      $Action `
    -Trigger     $OnceTrigger `
    -Settings    $Settings `
    -Principal   $Principal `
    -Description $OnceDesc

Register-ScheduledTask `
    -TaskName    $OnceTaskName `
    -InputObject $OnceTask `
    -Force | Out-Null

Write-Success "Taak '$OnceTaskName' aangemaakt."

# -----------------------------------------------
#  SAMENVATTING
# -----------------------------------------------

Write-Host ""
Write-Host "============================================" -ForegroundColor DarkGray
Write-Host "  Taken succesvol aangemaakt" -ForegroundColor White
Write-Host "============================================" -ForegroundColor DarkGray

foreach ($Name in @($WeeklyTaskName, $OnceTaskName)) {
    $Task = Get-ScheduledTask -TaskName $Name
    $Info = Get-ScheduledTaskInfo -TaskName $Name
    Write-Host ""
    Write-Host "  Taak      : $($Task.TaskName)" -ForegroundColor White
    Write-Host "  Status    : $($Task.State)"
    Write-Host "  Volgende  : $($Info.NextRunTime)"
}

Write-Host ""
Write-Host "  Beheer via 'Get-ScheduledTask' of taskschd.msc" -ForegroundColor DarkGray
Write-Host ""
