#Requires -RunAsAdministrator
# ============================================================
#  Get-ServerHealth.ps1
#  Geeft een snel overzicht van de gezondheid van de server:
#  CPU, geheugen, schijven, services en recente fouten.
#  Optioneel: exporteer rapport als HTML.
# ============================================================

Add-Type -AssemblyName System.Windows.Forms

# ------------------------------------------------------------
#  DREMPELWAARDEN
# ------------------------------------------------------------
$DrempelCPU       = 80   # % - waarschuwing boven deze waarde
$DrempelRAM       = 85   # % - waarschuwing boven deze waarde
$DrempelSchijfWarn = 20  # % vrij - waarschuwing onder deze waarde
$DrempelSchijfKrit = 10  # % vrij - kritiek onder deze waarde
$EventLogUren     = 24   # uur terug voor event log analyse

$KritiekeSvcs = @(
    "EventLog",
    "RpcSs",
    "Dnscache",
    "LanmanServer",
    "LanmanWorkstation",
    "Schedule",
    "Winmgmt",
    "wuauserv"
)

# ------------------------------------------------------------
#  HULPFUNCTIES
# ------------------------------------------------------------

$Rapport = [System.Collections.Generic.List[hashtable]]::new()

function Write-Sectie {
    param([string]$Titel)
    Write-Host ""
    Write-Host "  [$Titel]" -ForegroundColor Cyan
}

function Write-Rij {
    param([string]$Label, [string]$Waarde, [string]$Status = "OK")
    $kleur = switch ($Status) {
        "OK"    { "Green"  }
        "WARN"  { "Yellow" }
        "KRIT"  { "Red"    }
        default { "White"  }
    }
    $tag = switch ($Status) {
        "OK"    { "[OK] " }
        "WARN"  { "[!]  " }
        "KRIT"  { "[!!] " }
        default { "     " }
    }
    $labelPad = $Label.PadRight(18)
    Write-Host "    $tag" -NoNewline -ForegroundColor $kleur
    Write-Host "$labelPad : $Waarde"
    $Rapport.Add(@{ Label = $Label; Waarde = $Waarde; Status = $Status })
}

function Get-StatusKleur {
    param([string]$Status)
    switch ($Status) {
        "OK"   { return "#2ecc71" }
        "WARN" { return "#f39c12" }
        "KRIT" { return "#e74c3c" }
        default { return "#95a5a6" }
    }
}

# ------------------------------------------------------------
#  HEADER
# ------------------------------------------------------------

Clear-Host
$nu = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
Write-Host ""
Write-Host "  ============================================================" -ForegroundColor DarkGray
Write-Host "  Server Health Check" -ForegroundColor Cyan
Write-Host "  $($env:COMPUTERNAME)  |  $nu" -ForegroundColor White
Write-Host "  ============================================================" -ForegroundColor DarkGray

# ------------------------------------------------------------
#  SYSTEEM INFO
# ------------------------------------------------------------

Write-Sectie "SYSTEEM"

$os       = Get-CimInstance Win32_OperatingSystem
$lastBoot = $os.LastBootUpTime
$uptime   = (Get-Date) - $lastBoot
$uptimeTxt = "$([int]$uptime.TotalDays)d $($uptime.Hours)u $($uptime.Minutes)m"

Write-Rij "OS"          $os.Caption                           "INFO"
Write-Rij "Versie"      $os.Version                           "INFO"
Write-Rij "Laatste boot" $lastBoot.ToString("dd/MM/yyyy HH:mm") "INFO"
Write-Rij "Uptime"      $uptimeTxt                            "INFO"

# ------------------------------------------------------------
#  CPU
# ------------------------------------------------------------

Write-Sectie "CPU"

$cpu      = (Get-CimInstance Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average
$cpuStatus = if ($cpu -ge $DrempelCPU) { "KRIT" } elseif ($cpu -ge ($DrempelCPU * 0.75)) { "WARN" } else { "OK" }
Write-Rij "Gebruik" "$cpu%" $cpuStatus

# ------------------------------------------------------------
#  GEHEUGEN
# ------------------------------------------------------------

Write-Sectie "GEHEUGEN"

$totaalRAM  = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
$vrijRAM    = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
$gebruiktRAM = [math]::Round($totaalRAM - $vrijRAM, 1)
$pctRAM     = [math]::Round(($gebruiktRAM / $totaalRAM) * 100)
$ramStatus  = if ($pctRAM -ge $DrempelRAM) { "KRIT" } elseif ($pctRAM -ge ($DrempelRAM * 0.9)) { "WARN" } else { "OK" }

Write-Rij "Totaal"   "$totaalRAM GB"             "INFO"
Write-Rij "Gebruikt" "$gebruiktRAM GB ($pctRAM%)" $ramStatus
Write-Rij "Vrij"     "$vrijRAM GB"               "INFO"

# ------------------------------------------------------------
#  SCHIJVEN
# ------------------------------------------------------------

Write-Sectie "SCHIJVEN"

$schijven = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
foreach ($s in $schijven) {
    $totaal   = [math]::Round($s.Size / 1GB, 1)
    $vrij     = [math]::Round($s.FreeSpace / 1GB, 1)
    $gebruikt = [math]::Round($totaal - $vrij, 1)
    $pctVrij  = [math]::Round(($vrij / $totaal) * 100)
    $pctGebruikt = 100 - $pctVrij
    $status   = if ($pctVrij -le $DrempelSchijfKrit) { "KRIT" } elseif ($pctVrij -le $DrempelSchijfWarn) { "WARN" } else { "OK" }
    Write-Rij "$($s.DeviceID)" "$gebruikt GB / $totaal GB gebruikt ($pctGebruikt%)  -  $vrij GB vrij" $status
}

# ------------------------------------------------------------
#  NETWERK
# ------------------------------------------------------------

Write-Sectie "NETWERK"

$adapters = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
    Where-Object { $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.*" }

foreach ($a in $adapters) {
    $adapter = Get-NetAdapter -InterfaceIndex $a.InterfaceIndex -ErrorAction SilentlyContinue
    if ($adapter -and $adapter.Status -eq "Up") {
        Write-Rij $adapter.Name "$($a.IPAddress)/$($a.PrefixLength)" "OK"
    }
}

# ------------------------------------------------------------
#  SERVICES
# ------------------------------------------------------------

Write-Sectie "SERVICES"

$aantalGestopt = 0
foreach ($svcNaam in ($KritiekeSvcs | Sort-Object)) {
    $svc = Get-Service -Name $svcNaam -ErrorAction SilentlyContinue
    if (-not $svc) {
        Write-Rij $svcNaam "Niet gevonden" "WARN"
    } elseif ($svc.Status -eq "Running") {
        Write-Rij $svcNaam "Actief" "OK"
    } else {
        Write-Rij $svcNaam "GESTOPT ($($svc.Status))" "KRIT"
        $aantalGestopt++
    }
}

# ------------------------------------------------------------
#  EVENT LOG
# ------------------------------------------------------------

Write-Sectie "EVENT LOG (laatste $EventLogUren uur)"

$grens = (Get-Date).AddHours(-$EventLogUren)
$events = Get-WinEvent -FilterHashtable @{
    LogName   = "System", "Application"
    Level     = 1, 2
    StartTime = $grens
} -ErrorAction SilentlyContinue | Select-Object -First 50

if (-not $events -or $events.Count -eq 0) {
    Write-Rij "Geen fouten" "Geen kritieke of fout-events gevonden" "OK"
} else {
    $kritiek = ($events | Where-Object { $_.Level -eq 1 }).Count
    $fouten  = ($events | Where-Object { $_.Level -eq 2 }).Count
    if ($kritiek -gt 0) { Write-Rij "Kritiek"  "$kritiek event(s)" "KRIT" }
    if ($fouten  -gt 0) { Write-Rij "Fouten"   "$fouten event(s)"  "WARN" }

    Write-Host ""
    Write-Host "    Laatste 5 events:" -ForegroundColor DarkGray
    $events | Select-Object -First 5 | ForEach-Object {
        $lvl   = if ($_.Level -eq 1) { "[KRIT]" } else { "[FOUT]" }
        $kleur = if ($_.Level -eq 1) { "Red" } else { "Yellow" }
        $tijd  = $_.TimeCreated.ToString("dd/MM HH:mm")
        Write-Host "      $lvl $tijd  $($_.ProviderName): $($_.Message.Split("`n")[0].Substring(0, [math]::Min(80, $_.Message.Length)))" -ForegroundColor $kleur
    }
}

# ------------------------------------------------------------
#  SAMENVATTING
# ------------------------------------------------------------

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor DarkGray

$aantalKrit = ($Rapport | Where-Object { $_.Status -eq "KRIT" }).Count
$aantalWarn = ($Rapport | Where-Object { $_.Status -eq "WARN" }).Count

if ($aantalKrit -gt 0) {
    $ovStatus = "KRITIEK"
    $ovKleur  = "Red"
} elseif ($aantalWarn -gt 0) {
    $ovStatus = "WAARSCHUWING"
    $ovKleur  = "Yellow"
} else {
    $ovStatus = "OK"
    $ovKleur  = "Green"
}

Write-Host "  Eindstatus  : " -NoNewline
Write-Host $ovStatus -ForegroundColor $ovKleur
Write-Host "  Kritiek     : $aantalKrit  |  Waarschuwing: $aantalWarn" -ForegroundColor White
Write-Host "  ============================================================" -ForegroundColor DarkGray

# ------------------------------------------------------------
#  HTML EXPORT (optioneel)
# ------------------------------------------------------------

$exporteren = [System.Windows.Forms.MessageBox]::Show(
    "Health check voltooid.`n`nStatus  : $ovStatus`nKritiek : $aantalKrit`nWaarsch.: $aantalWarn`n`nWil je een HTML-rapport exporteren?",
    "Server Health - $($env:COMPUTERNAME)", "YesNo", "Information"
)

if ($exporteren -eq "Yes") {
    $datum   = Get-Date -Format "yyyyMMdd-HHmm"
    $htmlPad = "C:\Scripts\$($env:COMPUTERNAME)-Health-$datum.html"

    # Overzichtstabel rijen
    $rijen = ($Rapport | ForEach-Object {
        $kleur = Get-StatusKleur $_.Status
        $tag   = switch ($_.Status) { "OK" { "OK" } "WARN" { "!" } "KRIT" { "!!" } default { "-" } }
        "<tr><td style='color:$kleur;font-weight:bold'>[$tag]</td><td>$($_.Label)</td><td>$($_.Waarde)</td></tr>"
    }) -join "`n"

    # Event log tabel rijen
    if ($events -and $events.Count -gt 0) {
        $eventRijen = ($events | ForEach-Object {
            $lvlTekst  = if ($_.Level -eq 1) { "Kritiek" } else { "Fout" }
            $lvlKleur  = if ($_.Level -eq 1) { "#e74c3c" } else { "#f39c12" }
            $tijd      = $_.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss")
            $bericht   = $_.Message -replace "&","&amp;" -replace "<","&lt;" -replace ">","&gt;" -replace "`n","<br>"
            "<tr>
                <td style='color:$lvlKleur;font-weight:bold;white-space:nowrap'>$lvlTekst</td>
                <td style='white-space:nowrap'>$tijd</td>
                <td style='white-space:nowrap'>$($_.LogName)</td>
                <td style='white-space:nowrap'>$($_.Id)</td>
                <td style='white-space:nowrap'>$($_.ProviderName)</td>
                <td style='font-size:12px'>$bericht</td>
            </tr>"
        }) -join "`n"

        $eventSectie = @"
<h3 style='color:#00d4ff;margin-top:40px'>Event Log - Laatste $EventLogUren uur ($($events.Count) events)</h3>
<table>
<tr>
  <th>Level</th>
  <th>Tijdstip</th>
  <th>Log</th>
  <th>Event ID</th>
  <th>Bron</th>
  <th>Bericht</th>
</tr>
$eventRijen
</table>
"@
    } else {
        $eventSectie = "<p style='color:#2ecc71'>Geen kritieke events of fouten gevonden in de laatste $EventLogUren uur.</p>"
    }

    $html = @"
<!DOCTYPE html>
<html lang='nl'>
<head>
<meta charset='UTF-8'>
<title>Server Health - $($env:COMPUTERNAME)</title>
<style>
  body  { font-family: Segoe UI, sans-serif; background:#1a1a2e; color:#eee; margin:40px; }
  h1    { color:#00d4ff; margin-bottom:4px; }
  h3    { color:#00d4ff; }
  .sub  { color:#aaa; font-size:13px; margin-top:0; margin-bottom:20px; }
  table { border-collapse:collapse; width:100%; margin-top:10px; }
  th    { background:#16213e; color:#00d4ff; padding:10px; text-align:left; font-size:13px; }
  td    { padding:8px 10px; border-bottom:1px solid #2a2a4a; vertical-align:top; font-size:13px; }
  tr:hover td { background:#16213e; }
  .badge { display:inline-block; padding:5px 14px; border-radius:4px; font-weight:bold; color:#fff; }
</style>
</head>
<body>
<h1>Server Health Check</h1>
<p class='sub'>$($env:COMPUTERNAME) &nbsp;|&nbsp; $nu</p>

<p>Eindstatus: <span class='badge' style='background:$(Get-StatusKleur $ovStatus)'>$ovStatus</span>
&nbsp;&nbsp; Kritiek: <b>$aantalKrit</b> &nbsp; Waarschuwing: <b>$aantalWarn</b></p>

<h3>Overzicht</h3>
<table>
<tr><th>Status</th><th>Onderdeel</th><th>Waarde</th></tr>
$rijen
</table>

$eventSectie

</body>
</html>
"@

    [System.IO.File]::WriteAllText($htmlPad, $html, [System.Text.Encoding]::UTF8)

    [System.Windows.Forms.MessageBox]::Show(
        "Rapport opgeslagen:`n$htmlPad",
        "HTML Rapport", "OK", "Information"
    )

    Start-Process $htmlPad
}
