#Requires -RunAsAdministrator
# ============================================================
#  Get-ServerHealth.ps1
#  Geeft een snel overzicht van de gezondheid van de server:
#  CPU, geheugen, schijven, services, recente fouten en aanmeldingsbeveiliging.
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
$AanmeldLogUren   = 24   # uur terug voor aanmeldingsanalyse
$DrempelLoginWarn = 5    # waarschuwing vanaf dit aantal mislukte pogingen
$DrempelLoginKrit = 20   # kritiek vanaf dit aantal mislukte pogingen
$MaxAanmeldEvents = 2000 # maximum aantal Security-events om te analyseren

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

function Format-EventTime {
    param(
        [AllowNull()][object]$TimeCreated,
        [string]$Format = "dd/MM HH:mm"
    )

    if ($null -eq $TimeCreated) { return "Onbekend" }
    return ([datetime]$TimeCreated).ToString($Format)
}

function Format-EventMessage {
    param(
        [AllowNull()][string]$Message,
        [switch]$SingleLine,
        [int]$MaxLength = 0,
        [switch]$Html
    )

    $tekst = if ([string]::IsNullOrWhiteSpace($Message)) {
        "(Geen bericht beschikbaar)"
    } else {
        $Message
    }

    if ($SingleLine) { $tekst = ($tekst -split "\r?\n", 2)[0] }
    if ($MaxLength -gt 0 -and $tekst.Length -gt $MaxLength) {
        $tekst = $tekst.Substring(0, $MaxLength)
    }
    if ($Html) {
        $tekst = $tekst -replace "&", "&amp;" -replace "<", "&lt;" -replace ">", "&gt;" -replace "`r`n|`n|`r", "<br>"
    }

    return $tekst
}

function ConvertTo-HtmlText {
    param([AllowNull()][object]$Value)
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function ConvertFrom-SecurityEvent {
    param([Parameter(Mandatory)][object]$EventRecord)

    try {
        [xml]$xml = $EventRecord.ToXml()
    } catch {
        return $null
    }

    $data = @{}
    foreach ($node in @($xml.SelectNodes("//*[local-name()='EventData']/*[local-name()='Data']"))) {
        $name = [string]$node.GetAttribute("Name")
        if ($name) { $data[$name] = [string]$node.InnerText }
    }

    $isPrivilegeEvent = [int]$EventRecord.Id -eq 4672
    $logonType = 0
    [void][int]::TryParse([string]$data["LogonType"], [ref]$logonType)

    [pscustomobject]@{
        Id          = [int]$EventRecord.Id
        TimeCreated = $EventRecord.TimeCreated
        Account     = if ($isPrivilegeEvent) { $data["SubjectUserName"] } else { $data["TargetUserName"] }
        Domain      = if ($isPrivilegeEvent) { $data["SubjectDomainName"] } else { $data["TargetDomainName"] }
        LogonType   = $logonType
        IpAddress   = $data["IpAddress"]
        Workstation = $data["WorkstationName"]
        LogonId     = if ($isPrivilegeEvent) { $data["SubjectLogonId"] } else { $data["TargetLogonId"] }
        Elevated    = $data["ElevatedToken"]
        Privileges  = $data["PrivilegeList"]
        Status      = $data["Status"]
        SubStatus   = $data["SubStatus"]
    }
}

function Test-IsPublicIPAddress {
    param([AllowNull()][string]$Address)

    $ip = $null
    if (-not [System.Net.IPAddress]::TryParse($Address, [ref]$ip)) { return $false }
    if ([System.Net.IPAddress]::IsLoopback($ip)) { return $false }
    if ($ip.IsIPv4MappedToIPv6) { $ip = $ip.MapToIPv4() }

    if ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
        $bytes = $ip.GetAddressBytes()
        if ($bytes[0] -in 0, 10, 127) { return $false }
        if ($bytes[0] -eq 100 -and $bytes[1] -ge 64 -and $bytes[1] -le 127) { return $false }
        if ($bytes[0] -eq 169 -and $bytes[1] -eq 254) { return $false }
        if ($bytes[0] -eq 172 -and $bytes[1] -ge 16 -and $bytes[1] -le 31) { return $false }
        if ($bytes[0] -eq 192 -and $bytes[1] -eq 168) { return $false }
        if ($bytes[0] -ge 224) { return $false }
        return $true
    }

    $bytes = $ip.GetAddressBytes()
    if ($ip.IsIPv6LinkLocal -or $ip.IsIPv6Multicast -or $ip.IsIPv6SiteLocal) { return $false }
    if (($bytes[0] -band 0xFE) -eq 0xFC) { return $false }
    return $true
}

function Test-IsSystemSecurityAccount {
    param(
        [AllowNull()][string]$Account,
        [AllowNull()][string]$Domain
    )

    if ([string]::IsNullOrWhiteSpace($Account)) { return $true }
    if ($Account.EndsWith("$")) { return $true }

    $accountUpper = $Account.ToUpperInvariant()
    $identity = "$Domain\$Account".ToUpperInvariant()
    return $identity -in @(
        "NT AUTHORITY\SYSTEM",
        "NT AUTHORITY\LOCAL SERVICE",
        "NT AUTHORITY\NETWORK SERVICE"
    ) -or $accountUpper -in @("SYSTEM", "LOCAL SERVICE", "NETWORK SERVICE")
}

function Get-FailedLogonFindings {
    param(
        [array]$Events,
        [int]$WarnThreshold = 5,
        [int]$CriticalThreshold = 20
    )

    $groups = @()
    $groups += @($Events | Where-Object {
        $_.Id -eq 4625 -and [System.Net.IPAddress]::TryParse([string]$_.IpAddress, [ref]([System.Net.IPAddress]$null))
    } | Group-Object IpAddress | ForEach-Object {
        [pscustomobject]@{ Kind = "Bron-IP"; Value = $_.Name; Count = $_.Count }
    })
    $groups += @($Events | Where-Object {
        $_.Id -eq 4625 -and -not [string]::IsNullOrWhiteSpace([string]$_.Account) -and $_.Account -ne "-"
    } | Group-Object Account | ForEach-Object {
        [pscustomobject]@{ Kind = "Account"; Value = $_.Name; Count = $_.Count }
    })

    @($groups | Where-Object { $_.Count -ge $WarnThreshold } | ForEach-Object {
        [pscustomobject]@{
            Kind   = $_.Kind
            Value  = $_.Value
            Count  = $_.Count
            Status = if ($_.Count -ge $CriticalThreshold) { "KRIT" } else { "WARN" }
        }
    } | Sort-Object @{ Expression = "Count"; Descending = $true }, Kind, Value)
}

function Get-LogonSecurityAnalysis {
    param(
        [array]$Events,
        [int]$WarnThreshold = 5,
        [int]$CriticalThreshold = 20
    )

    $sessionsByLogonId = @{}
    foreach ($session in @($Events | Where-Object { $_.Id -eq 4624 -and -not [string]::IsNullOrWhiteSpace([string]$_.LogonId) })) {
        if (-not $sessionsByLogonId.ContainsKey([string]$session.LogonId)) {
            $sessionsByLogonId[[string]$session.LogonId] = $session
        }
    }

    $privilegedLogons = @($Events | Where-Object {
        $_.Id -eq 4672 -and -not (Test-IsSystemSecurityAccount -Account $_.Account -Domain $_.Domain)
    } | ForEach-Object {
        $privileged = $_
        $session = $sessionsByLogonId[[string]$privileged.LogonId]
        [pscustomobject]@{
            Id          = $privileged.Id
            TimeCreated = $privileged.TimeCreated
            Account     = $privileged.Account
            Domain      = $privileged.Domain
            LogonType   = if ($session) { $session.LogonType } else { $privileged.LogonType }
            IpAddress   = if ($session) { $session.IpAddress } else { $privileged.IpAddress }
            Workstation = if ($session) { $session.Workstation } else { $privileged.Workstation }
            LogonId     = $privileged.LogonId
            Elevated    = if ($session) { $session.Elevated } else { $privileged.Elevated }
            Privileges  = $privileged.Privileges
            Status      = $privileged.Status
            SubStatus   = $privileged.SubStatus
        }
    })

    [pscustomobject]@{
        SuccessfulInteractive = @($Events | Where-Object { $_.Id -eq 4624 -and $_.LogonType -in 2, 10 })
        Failed                 = @($Events | Where-Object { $_.Id -eq 4625 })
        Findings               = @(Get-FailedLogonFindings -Events $Events -WarnThreshold $WarnThreshold -CriticalThreshold $CriticalThreshold)
        PublicRdpLogons        = @($Events | Where-Object { $_.Id -eq 4624 -and $_.LogonType -eq 10 -and (Test-IsPublicIPAddress $_.IpAddress) })
        PrivilegedLogons       = $privilegedLogons
    }
}

function Get-RecentSecurityEvents {
    param(
        [datetime]$StartTime,
        [int]$MaxEvents = 2000,
        [scriptblock]$EventReader
    )

    $filter = @{
        LogName   = "Security"
        Id        = 4624, 4625, 4672
        StartTime = $StartTime
    }

    try {
        $rawEvents = if ($EventReader) {
            & $EventReader $filter
        } else {
            Get-WinEvent -FilterHashtable $filter -ErrorAction Stop
        }
        [pscustomobject]@{
            Events = @($rawEvents | Select-Object -First $MaxEvents)
            Error  = $null
        }
    } catch {
        [pscustomobject]@{
            Events = @()
            Error  = $_.Exception.Message
        }
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
        $tijd    = Format-EventTime $_.TimeCreated
        $bericht = Format-EventMessage $_.Message -SingleLine -MaxLength 80
        Write-Host "      $lvl $tijd  $($_.ProviderName): $bericht" -ForegroundColor $kleur
    }
}

# ------------------------------------------------------------
#  AANMELDINGSBEVEILIGING
# ------------------------------------------------------------

Write-Sectie "AANMELDINGSBEVEILIGING (laatste $AanmeldLogUren uur)"

$securityResult = Get-RecentSecurityEvents `
    -StartTime (Get-Date).AddHours(-$AanmeldLogUren) `
    -MaxEvents $MaxAanmeldEvents
$securityEvents = @()
$aanmeldAnalyse = $null
$securityFout = $securityResult.Error

if ($securityFout) {
    Write-Rij "Security-log" "Niet leesbaar: $securityFout" "WARN"
} elseif ($securityResult.Events.Count -eq 0) {
    Write-Rij "Aanmeldingen" "Geen events gevonden; controleer het auditbeleid" "WARN"
} else {
    $securityEvents = @($securityResult.Events | ForEach-Object {
        ConvertFrom-SecurityEvent $_
    } | Where-Object { $null -ne $_ })

    if ($securityEvents.Count -eq 0) {
        Write-Rij "Aanmeldingen" "Events konden niet worden verwerkt; controleer het auditbeleid" "WARN"
    } else {
        $aanmeldAnalyse = Get-LogonSecurityAnalysis `
            -Events $securityEvents `
            -WarnThreshold $DrempelLoginWarn `
            -CriticalThreshold $DrempelLoginKrit

        $rdpAantal = @($aanmeldAnalyse.SuccessfulInteractive | Where-Object LogonType -eq 10).Count
        Write-Rij "Interactief/RDP" "$($aanmeldAnalyse.SuccessfulInteractive.Count) geslaagd ($rdpAantal RDP)" "INFO"
        Write-Rij "Mislukt" "$($aanmeldAnalyse.Failed.Count) poging(en)" "INFO"

        foreach ($finding in $aanmeldAnalyse.Findings) {
            Write-Rij "Verdacht $($finding.Kind)" "$($finding.Value): $($finding.Count) poging(en)" $finding.Status
        }

        if ($aanmeldAnalyse.PublicRdpLogons.Count -gt 0) {
            Write-Rij "Publieke RDP" "$($aanmeldAnalyse.PublicRdpLogons.Count) geslaagde aanmelding(en)" "WARN"
        } else {
            Write-Rij "Publieke RDP" "Geen geslaagde RDP-aanmeldingen vanaf publieke IP-adressen" "OK"
        }

        Write-Rij "Privileged" "$($aanmeldAnalyse.PrivilegedLogons.Count) niet-systeem event(s)" "INFO"

        Write-Host ""
        Write-Host "    Laatste 10 relevante aanmeldingen:" -ForegroundColor DarkGray
        $securityEvents | Where-Object { $_.Id -in 4624, 4625 -and ($_.Id -eq 4625 -or $_.LogonType -in 2, 10) } |
            Sort-Object TimeCreated -Descending | Select-Object -First 10 | ForEach-Object {
                $tijd = Format-EventTime $_.TimeCreated
                $soort = if ($_.Id -eq 4625) { "MISLUKT" } elseif ($_.LogonType -eq 10) { "RDP" } else { "INTERACTIEF" }
                $kleur = if ($_.Id -eq 4625) { "Yellow" } elseif ($_.LogonType -eq 10 -and (Test-IsPublicIPAddress $_.IpAddress)) { "Yellow" } else { "Gray" }
                $account = if ([string]::IsNullOrWhiteSpace($_.Account)) { "-" } else { "$($_.Domain)\$($_.Account)".TrimStart("\") }
                $adres = if ([string]::IsNullOrWhiteSpace($_.IpAddress) -or $_.IpAddress -eq "-") { "lokaal/onbekend" } else { $_.IpAddress }
                Write-Host "      [$soort] $tijd  $account  vanaf $adres" -ForegroundColor $kleur
            }
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
        "<tr><td style='color:$kleur;font-weight:bold'>[$tag]</td><td>$(ConvertTo-HtmlText $_.Label)</td><td>$(ConvertTo-HtmlText $_.Waarde)</td></tr>"
    }) -join "`n"

    # Aanmeldingsbeveiliging
    if ($securityFout) {
        $securitySectie = "<h3>Aanmeldingsbeveiliging</h3><p style='color:#f39c12'>Security-log niet leesbaar: $(ConvertTo-HtmlText $securityFout)</p>"
    } elseif ($null -eq $aanmeldAnalyse) {
        $securitySectie = "<h3>Aanmeldingsbeveiliging</h3><p style='color:#f39c12'>Geen bruikbare aanmeldingsevents gevonden. Controleer het auditbeleid.</p>"
    } else {
        $findingRijen = @($aanmeldAnalyse.Findings | ForEach-Object {
            $kleur = Get-StatusKleur $_.Status
            "<tr><td style='color:$kleur;font-weight:bold'>$(ConvertTo-HtmlText $_.Status)</td><td>$(ConvertTo-HtmlText $_.Kind)</td><td>$(ConvertTo-HtmlText $_.Value)</td><td>$(ConvertTo-HtmlText $_.Count)</td></tr>"
        }) -join "`n"
        if ([string]::IsNullOrWhiteSpace($findingRijen)) {
            $findingRijen = "<tr><td style='color:#2ecc71;font-weight:bold'>OK</td><td colspan='3'>Geen groepen boven de waarschuwingsdrempel.</td></tr>"
        }

        $aanmeldRijen = @($aanmeldAnalyse.SuccessfulInteractive | Sort-Object TimeCreated -Descending | Select-Object -First 50 | ForEach-Object {
            $isPublicRdp = $_.LogonType -eq 10 -and (Test-IsPublicIPAddress $_.IpAddress)
            $status = if ($isPublicRdp) { "WARN" } else { "INFO" }
            $kleur = Get-StatusKleur $status
            $type = if ($_.LogonType -eq 10) { "RDP" } else { "Interactief" }
            "<tr><td style='color:$kleur;font-weight:bold'>$(ConvertTo-HtmlText $status)</td><td>$(ConvertTo-HtmlText (Format-EventTime $_.TimeCreated -Format 'dd/MM/yyyy HH:mm:ss'))</td><td>$(ConvertTo-HtmlText $type)</td><td>$(ConvertTo-HtmlText $_.Domain)\$(ConvertTo-HtmlText $_.Account)</td><td>$(ConvertTo-HtmlText $_.IpAddress)</td><td>$(ConvertTo-HtmlText $_.Workstation)</td></tr>"
        }) -join "`n"
        if ([string]::IsNullOrWhiteSpace($aanmeldRijen)) {
            $aanmeldRijen = "<tr><td colspan='6'>Geen interactieve of RDP-aanmeldingen gevonden.</td></tr>"
        }

        $privilegedRijen = @($aanmeldAnalyse.PrivilegedLogons | Sort-Object TimeCreated -Descending | Select-Object -First 50 | ForEach-Object {
            "<tr><td>$(ConvertTo-HtmlText (Format-EventTime $_.TimeCreated -Format 'dd/MM/yyyy HH:mm:ss'))</td><td>$(ConvertTo-HtmlText $_.Domain)\$(ConvertTo-HtmlText $_.Account)</td><td>$(ConvertTo-HtmlText $_.LogonId)</td><td>$(ConvertTo-HtmlText $_.Privileges)</td></tr>"
        }) -join "`n"
        if ([string]::IsNullOrWhiteSpace($privilegedRijen)) {
            $privilegedRijen = "<tr><td colspan='4'>Geen niet-systeem privileged logons gevonden.</td></tr>"
        }

        $securitySectie = @"
<h3 style='color:#00d4ff;margin-top:40px'>Aanmeldingsbeveiliging - Laatste $AanmeldLogUren uur</h3>
<p>
Geslaagd interactief/RDP: <b>$($aanmeldAnalyse.SuccessfulInteractive.Count)</b> &nbsp;
Mislukt: <b>$($aanmeldAnalyse.Failed.Count)</b> &nbsp;
Publieke RDP: <b>$($aanmeldAnalyse.PublicRdpLogons.Count)</b> &nbsp;
Privileged: <b>$($aanmeldAnalyse.PrivilegedLogons.Count)</b>
</p>

<h4>Verdachte mislukte aanmeldingen</h4>
<table>
<tr><th>Status</th><th>Groepering</th><th>Waarde</th><th>Aantal</th></tr>
$findingRijen
</table>

<h4>Recente interactieve en RDP-aanmeldingen</h4>
<table>
<tr><th>Status</th><th>Tijdstip</th><th>Type</th><th>Account</th><th>Bron-IP</th><th>Workstation</th></tr>
$aanmeldRijen
</table>

<h4>Niet-systeem privileged logons</h4>
<table>
<tr><th>Tijdstip</th><th>Account</th><th>Logon-ID</th><th>Privileges</th></tr>
$privilegedRijen
</table>
"@
    }

    # Event log tabel rijen
    if ($events -and $events.Count -gt 0) {
        $eventRijen = ($events | ForEach-Object {
            $lvlTekst  = if ($_.Level -eq 1) { "Kritiek" } else { "Fout" }
            $lvlKleur  = if ($_.Level -eq 1) { "#e74c3c" } else { "#f39c12" }
            $tijd      = Format-EventTime $_.TimeCreated -Format "dd/MM/yyyy HH:mm:ss"
            $bericht   = Format-EventMessage $_.Message -Html
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

$securitySectie

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
