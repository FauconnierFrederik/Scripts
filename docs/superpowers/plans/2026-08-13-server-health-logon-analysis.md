# Server Health Logon Analysis Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Voeg een taal-onafhankelijke analyse van verdachte Windows-aanmeldingen toe aan de console- en HTML-uitvoer van `Get-ServerHealth.ps1`.

**Architecture:** Security-events 4624, 4625 en 4672 worden één keer ingelezen en via benoemde XML-velden genormaliseerd. Kleine pure functies classificeren IP-adressen, filteren systeemaccounts en berekenen bevindingen; console en HTML consumeren hetzelfde analyseobject.

**Tech Stack:** Windows PowerShell 5.1, `Get-WinEvent`, PowerShell XML/AST, Pester 3.4, Git

---

### Task 1: Falende tests voor parsing en detectieregels

**Files:**
- Modify locally (ignored by repository policy): `tests/Get-ServerHealth.Tests.ps1`
- Test: `tests/Get-ServerHealth.Tests.ps1`

- [ ] **Step 1: Breid de AST-import en XML-fixtures uit**

Voeg deze productiefuncties toe aan `$targetNames`:

```powershell
"ConvertTo-HtmlText",
"ConvertFrom-SecurityEvent",
"Test-IsPublicIPAddress",
"Test-IsSystemSecurityAccount",
"Get-FailedLogonFindings",
"Get-LogonSecurityAnalysis",
"Get-RecentSecurityEvents"
```

Voeg een fixturefunctie toe die een WinEvent-achtig object met `ToXml()` teruggeeft:

```powershell
function New-TestSecurityEvent {
    param([int]$Id, [string]$Xml, [datetime]$TimeCreated = (Get-Date))
    $event = [pscustomobject]@{ Id = $Id; Xml = $Xml; TimeCreated = $TimeCreated }
    $event | Add-Member ScriptMethod ToXml { return $this.Xml }
    return $event
}
```

- [ ] **Step 2: Voeg gerichte rode tests toe**

Schrijf Pester-tests die controleren dat:

```powershell
# 4624 XML wordt Account=admin, LogonType=10, IpAddress=203.0.113.10.
# Ontbrekende velden leveren lege waarden en geen uitzondering.
(Test-IsPublicIPAddress "8.8.8.8") | Should Be $true
(Test-IsPublicIPAddress "10.0.0.1") | Should Be $false
(Test-IsPublicIPAddress "100.64.0.1") | Should Be $false
(Test-IsPublicIPAddress "fc00::1") | Should Be $false
(Test-IsPublicIPAddress "2001:4860:4860::8888") | Should Be $true
(Test-IsSystemSecurityAccount -Account "SYSTEM" -Domain "NT AUTHORITY") | Should Be $true
(Test-IsSystemSecurityAccount -Account "SERVER01$" -Domain "DOMAIN") | Should Be $true
(Test-IsSystemSecurityAccount -Account "admin" -Domain "DOMAIN") | Should Be $false
# Groepen met 4/5/19/20 events krijgen respectievelijk geen finding/WARN/WARN/KRIT.
# Een type-10 4624 vanaf een publiek IP verschijnt in PublicRdpLogons.
# Een reader die 'Access denied' gooit levert Error en nul Events.
(ConvertTo-HtmlText "<admin>&") | Should Be "&lt;admin&gt;&amp;"
```

Gebruik bij ontbrekende functies dezelfde `__MISSING__`-techniek als de bestaande tests, zodat de rode tests falen door ontbrekende productiecode.

- [ ] **Step 3: Controleer de rode toestand**

Run:

```powershell
$result = Invoke-Pester .\tests\Get-ServerHealth.Tests.ps1 -PassThru
```

Expected: de bestaande vier tests slagen; de nieuwe tests mislukken omdat de zeven functies nog ontbreken.

### Task 2: Pure helpers en Security-eventreader implementeren

**Files:**
- Modify: `Get-ServerHealth.ps1:14-18`
- Modify: `Get-ServerHealth.ps1:31-106`
- Test: `tests/Get-ServerHealth.Tests.ps1`

- [ ] **Step 1: Voeg configuratie toe**

```powershell
$AanmeldLogUren       = 24
$DrempelLoginWarn     = 5
$DrempelLoginKrit     = 20
$MaxAanmeldEvents     = 2000
```

- [ ] **Step 2: Voeg HTML-, XML- en classificatiehelpers toe**

Implementeer:

```powershell
function ConvertTo-HtmlText {
    param([AllowNull()][object]$Value)
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function ConvertFrom-SecurityEvent {
    param([Parameter(Mandatory)][object]$Event)
    try { [xml]$xml = $Event.ToXml() } catch { return $null }
    $data = @{}
    foreach ($node in @($xml.Event.EventData.Data)) {
        $name = [string]$node.Name
        if ($name) { $data[$name] = [string]$node.InnerText }
    }
    $isPrivilegeEvent = [int]$Event.Id -eq 4672
    $logonType = 0
    [void][int]::TryParse($data["LogonType"], [ref]$logonType)
    return [pscustomobject]@{
        Id          = [int]$Event.Id
        TimeCreated = $Event.TimeCreated
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
        $b = $ip.GetAddressBytes()
        if ($b[0] -in 0, 10, 127) { return $false }
        if ($b[0] -eq 100 -and $b[1] -ge 64 -and $b[1] -le 127) { return $false }
        if ($b[0] -eq 169 -and $b[1] -eq 254) { return $false }
        if ($b[0] -eq 172 -and $b[1] -ge 16 -and $b[1] -le 31) { return $false }
        if ($b[0] -eq 192 -and $b[1] -eq 168) { return $false }
        if ($b[0] -ge 224) { return $false }
        return $true
    }
    $bytes = $ip.GetAddressBytes()
    if ($ip.IsIPv6LinkLocal -or $ip.IsIPv6Multicast -or $ip.IsIPv6SiteLocal) { return $false }
    if (($bytes[0] -band 0xFE) -eq 0xFC) { return $false }
    return $true
}

function Test-IsSystemSecurityAccount {
    param([AllowNull()][string]$Account, [AllowNull()][string]$Domain)
    if ([string]::IsNullOrWhiteSpace($Account)) { return $true }
    if ($Account.EndsWith("$")) { return $true }
    $identity = "$Domain\$Account".ToUpperInvariant()
    return $identity -in @(
        "NT AUTHORITY\SYSTEM",
        "NT AUTHORITY\LOCAL SERVICE",
        "NT AUTHORITY\NETWORK SERVICE"
    ) -or $Account.ToUpperInvariant() -in @("SYSTEM", "LOCAL SERVICE", "NETWORK SERVICE")
}
```

- [ ] **Step 3: Voeg groepering, analyse en reader toe**

`Get-FailedLogonFindings` groepeert event 4625 afzonderlijk per valide IP en per niet-leeg account. Alleen groepen vanaf `$WarnThreshold` worden teruggegeven met `Kind`, `Value`, `Count` en `Status`; `$CriticalThreshold` bepaalt `KRIT`.

`Get-LogonSecurityAnalysis` retourneert exact:

```powershell
[pscustomobject]@{
    SuccessfulInteractive = @($Events | Where-Object { $_.Id -eq 4624 -and $_.LogonType -in 2, 10 })
    Failed                 = @($Events | Where-Object Id -eq 4625)
    Findings               = @(Get-FailedLogonFindings ...)
    PublicRdpLogons        = @($Events | Where-Object { $_.Id -eq 4624 -and $_.LogonType -eq 10 -and (Test-IsPublicIPAddress $_.IpAddress) })
    PrivilegedLogons       = @($Events | Where-Object { $_.Id -eq 4672 -and -not (Test-IsSystemSecurityAccount $_.Account $_.Domain) })
}
```

`Get-RecentSecurityEvents` accepteert `StartTime`, `MaxEvents` en een optionele `EventReader`. De standaardreader gebruikt `Get-WinEvent -FilterHashtable @{ LogName='Security'; Id=4624,4625,4672; StartTime=$StartTime } -ErrorAction Stop`. De functie retourneert altijd `{ Events; Error }` en vangt toegangsfouten af.

- [ ] **Step 4: Controleer helpertests**

Run `Invoke-Pester .\tests\Get-ServerHealth.Tests.ps1 -PassThru`.

Expected: alle parsing-, IP-, account-, drempel-, analyse-, HTML- en reader-tests slagen.

### Task 3: Consoleanalyse integreren

**Files:**
- Modify: `Get-ServerHealth.ps1` direct na de bestaande algemene EVENT LOG-sectie

- [ ] **Step 1: Lees en normaliseer Security-events**

```powershell
Write-Sectie "AANMELDINGSBEVEILIGING (laatste $AanmeldLogUren uur)"
$securityResult = Get-RecentSecurityEvents -StartTime (Get-Date).AddHours(-$AanmeldLogUren) -MaxEvents $MaxAanmeldEvents
$aanmeldAnalyse = $null
$securityFout = $securityResult.Error
if ($securityFout) {
    Write-Rij "Security-log" "Niet leesbaar: $securityFout" "WARN"
} elseif ($securityResult.Events.Count -eq 0) {
    Write-Rij "Aanmeldingen" "Geen events gevonden; controleer het auditbeleid" "WARN"
} else {
    $securityEvents = @($securityResult.Events | ForEach-Object { ConvertFrom-SecurityEvent $_ } | Where-Object { $null -ne $_ })
    $aanmeldAnalyse = Get-LogonSecurityAnalysis -Events $securityEvents -WarnThreshold $DrempelLoginWarn -CriticalThreshold $DrempelLoginKrit
}
```

- [ ] **Step 2: Toon samenvatting en details**

Bij een analyse schrijft `Write-Rij` INFO-totalen voor geslaagde interactieve/RDP-logons en mislukkingen. Iedere finding wordt `Write-Rij "Verdacht <Kind>" "<Value>: <Count> poging(en)" <Status>`. Publieke RDP-logons worden samengevat met `WARN`. Privileged logons worden met `INFO` getoond. Schrijf vervolgens maximaal tien recente relevante events met `Write-Host`, veilige tijdopmaak en lege waarden als `-`.

- [ ] **Step 3: Parseer en voer Pester opnieuw uit**

Expected: nul PowerShell-parsefouten en alle Pester-tests groen.

### Task 4: HTML-beveiligingssectie integreren

**Files:**
- Modify: `Get-ServerHealth.ps1` in het bestaande HTML-exportblok

- [ ] **Step 1: Bouw veilige HTML-rijen**

Maak `$securitySectie` vóór de hoofd-here-string. Bij `$securityFout` of ontbrekende analyse bevat ze een HTML-gecodeerde waarschuwing. Anders bouwt ze tabellen voor `Findings`, `SuccessfulInteractive`, `PublicRdpLogons` en `PrivilegedLogons`; alle account-, domein-, IP-, workstation- en statuswaarden gaan door `ConvertTo-HtmlText`.

- [ ] **Step 2: Voeg de sectie aan het rapport toe**

Plaats `$securitySectie` vóór `$eventSectie` in de bestaande HTML-body, zodat beveiligingssignalen boven de algemene eventlogdetails staan.

- [ ] **Step 3: Test HTML-codering en parsebaarheid**

Run Pester en de PowerShell-parser. Expected: alle tests groen en nul parsefouten.

### Task 5: Verifiëren, committen en publiceren

**Files:**
- Modify: `Get-ServerHealth.ps1`
- Add: `docs/superpowers/plans/2026-08-13-server-health-logon-analysis.md`
- Local-only test: `tests/Get-ServerHealth.Tests.ps1`

- [ ] **Step 1: Voer volledige verificatie uit**

Run Pester, parseer productie- en testscript, voer `git diff --check` uit en inspecteer `git diff -- Get-ServerHealth.ps1`.

- [ ] **Step 2: Commit de getrackte bestanden**

```powershell
git add Get-ServerHealth.ps1 docs/superpowers/plans/2026-08-13-server-health-logon-analysis.md
git commit -m "Add suspicious logon analysis to server health"
```

- [ ] **Step 3: Integreer eventuele nieuwe remote commits zonder force-push**

Run `git fetch origin main`; als `origin/main` vooruitgelopen is, inspecteer de wijzigingen en rebase alleen bij niet-overlappende of veilig oplosbare wijzigingen. Voer daarna de volledige verificatie opnieuw uit.

- [ ] **Step 4: Push en controleer GitHub main**

Run `git push origin main`. Haal daarna `Get-ServerHealth.ps1` via de GitHub-koppeling op en bevestig dat de Security-eventquery, detectiedrempels, publieke-RDP-detectie en `$securitySectie` aanwezig zijn.
