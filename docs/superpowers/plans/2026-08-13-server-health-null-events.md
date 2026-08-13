# Server Health Null Events Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Voorkom dat eventrecords zonder tijdstip of bericht de consoleweergave of HTML-export van `Get-ServerHealth.ps1` afbreken.

**Architecture:** Twee kleine formatteringsfuncties centraliseren null-afhandeling voor eventtijd en eventbericht. De bestaande console- en HTML-paden roepen deze functies aan; een Pester-test importeert alleen deze functies via de PowerShell-AST zodat de interactieve healthcheck niet wordt uitgevoerd.

**Tech Stack:** Windows PowerShell 5.1, Pester 3.4, PowerShell AST-parser, Git

---

### Task 1: Regressietest voor ontbrekende eventvelden

**Files:**
- Create locally (ignored by repository policy): `tests/Get-ServerHealth.Tests.ps1`
- Test: `tests/Get-ServerHealth.Tests.ps1`

- [x] **Step 1: Schrijf de falende test**

Maak een Pester-test die de definities `Format-EventTime` en `Format-EventMessage` uit `Get-ServerHealth.ps1` importeert via `[System.Management.Automation.Language.Parser]::ParseFile()`. Test exact deze resultaten:

```powershell
Format-EventTime $null                                  # Onbekend
Format-EventMessage $null                              # (Geen bericht beschikbaar)
Format-EventMessage ('x' * 81) -SingleLine -MaxLength 80 # 80 tekens
Format-EventMessage "A&B`r`n<C>" -Html                 # A&amp;B<br>&lt;C&gt;
```

Als een formatter nog niet bestaat, geeft de test `__MISSING__` terug zodat de assertions aantoonbaar falen wegens ontbrekende productiecode.

- [x] **Step 2: Controleer de rode test**

Run:

```powershell
Invoke-Pester .\tests\Get-ServerHealth.Tests.ps1 -PassThru
```

Expected: vier falende tests met verwachte veilige waarden en werkelijke waarde `__MISSING__`.

### Task 2: Null-safe formatteringsfuncties implementeren

**Files:**
- Modify: `Get-ServerHealth.ps1:66`
- Modify: `Get-ServerHealth.ps1:205-206`
- Modify: `Get-ServerHealth.ps1:257-258`
- Test: `tests/Get-ServerHealth.Tests.ps1`

- [x] **Step 1: Voeg de minimale formatteringsfuncties toe**

```powershell
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
```

- [x] **Step 2: Gebruik de functies in beide uitvoerpaden**

Console:

```powershell
$tijd    = Format-EventTime $_.TimeCreated
$bericht = Format-EventMessage $_.Message -SingleLine -MaxLength 80
Write-Host "      $lvl $tijd  $($_.ProviderName): $bericht" -ForegroundColor $kleur
```

HTML:

```powershell
$tijd    = Format-EventTime $_.TimeCreated -Format "dd/MM/yyyy HH:mm:ss"
$bericht = Format-EventMessage $_.Message -Html
```

- [x] **Step 3: Controleer de groene test**

Run:

```powershell
Invoke-Pester .\tests\Get-ServerHealth.Tests.ps1 -PassThru
```

Expected: vier tests geslaagd, nul mislukt.

### Task 3: Volledige verificatie en publicatie

**Files:**
- Modify: `docs/superpowers/plans/2026-08-13-server-health-null-events.md`

- [x] **Step 1: Parseer productie- en testscript**

Run `[System.Management.Automation.Language.Parser]::ParseFile()` voor beide `.ps1`-bestanden en faal als `ParseError.Count` groter dan nul is.

- [x] **Step 2: Controleer diff en repository-status**

Run:

```powershell
git diff --check
git status --short
git diff -- Get-ServerHealth.ps1 tests/Get-ServerHealth.Tests.ps1
```

Expected: alleen de geplande getrackte bestanden en de bestaande niet-getrackte `.memdb/`; geen whitespacefouten. De lokale regressietest verschijnt niet omdat `Tests/` volgens `.gitignore` bewust niet wordt gepubliceerd.

- [ ] **Step 3: Commit de implementatie**

```powershell
git add Get-ServerHealth.ps1 docs/superpowers/plans/2026-08-13-server-health-null-events.md
git commit -m "Fix server health handling of null event fields"
```

- [ ] **Step 4: Push en controleer GitHub**

```powershell
git push origin main
```

Haal daarna `Get-ServerHealth.ps1` via de GitHub-koppeling van `main` op en bevestig dat beide formatteraanroepen aanwezig zijn op de gepushte commit.
