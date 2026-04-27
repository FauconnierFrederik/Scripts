# ============================================================
#  Launcher.ps1
#  Centraal startpunt voor beheerscripts.
#  Meerdere scripts tegelijk selecteerbaar via nummers.
#  Uitvoeren als Administrator aanbevolen.
# ============================================================

#Requires -RunAsAdministrator

# ════════════════════════════════════════════════════════════
#  BASISPAD — scripts worden gezocht in dezelfde map als de launcher
#  Wil je scripts in een andere map? Pas $ScriptBase aan.
# ════════════════════════════════════════════════════════════
$ScriptBase = $PSScriptRoot

# ════════════════════════════════════════════════════════════
#  SCRIPTS TOEVOEGEN — pas alleen dit blok aan
#
#  Formaat per entry:
#  @{
#      Name        = "Korte weergavenaam"
#      Description = "Wat doet dit script?"
#      Path        = "$ScriptBase\NaamVanScript.ps1"
#  }
# ════════════════════════════════════════════════════════════
$Scripts = @(

    @{
        Name        = "Install – Check PSCM Service"
        Description = "Installeert Watch-CagService.ps1 en registreert de Scheduled Task."
        Path        = "$ScriptBase\Install-CheckPSCMService.ps1"
    }

    @{
        Name        = "Rapport – MFA Type overzicht"
        Description = "Genereert een CSV-rapport van MFA-registratie voor alle M365-gebruikers via MSOnline."
        Path        = "$ScriptBase\MFATypeReport.ps1"
    }

    @{
        Name        = "Setup – RDP certificaat en bestand"
        Description = "Maakt een self-signed certificaat aan, configureert RDP-Tcp en genereert een gesigned RDP-bestand."
        Path        = "$ScriptBase\Setup-RDPV4.ps1"
    }

    @{
        Name        = "Taken – Reboot scheduled tasks aanmaken"
        Description = "Maakt een wekelijkse en een eenmalige reboot-taak aan onder het SYSTEM-account."
        Path        = "$ScriptBase\Create-RebootTasks.ps1"
    }

    @{
        Name        = "Exchange – Online Archief activeren"
        Description = "Activeert het Exchange Online archief voor een mailbox en koppelt een retention policy van 3 jaar."
        Path        = "$ScriptBase\Enable-Archive.ps1"
    }

    @{
        Name        = "Setup – Domain Controller (nieuw forest)"
        Description = "Installeert AD DS en promoveert de server tot eerste DC in een nieuw domein. Herstart automatisch."
        Path        = "$ScriptBase\Setup-DomainController.ps1"
    }

    @{
        Name        = "WinUtil – Chris Titus Tech"
        Description = "Start WinUtil: software installeren, Windows tweaks, optimalisaties en reparaties. (vereist internet)"
        Path        = "$ScriptBase\Launch-WinUtil.ps1"
    }

    # ── Voeg hier nieuwe scripts toe ──────────────────────────

    # @{
    #     Name        = "Mijn Nieuw Script"
    #     Description = "Korte uitleg over wat dit script doet."
    #     Path        = "$ScriptBase\MijnNieuwScript.ps1"
    # }

)

# ════════════════════════════════════════════════════════════
#  Hieronder niets aanpassen — dit is de engine van de launcher
# ════════════════════════════════════════════════════════════

# ── Kleuren & opmaak ─────────────────────────────────────────
$C = @{
    Title   = "Cyan"
    OK      = "Green"
    Warn    = "Yellow"
    Error   = "Red"
    Dim     = "DarkGray"
    Default = "White"
}

function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  ██╗   ██╗██╗   ██╗██╗ ██████╗ ███╗   ██╗" -ForegroundColor $C.Title
    Write-Host "  ██║   ██║██║   ██║██║██╔═══██╗████╗  ██║" -ForegroundColor $C.Title
    Write-Host "  ██║   ██║██║   ██║██║██║   ██║██╔██╗ ██║" -ForegroundColor $C.Title
    Write-Host "  ██║   ██║╚██╗ ██╔╝██║██║   ██║██║╚██╗██║" -ForegroundColor $C.Title
    Write-Host "  ╚██████╔╝ ╚████╔╝ ██║╚██████╔╝██║ ╚████║" -ForegroundColor $C.Title
    Write-Host "   ╚═════╝   ╚═══╝  ╚═╝ ╚═════╝ ╚═╝  ╚═══╝" -ForegroundColor $C.Title
    Write-Host ""
    Write-Host "  Script Launcher" -ForegroundColor $C.Dim
    Write-Host "  $('─' * 50)" -ForegroundColor $C.Dim
    Write-Host ""
}

function Show-Menu {
    param([array]$ScriptList)

    Write-Host "  Beschikbare scripts:" -ForegroundColor $C.Default
    Write-Host ""

    for ($i = 0; $i -lt $ScriptList.Count; $i++) {
        $num  = "[$($i + 1)]"
        $name = $ScriptList[$i].Name
        $desc = $ScriptList[$i].Description
        $path = $ScriptList[$i].Path

        $exists = Test-Path $path
        $status = if ($exists) { "[OK]" } else { "[!!]" }
        $statusColor = if ($exists) { $C.OK } else { $C.Error }

        Write-Host "  " -NoNewline
        Write-Host $num -ForegroundColor $C.Title -NoNewline
        Write-Host "  $status " -ForegroundColor $statusColor -NoNewline
        Write-Host $name -ForegroundColor $C.Default
        Write-Host "       $desc" -ForegroundColor $C.Dim
        Write-Host ""
    }

    Write-Host "  $('─' * 50)" -ForegroundColor $C.Dim
    Write-Host "  [A] Alle scripts uitvoeren" -ForegroundColor $C.Warn
    Write-Host "  [Q] Afsluiten" -ForegroundColor $C.Dim
    Write-Host ""
}

function Invoke-Script {
    param([hashtable]$Script)

    $border = "─" * 50
    Write-Host ""
    Write-Host "  $border" -ForegroundColor $C.Dim
    Write-Host "  >> Uitvoeren: $($Script.Name)" -ForegroundColor $C.Title
    Write-Host "  $border" -ForegroundColor $C.Dim

    if (-not (Test-Path $Script.Path)) {
        Write-Host "  [!!] Script niet gevonden: $($Script.Path)" -ForegroundColor $C.Error
        Write-Host ""
        return
    }

    try {
        & $Script.Path
        Write-Host ""
        Write-Host "  [OK] '$($Script.Name)' succesvol afgerond." -ForegroundColor $C.OK
    }
    catch {
        Write-Host ""
        Write-Host "  [!!] Fout bij uitvoeren van '$($Script.Name)':" -ForegroundColor $C.Error
        Write-Host "    $_" -ForegroundColor $C.Error
    }

    Write-Host ""
}

function Read-Selection {
    param([int]$Max)

    Write-Host "  Geef nummers in (bijv. 1  of  1,3  of  2-4) of A/Q:" -ForegroundColor $C.Default
    Write-Host "  > " -ForegroundColor $C.Title -NoNewline
    $raw = Read-Host

    $raw = $raw.Trim().ToUpper()

    if ($raw -eq "Q") { return "QUIT" }
    if ($raw -eq "A") { return 1..$Max }

    # Verwerk komma's, spaties en bereiken (bijv. 2-4)
    $indices = @()
    $parts   = $raw -split "[,\s]+" | Where-Object { $_ -ne "" }

    foreach ($part in $parts) {
        if ($part -match "^(\d+)-(\d+)$") {
            $from = [int]$Matches[1]
            $to   = [int]$Matches[2]
            if ($from -le $to -and $from -ge 1 -and $to -le $Max) {
                $indices += $from..$to
            }
            else {
                Write-Host "  [!] Ongeldig bereik: $part (max $Max)" -ForegroundColor $C.Warn
            }
        }
        elseif ($part -match "^\d+$") {
            $num = [int]$part
            if ($num -ge 1 -and $num -le $Max) {
                $indices += $num
            }
            else {
                Write-Host "  [!] Nummer buiten bereik: $part" -ForegroundColor $C.Warn
            }
        }
        else {
            Write-Host "  [!] Ongeldige invoer genegeerd: '$part'" -ForegroundColor $C.Warn
        }
    }

    return ($indices | Select-Object -Unique | Sort-Object)
}

# ── Hoofdlus ─────────────────────────────────────────────────
do {
    Show-Header
    Show-Menu -ScriptList $Scripts

    $selection = Read-Selection -Max $Scripts.Count

    if ($selection -eq "QUIT") {
        Write-Host ""
        Write-Host "  Tot ziens!" -ForegroundColor $C.Dim
        Write-Host ""
        break
    }

    if (-not $selection -or $selection.Count -eq 0) {
        Write-Host "  Geen geldige keuze. Probeer opnieuw." -ForegroundColor $C.Warn
        Start-Sleep -Seconds 2
        continue
    }

    foreach ($index in $selection) {
        Invoke-Script -Script $Scripts[$index - 1]
    }

    Write-Host ""
    Write-Host "  $('─' * 50)" -ForegroundColor $C.Dim
    Write-Host "  Druk op een toets om terug te keren naar het menu..." -ForegroundColor $C.Dim
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

} while ($true)