#Requires -RunAsAdministrator
# ============================================================
#  Clear-DiskSpace.ps1
#  Ruimt veilig schijfruimte vrij op Windows Server.
#
#  Interactief:  & .\Clear-DiskSpace.ps1
#  Gepland:      & .\Clear-DiskSpace.ps1 -Scheduled
#                (stille uitvoering met veilige standaardtaken)
# ============================================================

param(
    [switch]$Scheduled
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ------------------------------------------------------------
#  HULPFUNCTIES
# ------------------------------------------------------------

function Write-Step    { param([string]$Msg) Write-Host "`n[*] $Msg" -ForegroundColor Cyan }
function Write-Success { param([string]$Msg) Write-Host "    [OK] $Msg" -ForegroundColor Green }
function Write-Warn    { param([string]$Msg) Write-Host "    [!]  $Msg" -ForegroundColor Yellow }
function Write-Err     { param([string]$Msg) Write-Host "    [!!] $Msg" -ForegroundColor Red }

function Get-FolderSizeMB {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return 0 }
    $bytes = (Get-ChildItem $Path -Recurse -Force -ErrorAction SilentlyContinue |
              Measure-Object -Property Length -Sum).Sum
    return [math]::Round($bytes / 1MB, 1)
}

function Remove-SafeFolder {
    param([string]$Path, [int]$OuderDanDagen = 0)
    if (-not (Test-Path $Path)) { return 0 }

    $voor = Get-FolderSizeMB $Path

    if ($OuderDanDagen -gt 0) {
        $grens = (Get-Date).AddDays(-$OuderDanDagen)
        Get-ChildItem $Path -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { -not $_.PSIsContainer -and $_.LastWriteTime -lt $grens } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    } else {
        Get-ChildItem $Path -Force -ErrorAction SilentlyContinue |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }

    $na = Get-FolderSizeMB $Path
    return [math]::Round($voor - $na, 1)
}

function Get-DriveInfo {
    param([string]$Drive = "C")
    $disk = Get-PSDrive -Name $Drive -ErrorAction SilentlyContinue
    if (-not $disk) { return $null }
    return @{
        Totaal   = [math]::Round(($disk.Used + $disk.Free) / 1GB, 1)
        Vrij     = [math]::Round($disk.Free / 1GB, 1)
        Gebruikt = [math]::Round($disk.Used / 1GB, 1)
    }
}

# ------------------------------------------------------------
#  CHECKBOX FORMULIER (interactieve modus)
# ------------------------------------------------------------

function Show-TaakKeuze {
    param([hashtable]$Schijf)

    $form                 = New-Object System.Windows.Forms.Form
    $form.Text            = "Schijf opruimen - Kies taken"
    $form.Size            = New-Object System.Drawing.Size(500, 540)
    $form.StartPosition   = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox     = $false

    $lblInfo           = New-Object System.Windows.Forms.Label
    $lblInfo.Text      = "Schijf C:  $($Schijf.Gebruikt) GB gebruikt  /  $($Schijf.Vrij) GB vrij  /  $($Schijf.Totaal) GB totaal"
    $lblInfo.Location  = New-Object System.Drawing.Point(12, 12)
    $lblInfo.Size      = New-Object System.Drawing.Size(468, 20)
    $lblInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $lblKies           = New-Object System.Windows.Forms.Label
    $lblKies.Text      = "Selecteer welke opruimingstaken uitgevoerd moeten worden:"
    $lblKies.Location  = New-Object System.Drawing.Point(12, 40)
    $lblKies.Size      = New-Object System.Drawing.Size(468, 20)

    $taken = @(
        @{ Key = "Temp";    Label = "Windows Temp bestanden  (C:\Windows\Temp)";            Checked = $true  }
        @{ Key = "WUCache"; Label = "Windows Update cache  (SoftwareDistribution\Download)"; Checked = $true  }
        @{ Key = "IIS";     Label = "IIS Logs ouder dan 30 dagen  (C:\inetpub\logs)";       Checked = $false }
        @{ Key = "CBS";     Label = "CBS & DISM Logs  (C:\Windows\Logs)";                   Checked = $true  }
        @{ Key = "WER";     Label = "Windows Error Reporting bestanden";                     Checked = $true  }
        @{ Key = "Dump";    Label = "MiniDump bestanden  (C:\Windows\MiniDump)";             Checked = $true  }
        @{ Key = "Bin";     Label = "Prullenbak leegmaken";                                  Checked = $false }
        @{ Key = "DISM";    Label = "WinSxS opruimen via DISM  (duurt 5-15 minuten)";       Checked = $false }
    )

    $checkboxes = @{}
    $y = 68
    foreach ($taak in $taken) {
        $cb          = New-Object System.Windows.Forms.CheckBox
        $cb.Text     = $taak.Label
        $cb.Location = New-Object System.Drawing.Point(16, $y)
        $cb.Size     = New-Object System.Drawing.Size(460, 22)
        $cb.Checked  = $taak.Checked
        $checkboxes[$taak.Key] = $cb
        $form.Controls.Add($cb)
        $y += 28
    }

    $yWarn   = $y + 10
    $yBtn    = $y + 40
    $yClient = $y + 84

    $lblWarn           = New-Object System.Windows.Forms.Label
    $lblWarn.Text      = "Alle acties zijn veilig en omkeerbaar via Windows Update / herinstallatie."
    $lblWarn.Location  = New-Object System.Drawing.Point(12, $yWarn)
    $lblWarn.Size      = New-Object System.Drawing.Size(468, 20)
    $lblWarn.ForeColor = [System.Drawing.Color]::Gray

    $btnOK              = New-Object System.Windows.Forms.Button
    $btnOK.Text         = "Uitvoeren"
    $btnOK.Location     = New-Object System.Drawing.Point(290, $yBtn)
    $btnOK.Size         = New-Object System.Drawing.Size(90, 30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel              = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Annuleren"
    $btnCancel.Location     = New-Object System.Drawing.Point(392, $yBtn)
    $btnCancel.Size         = New-Object System.Drawing.Size(90, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $form.Controls.AddRange(@($lblInfo, $lblKies, $lblWarn, $btnOK, $btnCancel))
    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel
    $form.ClientSize   = New-Object System.Drawing.Size(500, $yClient)

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $geselecteerd = @{}
        foreach ($key in $checkboxes.Keys) { $geselecteerd[$key] = $checkboxes[$key].Checked }
        return $geselecteerd
    }
    return $null
}

# ------------------------------------------------------------
#  SCHEDULED TASK AANMAKEN
# ------------------------------------------------------------

function Register-WeeklyCleanup {
    param([string]$ScriptPad)

    $dag = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Op welke dag moet de wekelijkse opruiming uitgevoerd worden?`n`nKies uit: Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday",
        "Wekelijkse opruiming - Dag",
        "Sunday"
    )
    if ([string]::IsNullOrWhiteSpace($dag)) { return }

    $tijd = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Op welk tijdstip? (HH:mm, 24-uurs formaat)`n`nVoorbeelden: 02:00  |  22:30",
        "Wekelijkse opruiming - Tijdstip",
        "02:00"
    )
    if ([string]::IsNullOrWhiteSpace($tijd)) { return }

    $taaknaam = "Wekelijkse Schijfopruiming"

    if (Get-ScheduledTask -TaskName $taaknaam -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $taaknaam -Confirm:$false
    }

    $action   = New-ScheduledTaskAction `
        -Execute  "powershell.exe" `
        -Argument "-NonInteractive -NoProfile -ExecutionPolicy Bypass -File `"$ScriptPad`" -Scheduled"

    $trigger  = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $dag -At $tijd

    $settings = New-ScheduledTaskSettingsSet `
        -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
        -StartWhenAvailable `
        -MultipleInstances IgnoreNew

    $principal = New-ScheduledTaskPrincipal `
        -UserId    "NT AUTHORITY\SYSTEM" `
        -LogonType ServiceAccount `
        -RunLevel  Highest

    Register-ScheduledTask `
        -TaskName    $taaknaam `
        -Action      $action `
        -Trigger     $trigger `
        -Settings    $settings `
        -Principal   $principal `
        -Description "Voert wekelijks een veilige schijfopruiming uit via Clear-DiskSpace.ps1" `
        -Force | Out-Null

    [System.Windows.Forms.MessageBox]::Show(
        "Scheduled Task aangemaakt!`n`nTaaknaam  : $taaknaam`nUitvoering: elke $dag om $tijd`nAccount   : SYSTEM`n`nBeheer via taskschd.msc",
        "Wekelijkse opruiming - Ingepland", "OK", "Information"
    )
}

# ------------------------------------------------------------
#  OPRUIMINGSTAKEN UITVOEREN
# ------------------------------------------------------------

function Invoke-Cleanup {
    param([hashtable]$Keuzes)

    $totaalVrij = 0

    if ($Keuzes["Temp"]) {
        Write-Step "Windows Temp bestanden verwijderen..."
        $v  = Remove-SafeFolder "C:\Windows\Temp"
        $v += Remove-SafeFolder $env:TEMP
        $totaalVrij += $v
        Write-Success "$v MB vrijgemaakt"
    }

    if ($Keuzes["WUCache"]) {
        Write-Step "Windows Update cache verwijderen..."
        try {
            Stop-Service -Name wuauserv, bits -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            $v = Remove-SafeFolder "C:\Windows\SoftwareDistribution\Download"
            $totaalVrij += $v
            Write-Success "$v MB vrijgemaakt"
        } finally {
            Start-Service -Name wuauserv, bits -ErrorAction SilentlyContinue
        }
    }

    if ($Keuzes["IIS"]) {
        Write-Step "IIS Logs ouder dan 30 dagen verwijderen..."
        $iisLogPad = "C:\inetpub\logs\LogFiles"
        if (Test-Path $iisLogPad) {
            $v = Remove-SafeFolder $iisLogPad -OuderDanDagen 30
            $totaalVrij += $v
            Write-Success "$v MB vrijgemaakt"
        } else {
            Write-Warn "IIS logmap niet gevonden."
        }
    }

    if ($Keuzes["CBS"]) {
        Write-Step "CBS en DISM logs verwijderen..."
        $v  = Remove-SafeFolder "C:\Windows\Logs\CBS"
        $v += Remove-SafeFolder "C:\Windows\Logs\DISM"
        $totaalVrij += $v
        Write-Success "$v MB vrijgemaakt"
    }

    if ($Keuzes["WER"]) {
        Write-Step "Windows Error Reporting bestanden verwijderen..."
        $werPad = "C:\ProgramData\Microsoft\Windows\WER"
        $v  = Remove-SafeFolder "$werPad\ReportQueue"
        $v += Remove-SafeFolder "$werPad\ReportArchive"
        $v += Remove-SafeFolder "$werPad\Temp"
        $totaalVrij += $v
        Write-Success "$v MB vrijgemaakt"
    }

    if ($Keuzes["Dump"]) {
        Write-Step "MiniDump bestanden verwijderen..."
        $v  = Remove-SafeFolder "C:\Windows\MiniDump"
        $v += Remove-SafeFolder "C:\Windows\minidump"
        $totaalVrij += $v
        Write-Success "$v MB vrijgemaakt"
    }

    if ($Keuzes["Bin"]) {
        Write-Step "Prullenbak leegmaken..."
        try {
            Clear-RecycleBin -Force -ErrorAction Stop
            Write-Success "Prullenbak geleegd"
        } catch {
            Write-Warn "Prullenbak al leeg of geen toegang."
        }
    }

    if ($Keuzes["DISM"]) {
        Write-Step "WinSxS opruimen via DISM (kan 5-15 minuten duren)..."
        Write-Warn "Even geduld..."
        $output = & dism.exe /Online /Cleanup-Image /StartComponentCleanup 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Success "WinSxS opruiming voltooid"
        } else {
            Write-Warn "DISM melding: $($output | Select-Object -Last 2 | Out-String)"
        }
    }

    return $totaalVrij
}

# ------------------------------------------------------------
#  HOOFDLOGICA
# ------------------------------------------------------------

Write-Host ""
Write-Host "  Schijf Opruimen" -ForegroundColor Cyan
Write-Host "  $('-' * 50)" -ForegroundColor DarkGray

$schijfVoor = Get-DriveInfo "C"

if ($Scheduled) {
    # Stille modus voor Scheduled Task - vaste veilige taken
    Write-Host "  Modus: Gepland (geen GUI)" -ForegroundColor DarkGray
    $keuzes = @{
        Temp    = $true
        WUCache = $true
        IIS     = $false
        CBS     = $true
        WER     = $true
        Dump    = $true
        Bin     = $false
        DISM    = $false
    }
} else {
    # Interactieve modus
    if (-not $schijfVoor) {
        [System.Windows.Forms.MessageBox]::Show("Schijf C: niet gevonden.", "Fout", "OK", "Error")
        exit
    }
    $keuzes = Show-TaakKeuze -Schijf $schijfVoor
    if (-not $keuzes) {
        Write-Host "`n  Geannuleerd." -ForegroundColor Yellow
        exit
    }
}

$totaalVrij = Invoke-Cleanup -Keuzes $keuzes
$schijfNa   = Get-DriveInfo "C"
$verschil   = [math]::Round($schijfNa.Vrij - $schijfVoor.Vrij, 2)

Write-Host ""
Write-Host "  $('-' * 50)" -ForegroundColor DarkGray
Write-Host "  Vrijgemaakt : $totaalVrij MB" -ForegroundColor Green
Write-Host "  Vrij voor   : $($schijfVoor.Vrij) GB" -ForegroundColor White
Write-Host "  Vrij na     : $($schijfNa.Vrij) GB" -ForegroundColor Green
Write-Host "  $('-' * 50)" -ForegroundColor DarkGray

if (-not $Scheduled) {
    Add-Type -AssemblyName Microsoft.VisualBasic

    $samenvatting  = "Opruiming voltooid!`n`n"
    $samenvatting += "Vrijgemaakt  : $totaalVrij MB`n"
    $samenvatting += "Vrij voor    : $($schijfVoor.Vrij) GB`n"
    $samenvatting += "Vrij na      : $($schijfNa.Vrij) GB`n"
    $samenvatting += "Verschil     : +$verschil GB`n`n"
    $samenvatting += "Wil je dit script wekelijks automatisch inplannen?"

    $plannen = [System.Windows.Forms.MessageBox]::Show(
        $samenvatting, "Schijf opruimen - Klaar", "YesNo", "Information"
    )

    if ($plannen -eq "Yes") {
        Register-WeeklyCleanup -ScriptPad $PSCommandPath
    }
}
