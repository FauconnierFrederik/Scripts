#Requires -RunAsAdministrator
# ============================================================
#  Expand-MainPartition.ps1
#  Twee modi:
#    1. C: uitbreiden + recovery herstellen  (geen reboot)
#    2. Alleen recovery partitie herstellen  (na mislukte uitbreiding)
#
#  Vereisten: GPT schijf, unallocated space aanwezig.
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ------------------------------------------------------------
#  HULPFUNCTIES
# ------------------------------------------------------------

function Write-Step    { param([string]$Msg) Write-Host "`n[*] $Msg" -ForegroundColor Cyan }
function Write-Success { param([string]$Msg) Write-Host "    [OK] $Msg" -ForegroundColor Green }
function Write-Warn    { param([string]$Msg) Write-Host "    [!]  $Msg" -ForegroundColor Yellow }
function Write-Err     { param([string]$Msg) Write-Host "    [!!] $Msg" -ForegroundColor Red }

function Invoke-Diskpart {
    param([string[]]$Regels)
    $tmp = [System.IO.Path]::Combine($env:TEMP, "dp_$(Get-Random).txt")
    $Regels | Out-File -FilePath $tmp -Encoding ASCII -Force
    $output = & diskpart.exe /s $tmp 2>&1
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    return $output
}

# ------------------------------------------------------------
#  GEMEENSCHAPPELIJKE PRE-FLIGHT
# ------------------------------------------------------------

$recoveryGuid = '{de94bba4-06d1-4d40-a16a-bfd50179d6ac}'
$winrePad     = "C:\Windows\System32\Recovery\winre.wim"

try {
    $cPart = Get-Partition -DriveLetter C -ErrorAction Stop
} catch {
    [System.Windows.Forms.MessageBox]::Show("Partitie C: niet gevonden.", "Fout", "OK", "Error")
    exit
}

$diskNum    = $cPart.DiskNumber
$disk       = Get-Disk -Number $diskNum -ErrorAction Stop
$cGrootteGB = [math]::Round($cPart.Size / 1GB, 1)

if ($disk.PartitionStyle -ne 'GPT') {
    [System.Windows.Forms.MessageBox]::Show(
        "Schijf $diskNum is MBR. Dit script ondersteunt enkel GPT-schijven.",
        "Niet ondersteund", "OK", "Error"
    )
    exit
}

# ------------------------------------------------------------
#  MODUS SELECTIE
# ------------------------------------------------------------

$rbUitbreiden = New-Object System.Windows.Forms.RadioButton
$rbHerstel    = New-Object System.Windows.Forms.RadioButton

$modusForm                 = New-Object System.Windows.Forms.Form
$modusForm.Text            = "Schijf $diskNum — Kies actie"
$modusForm.StartPosition   = "CenterScreen"
$modusForm.FormBorderStyle = "FixedDialog"
$modusForm.MaximizeBox     = $false

$y = 14

$lblModusTitel          = New-Object System.Windows.Forms.Label
$lblModusTitel.Text     = "Schijf $diskNum  ($($disk.FriendlyName))  —  C: $cGrootteGB GB"
$lblModusTitel.Location = New-Object System.Drawing.Point(14, $y)
$lblModusTitel.Size     = New-Object System.Drawing.Size(470, 20)
$lblModusTitel.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$modusForm.Controls.Add($lblModusTitel)
$y += 32

$rbUitbreiden.Text      = "C: uitbreiden + recovery herstellen"
$rbUitbreiden.Location  = New-Object System.Drawing.Point(14, $y)
$rbUitbreiden.Size      = New-Object System.Drawing.Size(470, 20)
$rbUitbreiden.Checked   = $true
$rbUitbreiden.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$modusForm.Controls.Add($rbUitbreiden)
$y += 22

$lblUitbreidenDesc           = New-Object System.Windows.Forms.Label
$lblUitbreidenDesc.Text      = "    Verwijdert recovery, breidt C: uit met unallocated space, plaatst recovery terug."
$lblUitbreidenDesc.Location  = New-Object System.Drawing.Point(14, $y)
$lblUitbreidenDesc.Size      = New-Object System.Drawing.Size(470, 18)
$lblUitbreidenDesc.ForeColor = [System.Drawing.Color]::DimGray
$modusForm.Controls.Add($lblUitbreidenDesc)
$y += 28

$rbHerstel.Text     = "Alleen recovery partitie herstellen"
$rbHerstel.Location = New-Object System.Drawing.Point(14, $y)
$rbHerstel.Size     = New-Object System.Drawing.Size(470, 20)
$rbHerstel.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$modusForm.Controls.Add($rbHerstel)
$y += 22

$lblHerstelDesc           = New-Object System.Windows.Forms.Label
$lblHerstelDesc.Text      = "    Maakt recovery partitie opnieuw aan vanuit winre.wim. Gebruik na een mislukte uitbreiding."
$lblHerstelDesc.Location  = New-Object System.Drawing.Point(14, $y)
$lblHerstelDesc.Size      = New-Object System.Drawing.Size(470, 18)
$lblHerstelDesc.ForeColor = [System.Drawing.Color]::DimGray
$modusForm.Controls.Add($lblHerstelDesc)
$y += 32

$btnModusOK              = New-Object System.Windows.Forms.Button
$btnModusOK.Text         = "Volgende"
$btnModusOK.Location     = New-Object System.Drawing.Point(298, $y)
$btnModusOK.Size         = New-Object System.Drawing.Size(90, 30)
$btnModusOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

$btnModusCancel              = New-Object System.Windows.Forms.Button
$btnModusCancel.Text         = "Annuleren"
$btnModusCancel.Location     = New-Object System.Drawing.Point(400, $y)
$btnModusCancel.Size         = New-Object System.Drawing.Size(90, 30)
$btnModusCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

$modusForm.Controls.AddRange(@($btnModusOK, $btnModusCancel))
$modusForm.AcceptButton = $btnModusOK
$modusForm.CancelButton = $btnModusCancel
[int]$modusFormH      = $y + 50
$modusForm.ClientSize = New-Object System.Drawing.Size(510, $modusFormH)

if ($modusForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "`n  Geannuleerd." -ForegroundColor Yellow
    exit
}

$modus = if ($rbUitbreiden.Checked) { "Uitbreiden" } else { "Herstel" }

# ============================================================
#  MODUS 1 — C: UITBREIDEN + RECOVERY HERSTELLEN
# ============================================================

if ($modus -eq "Uitbreiden") {

    # Pre-flight
    $recoveryPart = Get-Partition -DiskNumber $diskNum -ErrorAction SilentlyContinue |
        Where-Object { $_.GptType -eq $recoveryGuid } |
        Select-Object -Last 1

    if (-not $recoveryPart) {
        [System.Windows.Forms.MessageBox]::Show(
            "Geen Windows Recovery partitie gevonden op schijf $diskNum.`n`n" +
            "Als de partitie al verwijderd is, gebruik dan 'Alleen recovery herstellen'.`n`n" +
            "Vereist GPT-type: $recoveryGuid",
            "Niet gevonden", "OK", "Warning"
        )
        exit
    }

    $vrijSpace = $disk.LargestFreeExtent
    if ($vrijSpace -lt 100MB) {
        [System.Windows.Forms.MessageBox]::Show(
            "Geen bruikbare unallocated space op schijf $diskNum.`n`n" +
            "Vergroot eerst de virtuele schijf of verklein een andere partitie.",
            "Geen ruimte", "OK", "Warning"
        )
        exit
    }

    $cNaGB      = [math]::Round(($cPart.Size + $vrijSpace) / 1GB, 1)
    $recoveryMB = [math]::Round($recoveryPart.Size / 1MB, 0)
    $vrijMB     = [math]::Round($vrijSpace / 1MB, 0)

    # Bevestigingsscherm
    $form                 = New-Object System.Windows.Forms.Form
    $form.Text            = "Hoofdpartitie uitbreiden"
    $form.StartPosition   = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox     = $false

    $y = 14

    $lblTitel          = New-Object System.Windows.Forms.Label
    $lblTitel.Text     = "Schijfindeling — Schijf $diskNum  ($($disk.FriendlyName))"
    $lblTitel.Location = New-Object System.Drawing.Point(14, $y)
    $lblTitel.Size     = New-Object System.Drawing.Size(470, 20)
    $lblTitel.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblTitel)
    $y += 30

    foreach ($tekst in @(
        "  C: (huidig)           $cGrootteGB GB   →   na uitbreiding: ~$cNaGB GB"
        "  Recovery partitie     $recoveryMB MB       (partitie $($recoveryPart.PartitionNumber))"
        "  Unallocated space     $vrijMB MB"
    )) {
        $lbl          = New-Object System.Windows.Forms.Label
        $lbl.Text     = $tekst
        $lbl.Location = New-Object System.Drawing.Point(14, $y)
        $lbl.Size     = New-Object System.Drawing.Size(470, 18)
        $lbl.Font     = New-Object System.Drawing.Font("Consolas", 9)
        $form.Controls.Add($lbl)
        $y += 20
    }

    $y += 14

    $lblWatLabel          = New-Object System.Windows.Forms.Label
    $lblWatLabel.Text     = "Wat er gebeurt:"
    $lblWatLabel.Location = New-Object System.Drawing.Point(14, $y)
    $lblWatLabel.Size     = New-Object System.Drawing.Size(470, 18)
    $lblWatLabel.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblWatLabel)
    $y += 22

    foreach ($stap in @(
        "  1. WinRE uitschakelen  (winre.wim → C:\Windows\System32\Recovery)"
        "  2. Recovery partitie verwijderen"
        "  3. C: uitbreiden (${recoveryMB} MB gereserveerd voor recovery)"
        "  4. Nieuwe recovery partitie aanmaken aan het einde"
        "  5. WinRE herstellen en heractiveren"
    )) {
        $lbl           = New-Object System.Windows.Forms.Label
        $lbl.Text      = $stap
        $lbl.Location  = New-Object System.Drawing.Point(14, $y)
        $lbl.Size      = New-Object System.Drawing.Size(470, 18)
        $lbl.ForeColor = [System.Drawing.Color]::DimGray
        $form.Controls.Add($lbl)
        $y += 18
    }

    $y += 16

    $lblRisico           = New-Object System.Windows.Forms.Label
    $lblRisico.Text      = "Maak een systeemback-up voor je verdergaat. Geen reboot nodig."
    $lblRisico.Location  = New-Object System.Drawing.Point(14, $y)
    $lblRisico.Size      = New-Object System.Drawing.Size(470, 18)
    $lblRisico.ForeColor = [System.Drawing.Color]::DarkRed
    $form.Controls.Add($lblRisico)
    $y += 36

    $btnOK              = New-Object System.Windows.Forms.Button
    $btnOK.Text         = "Uitvoeren"
    $btnOK.Location     = New-Object System.Drawing.Point(298, $y)
    $btnOK.Size         = New-Object System.Drawing.Size(90, 30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel              = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Annuleren"
    $btnCancel.Location     = New-Object System.Drawing.Point(400, $y)
    $btnCancel.Size         = New-Object System.Drawing.Size(90, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $form.Controls.AddRange(@($btnOK, $btnCancel))
    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel
    [int]$formH        = $y + 50
    $form.ClientSize   = New-Object System.Drawing.Size(510, $formH)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "`n  Geannuleerd." -ForegroundColor Yellow
        exit
    }

    # Uitvoering
    $foutStap   = $null
    $nieuwePart = $null
    $tempLetter = $null

    Write-Host ""
    Write-Host "  Hoofdpartitie uitbreiden" -ForegroundColor Cyan
    Write-Host "  $('-' * 50)" -ForegroundColor DarkGray

    # Stap 1 — WinRE uitschakelen
    Write-Step "WinRE uitschakelen..."
    & reagentc.exe /disable 2>&1 | Out-Null

    if (-not (Test-Path $winrePad)) {
        Write-Err "winre.wim niet gevonden op $winrePad na uitschakelen."
        Write-Err "Geen wijzigingen doorgevoerd. Controleer: reagentc /info"
        [System.Windows.Forms.MessageBox]::Show(
            "WinRE uitschakelen mislukt.`nwinre.wim niet gevonden op:`n$winrePad`n`nGeen wijzigingen gemaakt.",
            "Fout", "OK", "Error"
        )
        exit
    }
    Write-Success "WinRE uitgeschakeld — winre.wim veilig op $winrePad"

    # Stap 2 — Bevestiging verwijderen
    $partOffset = [math]::Round($recoveryPart.Offset / 1GB, 2)
    $bevestig = [System.Windows.Forms.MessageBox]::Show(
        "Bevestig verwijderen van de volgende partitie:`n`n" +
        "  Schijf       : $diskNum  ($($disk.FriendlyName))`n" +
        "  Partitie nr  : $($recoveryPart.PartitionNumber)`n" +
        "  Grootte      : $recoveryMB MB`n" +
        "  Offset       : $partOffset GB`n" +
        "  GPT-type     : $recoveryGuid`n`n" +
        "WinRE is uitgeschakeld en winre.wim staat veilig op:`n  $winrePad`n`n" +
        "Na verwijderen wordt C: uitgebreid en de partitie herschapen.",
        "Partitie verwijderen bevestigen",
        "YesNo",
        "Warning"
    )

    if ($bevestig -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Warn "Geannuleerd door gebruiker — WinRE heractiveren..."
        & reagentc.exe /enable 2>&1 | Out-Null
        Write-Warn "WinRE heractiveerd. Geen wijzigingen aan schijfindeling."
        exit
    }

    # Stap 3 — Recovery partitie verwijderen
    Write-Step "Recovery partitie verwijderen (partitie $($recoveryPart.PartitionNumber))..."
    Invoke-Diskpart @(
        "select disk $diskNum"
        "select partition $($recoveryPart.PartitionNumber)"
        "delete partition override"
    ) | Out-Null

    $controle = Get-Partition -DiskNumber $diskNum -ErrorAction SilentlyContinue |
        Where-Object { $_.GptType -eq $recoveryGuid }

    if ($controle) {
        Write-Err "Recovery partitie nog aanwezig na verwijderen — afgebroken."
        Write-Warn "Heractiveer WinRE handmatig: reagentc /enable"
        [System.Windows.Forms.MessageBox]::Show(
            "Recovery partitie kon niet verwijderd worden.`n`n" +
            "C: is NIET uitgebreid.`n`n" +
            "Voer handmatig uit: reagentc /enable",
            "Mislukt", "OK", "Error"
        )
        exit
    }
    Write-Success "Recovery partitie verwijderd"

    # Stap 4 — C: uitbreiden (recovery ruimte reserveren)
    Write-Step "C: uitbreiden (${recoveryMB} MB gereserveerd voor recovery)..."
    try {
        $maxSize    = (Get-PartitionSupportedSize -DriveLetter C -ErrorAction Stop).SizeMax
        $reserveer  = [int64]$recoveryMB * 1MB
        $targetSize = $maxSize - $reserveer
        if ($targetSize -le $cPart.Size) {
            Write-Err "Berekende doelgrootte ($([math]::Round($targetSize/1GB,1)) GB) is niet groter dan huidig ($cGrootteGB GB) — afgebroken."
            $foutStap = "C: uitbreiden"
        } else {
            Resize-Partition -DriveLetter C -Size $targetSize -ErrorAction Stop
            $nieuweC = [math]::Round((Get-Partition -DriveLetter C).Size / 1GB, 1)
            Write-Success "C: uitgebreid naar $nieuweC GB ($recoveryMB MB vrij voor recovery)"
        }
    } catch {
        Write-Err "Uitbreiden mislukt: $_"
        $foutStap = "C: uitbreiden"
    }

    # Stap 5 — Nieuwe recovery partitie aanmaken
    if (-not $foutStap) {
        Write-Step "Nieuwe recovery partitie aanmaken ($recoveryMB MB)..."

        $gebruikteBrieven = (Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue).Name
        $tempLetter = 'Z','Y','X','W','V','U','T','S','R' |
            Where-Object { $_ -notin $gebruikteBrieven } | Select-Object -First 1

        if (-not $tempLetter) {
            Write-Err "Geen vrije driveletter beschikbaar voor tijdelijke toewijzing."
            $foutStap = "recovery partitie aanmaken"
        } else {
            Invoke-Diskpart @(
                "select disk $diskNum"
                "create partition primary size=$recoveryMB"
                "format quick fs=ntfs label=`"WinRE`""
                "set id=de94bba4-06d1-4d40-a16a-bfd50179d6ac"
                "gpt attributes=0x8000000000000001"
                "assign letter=$tempLetter"
            ) | Out-Null

            Start-Sleep -Seconds 2

            $nieuwePart = Get-Partition -DiskNumber $diskNum -ErrorAction SilentlyContinue |
                Where-Object { $_.GptType -eq $recoveryGuid } | Select-Object -Last 1

            if (-not $nieuwePart) {
                Write-Err "Nieuwe recovery partitie niet gevonden na aanmaken."
                $foutStap = "recovery partitie aanmaken"
            } else {
                Write-Success "Recovery partitie aangemaakt (partitie $($nieuwePart.PartitionNumber))"
            }
        }
    }

    # Stap 6 — WinRE terugzetten
    if (-not $foutStap) {
        Write-Step "WinRE bestanden terugzetten naar ${tempLetter}:..."
        try {
            $doelPad = "${tempLetter}:\Recovery\WindowsRE"
            New-Item -Path $doelPad -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Copy-Item -Path $winrePad -Destination "$doelPad\winre.wim" -Force -ErrorAction Stop
            Write-Success "winre.wim gekopieerd naar $doelPad"
        } catch {
            Write-Err "Kopiëren winre.wim mislukt: $_"
            $foutStap = "winre.wim kopiëren"
        }
    }

    # Stap 7 — WinRE heractiveren
    if (-not $foutStap) {
        Write-Step "WinRE heractiveren..."
        $enableOut = & reagentc.exe /enable 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Success "WinRE actief"
        } else {
            Write-Warn "reagentc /enable: $enableOut"
        }
    }

    # Tijdelijke driveletter verwijderen
    if ($nieuwePart -and $tempLetter) {
        Write-Step "Tijdelijke driveletter $tempLetter verwijderen..."
        Invoke-Diskpart @(
            "select disk $diskNum"
            "select partition $($nieuwePart.PartitionNumber)"
            "remove letter=$tempLetter"
        ) | Out-Null
        Write-Success "Driveletter verwijderd"
    }

    # Resultaat
    Write-Host ""
    Write-Host "  $('-' * 50)" -ForegroundColor DarkGray

    $eindC          = [math]::Round((Get-Partition -DriveLetter C -ErrorAction SilentlyContinue).Size / 1GB, 1)
    $reagentInfoRaw = & reagentc.exe /info 2>&1
    $reagentStatus  = ($reagentInfoRaw | Select-String "Windows RE status") -replace '\s+', ' '
    if (-not $reagentStatus) { $reagentStatus = "(status onbekend)" }

    if ($foutStap) {
        Write-Host "  Gedeeltelijk voltooid — '$foutStap' mislukt." -ForegroundColor Yellow
        Write-Host "  winre.wim bewaard op: $winrePad" -ForegroundColor Yellow
        [System.Windows.Forms.MessageBox]::Show(
            "Operatie gedeeltelijk voltooid.`n`n" +
            "Stap mislukt  : $foutStap`n" +
            "winre.wim op  : $winrePad`n`n" +
            "Voer handmatig uit:`n  reagentc /enable",
            "Gedeeltelijk voltooid", "OK", "Warning"
        )
    } else {
        Write-Host "  Klaar!" -ForegroundColor Green
        Write-Host "  C: voor   : $cGrootteGB GB" -ForegroundColor White
        Write-Host "  C: na     : $eindC GB" -ForegroundColor Green
        Write-Host "  $('-' * 50)" -ForegroundColor DarkGray
        [System.Windows.Forms.MessageBox]::Show(
            "Uitbreiding geslaagd!`n`n" +
            "C: voor      : $cGrootteGB GB`n" +
            "C: na        : $eindC GB`n" +
            "Recovery     : hersteld ($recoveryMB MB)`n" +
            "$reagentStatus`n`n" +
            "Geen reboot vereist.",
            "Hoofdpartitie uitgebreid", "OK", "Information"
        )
    }
}

# ============================================================
#  MODUS 2 — ALLEEN RECOVERY PARTITIE HERSTELLEN
# ============================================================

if ($modus -eq "Herstel") {

    Write-Host ""
    Write-Host "  Recovery partitie herstellen" -ForegroundColor Cyan
    Write-Host "  $('-' * 50)" -ForegroundColor DarkGray

    # winre.wim vereist
    if (-not (Test-Path $winrePad)) {
        Write-Err "winre.wim niet gevonden op $winrePad"
        [System.Windows.Forms.MessageBox]::Show(
            "winre.wim niet gevonden op:`n$winrePad`n`n" +
            "Kan recovery partitie niet herstellen zonder dit bestand.`n`n" +
            "Mogelijk staat winre.wim al op een bestaande recovery partitie.",
            "winre.wim ontbreekt", "OK", "Error"
        )
        exit
    }

    $wimGrootteMB = [math]::Round((Get-Item $winrePad).Length / 1MB, 0)

    # Check bestaande recovery partitie
    $bestaandeRecovery = Get-Partition -DiskNumber $diskNum -ErrorAction SilentlyContinue |
        Where-Object { $_.GptType -eq $recoveryGuid } | Select-Object -Last 1

    if ($bestaandeRecovery) {
        $antw = [System.Windows.Forms.MessageBox]::Show(
            "Er bestaat al een recovery partitie op schijf $diskNum:`n`n" +
            "  Partitie nr  : $($bestaandeRecovery.PartitionNumber)`n" +
            "  Grootte      : $([math]::Round($bestaandeRecovery.Size/1MB,0)) MB`n`n" +
            "Wil je alleen WinRE heractiveren (reagentc /enable) zonder nieuwe partitie aan te maken?",
            "Recovery partitie gevonden", "YesNo", "Question"
        )
        if ($antw -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Step "WinRE heractiveren..."
            $enableOut = & reagentc.exe /enable 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Success "WinRE actief"
                [System.Windows.Forms.MessageBox]::Show(
                    "WinRE heractiveerd.`n`nRecovery partitie was al aanwezig (partitie $($bestaandeRecovery.PartitionNumber)).",
                    "Klaar", "OK", "Information"
                )
            } else {
                Write-Warn "reagentc /enable: $enableOut"
                [System.Windows.Forms.MessageBox]::Show(
                    "reagentc /enable gaf een waarschuwing:`n$enableOut",
                    "Waarschuwing", "OK", "Warning"
                )
            }
        }
        exit
    }

    # Bepaal recovery grootte: winre.wim + 150 MB buffer, afgerond naar boven per 50 MB, min 500 MB
    $herstelMB = [math]::Max(500, [math]::Ceiling(($wimGrootteMB + 150) / 50) * 50)

    # Unallocated space check
    $disk      = Get-Disk -Number $diskNum -ErrorAction Stop
    $vrijSpace = $disk.LargestFreeExtent
    if ($vrijSpace -lt ($herstelMB * 1MB)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Onvoldoende unallocated space op schijf $diskNum.`n`n" +
            "Nodig      : $herstelMB MB`n" +
            "Beschikbaar: $([math]::Round($vrijSpace/1MB,0)) MB`n`n" +
            "Verklein C: handmatig om ruimte vrij te maken.",
            "Geen ruimte", "OK", "Warning"
        )
        exit
    }

    $vrijHerstelMB = [math]::Round($vrijSpace / 1MB, 0)

    # Bevestigingsscherm
    $herstelForm                 = New-Object System.Windows.Forms.Form
    $herstelForm.Text            = "Recovery partitie herstellen"
    $herstelForm.StartPosition   = "CenterScreen"
    $herstelForm.FormBorderStyle = "FixedDialog"
    $herstelForm.MaximizeBox     = $false

    $y = 14

    $lblHTitel          = New-Object System.Windows.Forms.Label
    $lblHTitel.Text     = "Schijfindeling — Schijf $diskNum  ($($disk.FriendlyName))"
    $lblHTitel.Location = New-Object System.Drawing.Point(14, $y)
    $lblHTitel.Size     = New-Object System.Drawing.Size(470, 20)
    $lblHTitel.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $herstelForm.Controls.Add($lblHTitel)
    $y += 30

    foreach ($tekst in @(
        "  C: (huidig)           $cGrootteGB GB"
        "  Unallocated space     $vrijHerstelMB MB  (recovery krijgt $herstelMB MB)"
        "  winre.wim             $wimGrootteMB MB   ($winrePad)"
    )) {
        $lbl          = New-Object System.Windows.Forms.Label
        $lbl.Text     = $tekst
        $lbl.Location = New-Object System.Drawing.Point(14, $y)
        $lbl.Size     = New-Object System.Drawing.Size(470, 18)
        $lbl.Font     = New-Object System.Drawing.Font("Consolas", 9)
        $herstelForm.Controls.Add($lbl)
        $y += 20
    }

    $y += 14

    $lblHWatLabel          = New-Object System.Windows.Forms.Label
    $lblHWatLabel.Text     = "Wat er gebeurt:"
    $lblHWatLabel.Location = New-Object System.Drawing.Point(14, $y)
    $lblHWatLabel.Size     = New-Object System.Drawing.Size(470, 18)
    $lblHWatLabel.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $herstelForm.Controls.Add($lblHWatLabel)
    $y += 22

    foreach ($stap in @(
        "  1. Nieuwe recovery partitie aanmaken ($herstelMB MB)"
        "  2. winre.wim kopiëren naar de nieuwe partitie"
        "  3. WinRE heractiveren"
    )) {
        $lbl           = New-Object System.Windows.Forms.Label
        $lbl.Text      = $stap
        $lbl.Location  = New-Object System.Drawing.Point(14, $y)
        $lbl.Size      = New-Object System.Drawing.Size(470, 18)
        $lbl.ForeColor = [System.Drawing.Color]::DimGray
        $herstelForm.Controls.Add($lbl)
        $y += 18
    }

    $y += 36

    $btnHOK              = New-Object System.Windows.Forms.Button
    $btnHOK.Text         = "Uitvoeren"
    $btnHOK.Location     = New-Object System.Drawing.Point(298, $y)
    $btnHOK.Size         = New-Object System.Drawing.Size(90, 30)
    $btnHOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnHCancel              = New-Object System.Windows.Forms.Button
    $btnHCancel.Text         = "Annuleren"
    $btnHCancel.Location     = New-Object System.Drawing.Point(400, $y)
    $btnHCancel.Size         = New-Object System.Drawing.Size(90, 30)
    $btnHCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $herstelForm.Controls.AddRange(@($btnHOK, $btnHCancel))
    $herstelForm.AcceptButton = $btnHOK
    $herstelForm.CancelButton = $btnHCancel
    [int]$herstelFormH      = $y + 50
    $herstelForm.ClientSize = New-Object System.Drawing.Size(510, $herstelFormH)

    if ($herstelForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "`n  Geannuleerd." -ForegroundColor Yellow
        exit
    }

    # Uitvoering
    $foutStap   = $null
    $nieuwePart = $null
    $tempLetter = $null

    # Stap 1 — Recovery partitie aanmaken
    Write-Step "Recovery partitie aanmaken ($herstelMB MB)..."

    $gebruikteBrieven = (Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue).Name
    $tempLetter = 'Z','Y','X','W','V','U','T','S','R' |
        Where-Object { $_ -notin $gebruikteBrieven } | Select-Object -First 1

    if (-not $tempLetter) {
        Write-Err "Geen vrije driveletter beschikbaar."
        $foutStap = "recovery partitie aanmaken"
    } else {
        Invoke-Diskpart @(
            "select disk $diskNum"
            "create partition primary size=$herstelMB"
            "format quick fs=ntfs label=`"WinRE`""
            "set id=de94bba4-06d1-4d40-a16a-bfd50179d6ac"
            "gpt attributes=0x8000000000000001"
            "assign letter=$tempLetter"
        ) | Out-Null

        Start-Sleep -Seconds 2

        $nieuwePart = Get-Partition -DiskNumber $diskNum -ErrorAction SilentlyContinue |
            Where-Object { $_.GptType -eq $recoveryGuid } | Select-Object -Last 1

        if (-not $nieuwePart) {
            Write-Err "Recovery partitie niet gevonden na aanmaken."
            $foutStap = "recovery partitie aanmaken"
        } else {
            Write-Success "Recovery partitie aangemaakt (partitie $($nieuwePart.PartitionNumber))"
        }
    }

    # Stap 2 — winre.wim kopiëren
    if (-not $foutStap) {
        Write-Step "winre.wim kopiëren naar ${tempLetter}:..."
        try {
            $doelPad = "${tempLetter}:\Recovery\WindowsRE"
            New-Item -Path $doelPad -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Copy-Item -Path $winrePad -Destination "$doelPad\winre.wim" -Force -ErrorAction Stop
            Write-Success "winre.wim gekopieerd naar $doelPad"
        } catch {
            Write-Err "Kopiëren winre.wim mislukt: $_"
            $foutStap = "winre.wim kopiëren"
        }
    }

    # Stap 3 — WinRE heractiveren
    if (-not $foutStap) {
        Write-Step "WinRE heractiveren..."
        $enableOut = & reagentc.exe /enable 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Success "WinRE actief"
        } else {
            Write-Warn "reagentc /enable: $enableOut"
        }
    }

    # Driveletter verwijderen
    if ($nieuwePart -and $tempLetter) {
        Write-Step "Tijdelijke driveletter $tempLetter verwijderen..."
        Invoke-Diskpart @(
            "select disk $diskNum"
            "select partition $($nieuwePart.PartitionNumber)"
            "remove letter=$tempLetter"
        ) | Out-Null
        Write-Success "Driveletter verwijderd"
    }

    # Resultaat
    Write-Host ""
    Write-Host "  $('-' * 50)" -ForegroundColor DarkGray

    $reagentInfoRaw = & reagentc.exe /info 2>&1
    $reagentStatus  = ($reagentInfoRaw | Select-String "Windows RE status") -replace '\s+', ' '
    if (-not $reagentStatus) { $reagentStatus = "(status onbekend)" }

    if ($foutStap) {
        Write-Host "  Gedeeltelijk voltooid — '$foutStap' mislukt." -ForegroundColor Yellow
        Write-Host "  winre.wim bewaard op: $winrePad" -ForegroundColor Yellow
        [System.Windows.Forms.MessageBox]::Show(
            "Herstel gedeeltelijk voltooid.`n`n" +
            "Stap mislukt : $foutStap`n" +
            "winre.wim op : $winrePad`n`n" +
            "Voer handmatig uit:`n  reagentc /enable",
            "Gedeeltelijk voltooid", "OK", "Warning"
        )
    } else {
        Write-Host "  Klaar!" -ForegroundColor Green
        Write-Host "  Recovery partitie hersteld: $herstelMB MB" -ForegroundColor Green
        Write-Host "  $reagentStatus" -ForegroundColor White
        Write-Host "  $('-' * 50)" -ForegroundColor DarkGray
        [System.Windows.Forms.MessageBox]::Show(
            "Recovery partitie hersteld!`n`n" +
            "Partitie grootte : $herstelMB MB`n" +
            "$reagentStatus`n`n" +
            "Geen reboot vereist.",
            "Recovery hersteld", "OK", "Information"
        )
    }
}
