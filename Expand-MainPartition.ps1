#Requires -RunAsAdministrator
# ============================================================
#  Expand-MainPartition.ps1
#  Breidt C: uit door de recovery partitie tijdelijk te
#  verwijderen en daarna terug te plaatsen. Geen reboot nodig.
#
#  Vereisten: GPT schijf, recovery partitie aanwezig,
#             unallocated space aanwezig op dezelfde schijf.
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
#  PRE-FLIGHT CHECKS
# ------------------------------------------------------------

$recoveryGuid = '{de94bba4-06d1-4d40-a16a-bfd50179d6ac}'

try {
    $cPart   = Get-Partition -DriveLetter C -ErrorAction Stop
} catch {
    [System.Windows.Forms.MessageBox]::Show("Partitie C: niet gevonden.", "Fout", "OK", "Error")
    exit
}

$diskNum = $cPart.DiskNumber
$disk    = Get-Disk -Number $diskNum -ErrorAction Stop

if ($disk.PartitionStyle -ne 'GPT') {
    [System.Windows.Forms.MessageBox]::Show(
        "Schijf $diskNum is MBR. Dit script ondersteunt enkel GPT-schijven.",
        "Niet ondersteund", "OK", "Error"
    )
    exit
}

$recoveryPart = Get-Partition -DiskNumber $diskNum -ErrorAction SilentlyContinue |
    Where-Object { $_.GptType -eq $recoveryGuid } |
    Select-Object -Last 1

if (-not $recoveryPart) {
    [System.Windows.Forms.MessageBox]::Show(
        "Geen Windows Recovery partitie gevonden op schijf $diskNum.`n`n" +
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

# ------------------------------------------------------------
#  BEVESTIGINGSSCHERM
# ------------------------------------------------------------

$cGrootteGB  = [math]::Round($cPart.Size / 1GB, 1)
$cNaGB       = [math]::Round(($cPart.Size + $vrijSpace) / 1GB, 1)
$recoveryMB  = [math]::Round($recoveryPart.Size / 1MB, 0)
$vrijMB      = [math]::Round($vrijSpace / 1MB, 0)

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
    "  3. C: uitbreiden naar maximum"
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

# ------------------------------------------------------------
#  UITVOERING
# ------------------------------------------------------------

$winrePad   = "C:\Windows\System32\Recovery\winre.wim"
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

# Stap 2 — Recovery partitie verwijderen (bevestiging)
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

# Stap 3 — C: uitbreiden
Write-Step "C: uitbreiden naar maximum..."
try {
    $maxSize = (Get-PartitionSupportedSize -DriveLetter C -ErrorAction Stop).SizeMax
    Resize-Partition -DriveLetter C -Size $maxSize -ErrorAction Stop
    $nieuweC = [math]::Round((Get-Partition -DriveLetter C).Size / 1GB, 1)
    Write-Success "C: uitgebreid naar $nieuweC GB"
} catch {
    Write-Err "Uitbreiden mislukt: $_"
    $foutStap = "C: uitbreiden"
}

# Stap 4 — Nieuwe recovery partitie aanmaken
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

# Stap 5 — WinRE terugzetten
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

# Stap 6 — WinRE heractiveren
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

# ------------------------------------------------------------
#  RESULTAAT
# ------------------------------------------------------------

Write-Host ""
Write-Host "  $('-' * 50)" -ForegroundColor DarkGray

$eindC = [math]::Round((Get-Partition -DriveLetter C -ErrorAction SilentlyContinue).Size / 1GB, 1)
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
