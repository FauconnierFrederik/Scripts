#Requires -RunAsAdministrator
# ============================================================
#  Setup-DomainController.ps1
#  Maakt van een kale Windows Server een nieuwe primaire
#  Domain Controller (eerste DC in een nieuw forest).
#
#  Wat doet dit script:
#    1. Invoer verzamelen via popups (domein, IP, wachtwoord)
#    2. Statisch IP instellen (optioneel)
#    3. AD DS-rol installeren
#    4. Server promoveren tot DC
#    5. Automatische herstart → domein is actief
#
#  Uitvoeren als Administrator op de toekomstige DC-server.
#  Compatibel met: Windows Server 2016 / 2019 / 2022 / 2025
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# ────────────────────────────────────────────────────────────
#  HULPFUNCTIES
# ────────────────────────────────────────────────────────────

function Write-Step    { param([string]$Msg) Write-Host "`n[*] $Msg" -ForegroundColor Cyan }
function Write-Success { param([string]$Msg) Write-Host "    [OK] $Msg" -ForegroundColor Green }
function Write-Warn    { param([string]$Msg) Write-Host "    [!]  $Msg" -ForegroundColor Yellow }
function Write-Err     { param([string]$Msg) Write-Host "    [!!] $Msg" -ForegroundColor Red }

function Get-MaskedInput {
    param([string]$Titel, [string]$Label)

    $form                  = New-Object System.Windows.Forms.Form
    $form.Text             = $Titel
    $form.Size             = New-Object System.Drawing.Size(430, 175)
    $form.StartPosition    = "CenterScreen"
    $form.FormBorderStyle  = "FixedDialog"
    $form.MaximizeBox      = $false
    $form.MinimizeBox      = $false

    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Text     = $Label
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.Size     = New-Object System.Drawing.Size(400, 38)

    $txt              = New-Object System.Windows.Forms.TextBox
    $txt.PasswordChar = [char]0x2022
    $txt.Location     = New-Object System.Drawing.Point(12, 56)
    $txt.Size         = New-Object System.Drawing.Size(395, 24)

    $btnOK              = New-Object System.Windows.Forms.Button
    $btnOK.Text         = "OK"
    $btnOK.Location     = New-Object System.Drawing.Point(228, 100)
    $btnOK.Size         = New-Object System.Drawing.Size(85, 28)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel              = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Annuleren"
    $btnCancel.Location     = New-Object System.Drawing.Point(322, 100)
    $btnCancel.Size         = New-Object System.Drawing.Size(85, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $form.Controls.AddRange(@($lbl, $txt, $btnOK, $btnCancel))
    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $txt.Text }
    return $null
}

function Get-NicKeuze {
    $nics = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
    if ($nics.Count -eq 1) { return $nics[0] }

    $form                 = New-Object System.Windows.Forms.Form
    $form.Text            = "DC Setup – Netwerkadapter kiezen"
    $form.Size            = New-Object System.Drawing.Size(480, 220)
    $form.StartPosition   = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox     = $false

    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Text     = "Kies de adapter waarvoor het statische IP ingesteld wordt:"
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.Size     = New-Object System.Drawing.Size(450, 28)

    $list          = New-Object System.Windows.Forms.ListBox
    $list.Location = New-Object System.Drawing.Point(12, 44)
    $list.Size     = New-Object System.Drawing.Size(450, 110)
    foreach ($nic in $nics) {
        $list.Items.Add("$($nic.Name)  –  $($nic.InterfaceDescription)") | Out-Null
    }
    $list.SelectedIndex = 0

    $btnOK              = New-Object System.Windows.Forms.Button
    $btnOK.Text         = "OK"
    $btnOK.Location     = New-Object System.Drawing.Point(378, 162)
    $btnOK.Size         = New-Object System.Drawing.Size(84, 28)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $form.Controls.AddRange(@($lbl, $list, $btnOK))
    $form.AcceptButton = $btnOK

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $nics[$list.SelectedIndex] }
    return $null
}

# ────────────────────────────────────────────────────────────
#  STAP 0 – Check: al een DC?
# ────────────────────────────────────────────────────────────

$domainRole = (Get-WmiObject Win32_ComputerSystem).DomainRole
if ($domainRole -ge 4) {
    [System.Windows.Forms.MessageBox]::Show(
        "Deze server is al een Domain Controller.`nScript wordt gestopt.",
        "DC Setup", "OK", "Warning"
    )
    exit
}

Write-Host ""
Write-Host "  DC Setup – UViON" -ForegroundColor Cyan
Write-Host "  $('─' * 50)" -ForegroundColor DarkGray
Write-Host ""

# ────────────────────────────────────────────────────────────
#  STAP 1 – Invoer verzamelen
# ────────────────────────────────────────────────────────────

Write-Step "Invoer verzamelen..."

# Domein FQDN
$DomainFQDN = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Volledige domeinnaam (FQDN):`n`nVoorbeelden:`n  bedrijf.local`n  corp.intern`n  klant.be",
    "DC Setup – Domeinnaam",
    "bedrijf.local"
)
if ([string]::IsNullOrWhiteSpace($DomainFQDN)) {
    [System.Windows.Forms.MessageBox]::Show("Geen domeinnaam ingegeven. Script gestopt.", "DC Setup", "OK", "Error"); exit
}
$DomainFQDN = $DomainFQDN.Trim().ToLower()

# NetBIOS naam (automatisch afgeleid van FQDN)
$NetBIOSDefault = (($DomainFQDN -split "\.")[0]).ToUpper()
if ($NetBIOSDefault.Length -gt 15) { $NetBIOSDefault = $NetBIOSDefault.Substring(0, 15) }

$NetBIOSName = [Microsoft.VisualBasic.Interaction]::InputBox(
    "NetBIOS naam van het domein (max. 15 tekens):`n`nDit is de korte domeinnaam, bijv. BEDRIJF`nGebruikt voor DOMEIN\Gebruiker-aanmeldingen.",
    "DC Setup – NetBIOS naam",
    $NetBIOSDefault
)
if ([string]::IsNullOrWhiteSpace($NetBIOSName)) {
    [System.Windows.Forms.MessageBox]::Show("Geen NetBIOS naam ingegeven. Script gestopt.", "DC Setup", "OK", "Error"); exit
}
$NetBIOSName = $NetBIOSName.Trim().ToUpper()

# DSRM wachtwoord (met bevestiging)
$DSRMPass   = $null
$pogingen   = 0
while ($null -eq $DSRMPass) {
    $pogingen++
    if ($pogingen -gt 3) {
        [System.Windows.Forms.MessageBox]::Show("Te veel pogingen. Script gestopt.", "DC Setup", "OK", "Error"); exit
    }

    $pw1 = Get-MaskedInput `
        -Titel "DC Setup – DSRM Wachtwoord" `
        -Label "Directory Services Restore Mode (DSRM) wachtwoord:`n`nVereisten: min. 8 tekens, hoofdletter, cijfer, speciaal teken."
    if ($null -eq $pw1) { exit }

    $pw2 = Get-MaskedInput `
        -Titel "DC Setup – DSRM Wachtwoord bevestigen" `
        -Label "Bevestig het DSRM wachtwoord:"
    if ($null -eq $pw2) { exit }

    if ($pw1 -ne $pw2) {
        [System.Windows.Forms.MessageBox]::Show("Wachtwoorden komen niet overeen. Probeer opnieuw.", "DC Setup", "OK", "Warning")
    } elseif ($pw1.Length -lt 8) {
        [System.Windows.Forms.MessageBox]::Show("Wachtwoord te kort (minimum 8 tekens).", "DC Setup", "OK", "Warning")
    } else {
        $DSRMPass = ConvertTo-SecureString $pw1 -AsPlainText -Force
    }
}

# Statisch IP instellen?
$staticIP     = $false
$IPAddress    = $null
$PrefixLength = $null
$Gateway      = $null
$selectedNic  = $null

$keuzeIP = [System.Windows.Forms.MessageBox]::Show(
    "Wil je een statisch IP-adres instellen op deze server?`n`nAanbevolen voor een Domain Controller.",
    "DC Setup – Statisch IP",
    "YesNo", "Question"
)

if ($keuzeIP -eq "Yes") {
    $staticIP    = $true
    $selectedNic = Get-NicKeuze
    if ($null -eq $selectedNic) { exit }

    $huidigIP = (Get-NetIPAddress -InterfaceIndex $selectedNic.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                 Where-Object { $_.PrefixOrigin -ne "WellKnown" }).IPAddress
    $huidigGW = (Get-NetRoute -InterfaceIndex $selectedNic.ifIndex -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue).NextHop

    $IPAddress = [Microsoft.VisualBasic.Interaction]::InputBox(
        "IP-adres voor deze server:`n`nVoorbeelden: 192.168.1.10  |  10.0.0.5",
        "DC Setup – IP-adres",
        $(if ($huidigIP) { $huidigIP } else { "192.168.1.10" })
    )
    if ([string]::IsNullOrWhiteSpace($IPAddress)) { exit }

    $PrefixStr = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Subnet prefix lengte:`n`n24 = 255.255.255.0  (/24)`n16 = 255.255.0.0    (/16)`n8  = 255.0.0.0      (/8)",
        "DC Setup – Subnetmasker",
        "24"
    )
    if ([string]::IsNullOrWhiteSpace($PrefixStr)) { exit }
    $PrefixLength = [int]($PrefixStr.Trim().TrimStart("/"))

    $Gateway = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Default Gateway (router):`n`nVoorbeelden: 192.168.1.1  |  10.0.0.1",
        "DC Setup – Gateway",
        $(if ($huidigGW) { $huidigGW } else { "192.168.1.1" })
    )
    if ([string]::IsNullOrWhiteSpace($Gateway)) { exit }
}

# ────────────────────────────────────────────────────────────
#  STAP 2 – Samenvatting en bevestiging
# ────────────────────────────────────────────────────────────

$samenvatting  = "Samenvatting DC-installatie`n"
$samenvatting += "══════════════════════════════════════`n`n"
$samenvatting += "Servernaam           : $($env:COMPUTERNAME)`n"
$samenvatting += "Domein (FQDN)        : $DomainFQDN`n"
$samenvatting += "NetBIOS naam         : $NetBIOSName`n"
$samenvatting += "Forest / Domain mode : WinThreshold (Server 2016+)`n"
$samenvatting += "DNS-server           : wordt op deze server geïnstalleerd`n"

if ($staticIP) {
    $samenvatting += "`nNetwerk adapter      : $($selectedNic.Name)`n"
    $samenvatting += "IP-adres             : $IPAddress/$PrefixLength`n"
    $samenvatting += "Gateway              : $Gateway`n"
}

$samenvatting += "`n══════════════════════════════════════`n"
$samenvatting += "Na voltooiing herstart de server automatisch.`n"
$samenvatting += "Na de herstart is het domein actief.`n`n"
$samenvatting += "Doorgaan met de installatie?"

$confirm = [System.Windows.Forms.MessageBox]::Show(
    $samenvatting, "DC Setup – Bevestiging", "YesNo", "Question"
)
if ($confirm -ne "Yes") {
    Write-Host "`n  Installatie geannuleerd door gebruiker." -ForegroundColor Yellow
    exit
}

# ────────────────────────────────────────────────────────────
#  STAP 3 – Statisch IP instellen
# ────────────────────────────────────────────────────────────

if ($staticIP) {
    Write-Step "Statisch IP instellen op adapter '$($selectedNic.Name)'..."

    Get-NetIPAddress -InterfaceIndex $selectedNic.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue

    Get-NetRoute -InterfaceIndex $selectedNic.ifIndex -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue |
        Remove-NetRoute -Confirm:$false -ErrorAction SilentlyContinue

    New-NetIPAddress `
        -InterfaceIndex $selectedNic.ifIndex `
        -IPAddress      $IPAddress `
        -PrefixLength   $PrefixLength `
        -DefaultGateway $Gateway | Out-Null

    # DNS tijdelijk naar eigen IP; wordt na promotie automatisch 127.0.0.1
    Set-DnsClientServerAddress -InterfaceIndex $selectedNic.ifIndex -ServerAddresses $IPAddress

    Write-Success "IP ingesteld: $IPAddress/$PrefixLength  –  Gateway: $Gateway"
}

# ────────────────────────────────────────────────────────────
#  STAP 4 – AD DS-rol installeren
# ────────────────────────────────────────────────────────────

Write-Step "AD DS-rol installeren (dit duurt enkele minuten)..."

$install = Install-WindowsFeature -Name AD-Domain-Services -IncludeManagementTools

if (-not $install.Success) {
    Write-Err "AD DS-rol kon niet worden geïnstalleerd!"
    [System.Windows.Forms.MessageBox]::Show(
        "De AD DS-rol kon niet worden geïnstalleerd.`nControleer de Windows Event Viewer voor details.",
        "DC Setup – Fout", "OK", "Error"
    )
    exit
}
Write-Success "AD DS-rol succesvol geïnstalleerd."

# ────────────────────────────────────────────────────────────
#  STAP 5 – Server promoveren tot Domain Controller
# ────────────────────────────────────────────────────────────

Write-Step "Server promoveren tot Domain Controller voor '$DomainFQDN'..."
Write-Warn "De server herstart automatisch na voltooiing van de promotie!"
Write-Host ""

Import-Module ADDSDeployment -ErrorAction Stop

Install-ADDSForest `
    -DomainName                    $DomainFQDN `
    -DomainNetbiosName             $NetBIOSName `
    -DomainMode                    "WinThreshold" `
    -ForestMode                    "WinThreshold" `
    -InstallDns:$true `
    -CreateDnsDelegation:$false `
    -DatabasePath                  "C:\Windows\NTDS" `
    -LogPath                       "C:\Windows\NTDS" `
    -SysvolPath                    "C:\Windows\SYSVOL" `
    -SafeModeAdministratorPassword $DSRMPass `
    -NoRebootOnCompletion:$false `
    -Force:$true

# Onderstaande code wordt enkel bereikt als NoRebootOnCompletion:$true was
Write-Host ""
Write-Success "Promotie voltooid. De server herstart nu..."
