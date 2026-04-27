# ============================================================
# Setup-RDP.ps1
# Generiek script voor aanmaken self-signed certificaat,
# RDP bestand en signing - bruikbaar bij elke klant
# Uitvoeren als Administrator op de TS-server
# ============================================================

# ---- CONFIGURATIE ----
$CertGeldigheid = 2    # Geldigheid certificaat in jaren
$SysvolPad      = ""   # Optioneel: \\domein\SYSVOL\domein\Certificates\
# ----------------------

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

$Hostname = $env:COMPUTERNAME
$RDPPad   = "C:\Scripts\$Hostname.rdp"
$CertPad  = "C:\Scripts\$Hostname.cer"

# ------------------------------------------------------------
# POPUP - Keuze: volledig of alleen RDP opnieuw aanmaken
# ------------------------------------------------------------
$keuze = [System.Windows.Forms.MessageBox]::Show(
    "Wat wil je uitvoeren?`n`nJa  = Volledig setup (certificaat + RDP aanmaken + signen)`nNee = Alleen RDP bestand opnieuw aanmaken en signen",
    "RDP Setup - Keuze",
    "YesNo",
    "Question"
)

# ------------------------------------------------------------
# DNS naam opvragen via popup
# ------------------------------------------------------------
# Bij alleen RDP: bestaande DNS naam proberen op te halen uit huidig RDP bestand
$defaultDns = $Hostname
if ($keuze -eq "No" -and (Test-Path $RDPPad)) {
    $bestaandeDns = Get-Content $RDPPad | Where-Object { $_ -like "full address*" } | Select-Object -First 1
    if ($bestaandeDns) {
        $defaultDns = $bestaandeDns -replace "full address:s:", ""
    }
}

$DnsNaam = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Voer de DNS naam in voor de RDP verbinding:`n`nVoorbeelden:`n  server.dyndns.org`n  SERVER02.domein.local",
    "RDP Setup - DNS naam",
    $defaultDns
)

if ([string]::IsNullOrWhiteSpace($DnsNaam)) {
    [System.Windows.Forms.MessageBox]::Show("Geen DNS naam ingegeven. Script wordt gestopt.", "Setup-RDP", "OK", "Error")
    exit
}

Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Setup RDP - $( if ($keuze -eq 'Yes') { 'Volledig' } else { 'Alleen RDP' } )" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Hostname   : $Hostname"
Write-Host " DNS naam   : $DnsNaam"
Write-Host " RDP bestand: $RDPPad"
Write-Host " Certificaat: $CertPad"
Write-Host "============================================" -ForegroundColor Cyan

# ------------------------------------------------------------
# VOLLEDIG SETUP - certificaat aanmaken en koppelen
# ------------------------------------------------------------
if ($keuze -eq "Yes") {

    # STAP 1 - Certificaat aanmaken
    Write-Host "`n[1/3] Certificaat aanmaken..." -ForegroundColor Yellow

    Get-ChildItem Cert:\LocalMachine\My | Where-Object {
        $_.Subject -like "*$Hostname*" -and $_.FriendlyName -like "RDP*"
    } | Remove-Item -Force -ErrorAction SilentlyContinue

    $cert = New-SelfSignedCertificate `
        -DnsName $DnsNaam, $Hostname `
        -CertStoreLocation "Cert:\LocalMachine\My" `
        -NotAfter (Get-Date).AddYears($CertGeldigheid) `
        -KeyUsage KeyEncipherment, DataEncipherment `
        -FriendlyName "RDP $Hostname"

    $thumbprint = $cert.Thumbprint
    Write-Host "    Certificaat aangemaakt: $thumbprint" -ForegroundColor Green
    Write-Host "    Geldig tot            : $($cert.NotAfter.ToString('dd/MM/yyyy'))" -ForegroundColor Green

    # STAP 2 - Certificaat koppelen aan RDP-Tcp
    Write-Host "`n[2/3] Certificaat koppelen aan RDP..." -ForegroundColor Yellow

    try {
        $path = (Get-WmiObject -Class "Win32_TSGeneralSetting" `
            -Namespace root\cimv2\terminalservices `
            -Filter "TerminalName='RDP-Tcp'").__path
        Set-WmiInstance -Path $path -Argument @{SSLCertificateSHA1Hash=$thumbprint} | Out-Null
        Write-Host "    Certificaat gekoppeld aan RDP-Tcp" -ForegroundColor Green
    } catch {
        Write-Host "    Fout bij koppelen: $_" -ForegroundColor Red
    }

    # Certificaat exporteren
    $certMap = Split-Path $CertPad
    if (!(Test-Path $certMap)) { New-Item -ItemType Directory -Path $certMap -Force | Out-Null }
    Export-Certificate -Cert "Cert:\LocalMachine\My\$thumbprint" -FilePath $CertPad -Force | Out-Null
    Write-Host "    Certificaat geexporteerd naar: $CertPad" -ForegroundColor Green

    # Optioneel naar SYSVOL kopieren
    if ($SysvolPad -ne "") {
        try {
            $sysvolMap = Split-Path $SysvolPad
            if (!(Test-Path $sysvolMap)) { New-Item -ItemType Directory -Path $sysvolMap -Force | Out-Null }
            Copy-Item -Path $CertPad -Destination $SysvolPad -Force
            Write-Host "    Certificaat gekopieerd naar SYSVOL: $SysvolPad" -ForegroundColor Green
        } catch {
            Write-Host "    Kon niet kopieren naar SYSVOL: $_" -ForegroundColor Red
        }
    }

    $stapRDP = "[3/3]"

} else {

    # Alleen RDP: bestaand certificaat ophalen
    Write-Host "`n[1/1] Bestaand certificaat ophalen..." -ForegroundColor Yellow

    $cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object {
        $_.FriendlyName -like "RDP*" -and $_.Subject -like "*$Hostname*"
    } | Sort-Object NotAfter -Descending | Select-Object -First 1

    if ($null -eq $cert) {
        [System.Windows.Forms.MessageBox]::Show(
            "Geen bestaand RDP certificaat gevonden voor $Hostname.`nStart het script opnieuw en kies 'Ja' voor een volledige setup.",
            "Setup-RDP - Fout", "OK", "Error"
        )
        exit
    }

    $thumbprint = $cert.Thumbprint
    Write-Host "    Certificaat gevonden: $thumbprint" -ForegroundColor Green
    Write-Host "    Geldig tot          : $($cert.NotAfter.ToString('dd/MM/yyyy'))" -ForegroundColor Green

    $stapRDP = "[2/2]"
}

# ------------------------------------------------------------
# GPO fixes - printer redirectie en EasyPrint (alleen bij volledig setup)
# ------------------------------------------------------------
if ($keuze -eq "Yes") {
    Write-Host "`n[+] GPO fixes toepassen voor printer redirectie..." -ForegroundColor Yellow

    $rdsPad = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services"

    # Maak registry pad aan indien het nog niet bestaat
    if (!(Test-Path $rdsPad)) {
        New-Item -Path $rdsPad -Force | Out-Null
    }

    # Printer redirectie toestaan (0 = toestaan, 1 = blokkeren)
    Set-ItemProperty -Path $rdsPad -Name "fDisableCpm" -Value 0 -Type DWord
    Write-Host "    Printer redirectie toegestaan" -ForegroundColor Green

    # EasyPrint als standaard driver uitschakelen zodat lokale drivers gebruikt worden
    Set-ItemProperty -Path $rdsPad -Name "fEnableEasyPrint" -Value 1 -Type DWord
    Write-Host "    EasyPrint driver inschakelen" -ForegroundColor Green

    # Altijd om wachtwoord vragen uitschakelen
    Set-ItemProperty -Path $rdsPad -Name "fPromptForPassword" -Value 0 -Type DWord
    Write-Host "    Altijd om wachtwoord vragen uitgeschakeld" -ForegroundColor Green

    # Beveiligingslaag instellen op RDP (2 = SSL/TLS)
    Set-ItemProperty -Path $rdsPad -Name "SecurityLayer" -Value 2 -Type DWord
    Write-Host "    Beveiligingslaag ingesteld op SSL/TLS" -ForegroundColor Green

    # Print Spooler service controleren en herstarten indien nodig
    $spooler = Get-Service -Name Spooler
    if ($spooler.Status -ne "Running") {
        Start-Service -Name Spooler
        Write-Host "    Print Spooler gestart" -ForegroundColor Green
    } else {
        Write-Host "    Print Spooler actief" -ForegroundColor Green
    }

    # GPO forceren
    & gpupdate /force | Out-Null
    Write-Host "    GPO bijgewerkt via gpupdate /force" -ForegroundColor Green
}

# ------------------------------------------------------------
# RDP bestand aanmaken
# ------------------------------------------------------------
Write-Host "`n$stapRDP RDP bestand aanmaken..." -ForegroundColor Yellow

$rdpMap = Split-Path $RDPPad
if (!(Test-Path $rdpMap)) { New-Item -ItemType Directory -Path $rdpMap -Force | Out-Null }

$rdpInhoud = @"
screen mode id:i:2
use multimon:i:0
desktopwidth:i:1920
desktopheight:i:1080
session bpp:i:32
compression:i:1
keyboardhook:i:2
audiocapturemode:i:0
videoplaybackmode:i:1
connection type:i:7
networkautodetect:i:1
bandwidthautodetect:i:1
displayconnectionbar:i:1
autoreconnection enabled:i:1
full address:s:$DnsNaam
alternate full address:s:$DnsNaam
audiomode:i:0
redirectprinters:i:1
redirectclipboard:i:1
redirectsmartcards:i:1
redirectcomports:i:0
redirectposdevices:i:0
authentication level:i:2
prompt for credentials:i:0
enablecredsspsupport:i:1
negotiate security layer:i:1
remoteapplicationmode:i:0
savecredentials:i:1
alternate shell:s:
shell working directory:s:
gatewayhostname:s:
gatewayusagemethod:i:4
gatewaycredentialssource:i:4
gatewayprofileusagemethod:i:0
promptcredentialonce:i:0
drivestoredirect:s:
"@

$rdpInhoud | Out-File -FilePath $RDPPad -Encoding unicode
Write-Host "    RDP bestand aangemaakt: $RDPPad" -ForegroundColor Green

# ------------------------------------------------------------
# RDP bestand signen
# ------------------------------------------------------------
Write-Host "`n$( if ($keuze -eq 'Yes') { '[3/3]' } else { '[2/2]' } ) RDP bestand signen..." -ForegroundColor Yellow

$signResult = & rdpsign /sha256 $thumbprint "$RDPPad" 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "    RDP bestand succesvol gesigned" -ForegroundColor Green
} else {
    Write-Host "    Fout bij signen: $signResult" -ForegroundColor Red
}

# ------------------------------------------------------------
# RDP service herstarten (alleen bij volledig setup)
# ------------------------------------------------------------
if ($keuze -eq "Yes") {
    Write-Host "`nRDP service herstarten..." -ForegroundColor Yellow
    Restart-Service TermService -Force
    Write-Host "    RDP service herstart" -ForegroundColor Green
}

# ------------------------------------------------------------
# SAMENVATTING popup
# ------------------------------------------------------------
$samenvatting = "Setup voltooid!`n`n"
$samenvatting += "Hostname               : $Hostname`n"
$samenvatting += "DNS naam               : $DnsNaam`n"
$samenvatting += "Certificaat thumbprint : $thumbprint`n"
$samenvatting += "Certificaat geldig tot : $($cert.NotAfter.ToString('dd/MM/yyyy'))`n"
$samenvatting += "RDP bestand locatie    : $RDPPad`n"

if ($keuze -eq "Yes") {
    $samenvatting += "Certificaat locatie    : $CertPad`n"
    $samenvatting += "`nVolgende stappen:`n"
    $samenvatting += "1. Importeer $Hostname.cer op de clientpc's via GPO of certlm.msc`n"
    $samenvatting += "2. Kopieer $Hostname.rdp naar de clientpc's"
} else {
    $samenvatting += "`nKopieer $Hostname.rdp naar de clientpc's"
}

[System.Windows.Forms.MessageBox]::Show($samenvatting, "Setup-RDP - Klaar", "OK", "Information")
