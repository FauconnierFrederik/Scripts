# ============================================================
# Exchange Online - Online Archief activeren voor een mailbox
# Vereiste module: ExchangeOnlineManagement
# Installeer via: Install-Module -Name ExchangeOnlineManagement
# ============================================================

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

# --- Variabelen ---
$mailbox = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Geef het e-mailadres in waarvoor het Online Archief geactiveerd moet worden:",
    "Exchange Online - Archief activeren",
    ""
)

if ([string]::IsNullOrWhiteSpace($mailbox)) {
    [System.Windows.Forms.MessageBox]::Show("Geen e-mailadres ingegeven. Script wordt gestopt.", "Archief activeren", "OK", "Error")
    exit
}
$tagName       = "Move to Archive - 3 Years"
$policyName    = "Retention Policy - Archive 3 Years"
$maxRetries    = 3
$retryDelay    = 180  # seconden (3 minuten)

# --- Stap 1: Verbinding maken met Exchange Online ---
Write-Host "Verbinding maken met Exchange Online..." -ForegroundColor Cyan
Connect-ExchangeOnline -ShowProgress $true

# --- Stap 2: Retention Tag aanmaken (Move to Archive na 3 jaar) ---
Write-Host "Retention Tag aanmaken..." -ForegroundColor Cyan

$existingTag = Get-RetentionPolicyTag -Identity $tagName -ErrorAction SilentlyContinue
if (-not $existingTag) {
    New-RetentionPolicyTag `
        -Name $tagName `
        -Type All `
        -RetentionEnabled $true `
        -AgeLimitForRetention 1095 `
        -RetentionAction MoveToArchive
    Write-Host "Retention Tag '$tagName' aangemaakt." -ForegroundColor Green
} else {
    Write-Host "Retention Tag '$tagName' bestaat al, overgeslagen." -ForegroundColor Yellow
}

# --- Stap 3: Retention Policy aanmaken en tag koppelen ---
Write-Host "Retention Policy aanmaken..." -ForegroundColor Cyan

$existingPolicy = Get-RetentionPolicy -Identity $policyName -ErrorAction SilentlyContinue
if (-not $existingPolicy) {
    New-RetentionPolicy `
        -Name $policyName `
        -RetentionPolicyTagLinks $tagName
    Write-Host "Retention Policy '$policyName' aangemaakt." -ForegroundColor Green
} else {
    Write-Host "Retention Policy '$policyName' bestaat al, overgeslagen." -ForegroundColor Yellow
}

# --- Stap 4: Online Archief activeren voor de mailbox ---
Write-Host "Online Archief activeren voor $mailbox..." -ForegroundColor Cyan
Enable-Mailbox -Identity $mailbox -Archive
Write-Host "Archief geactiveerd." -ForegroundColor Green

# --- Stap 5: Retention Policy toepassen op de mailbox ---
Write-Host "Retention Policy toepassen op $mailbox..." -ForegroundColor Cyan
Set-Mailbox -Identity $mailbox -RetentionPolicy $policyName
Write-Host "Retention Policy '$policyName' toegepast op $mailbox." -ForegroundColor Green

# --- Verificatie ---
Write-Host "`n--- Verificatie ---" -ForegroundColor Cyan
Get-Mailbox -Identity $mailbox | Select-Object DisplayName, ArchiveStatus, RetentionPolicy

# --- Stap 6: Managed Folder Assistant met retry-logica ---
Write-Host "`nManaged Folder Assistant starten op $mailbox..." -ForegroundColor Cyan

$attempt = 0
$success = $false

while ($attempt -lt $maxRetries -and -not $success) {
    $attempt++
    Write-Host "Poging $attempt van $maxRetries..." -ForegroundColor Cyan
    try {
        Start-ManagedFolderAssistant -Identity $mailbox -ErrorAction Stop
        Write-Host "Managed Folder Assistant succesvol gestart op poging $attempt." -ForegroundColor Green
        $success = $true
    } catch {
        Write-Host "Poging $attempt mislukt: $_" -ForegroundColor Red
        if ($attempt -lt $maxRetries) {
            Write-Host "Wacht $retryDelay seconden voor volgende poging..." -ForegroundColor Yellow
            Start-Sleep -Seconds $retryDelay
        } else {
            Write-Host "Managed Folder Assistant kon niet worden gestart na $maxRetries pogingen." -ForegroundColor Red
            Write-Host "Dit is niet kritiek - de assistent zal automatisch 's nachts worden uitgevoerd." -ForegroundColor Yellow
        }
    }
}

Write-Host "`nKlaar! Het archief is actief en de retentiepolicy is ingesteld." -ForegroundColor Green

# --- Verbinding verbreken ---
Disconnect-ExchangeOnline -Confirm:$false
