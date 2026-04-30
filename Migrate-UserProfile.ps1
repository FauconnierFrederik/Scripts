#Requires -Version 5.1
<#
.SYNOPSIS
    Migrate-UserProfile.ps1 - GUI tool voor gebruikersprofiel migratie naar nieuwe computer.
.DESCRIPTION
    Migreert: printers (PrintBrm), browser profielen (Chrome/Edge/Firefox),
    gebruikersbestanden, Outlook handtekeningen + PST, mapped drives,
    WiFi profielen, Sticky Notes, achtergrond, omgevingsvariabelen.
    Ondersteunt lokale en remote broncomputer via C$ share.
.NOTES
    Vereist administrator rechten.
    Remote migratie vereist admin credentials voor C$ share op broncomputer.
#>

[CmdletBinding()]
param()

# ─── Admin check – herstart verhoogd als nodig ────────────────────────────────
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $selfArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Start-Process powershell -Verb RunAs -ArgumentList $selfArgs
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ─── Globals ──────────────────────────────────────────────────────────────────
$script:LogEntries   = [System.Collections.Generic.List[string]]::new()
$script:TempMount    = $null
$script:BackupRoot   = $null
$script:SourceCred   = $null
$script:LogBox       = $null
$script:ProgressBar  = $null
$script:StatusLabel  = $null
$script:TotalSteps   = 0
$script:CurrentStep  = 0

# ─── Helper: logging ──────────────────────────────────────────────────────────
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','OK')]$Level = 'INFO'
    )
    $ts   = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts][$Level] $Message"
    $script:LogEntries.Add($line)

    if ($script:LogBox) {
        $color = switch ($Level) {
            'OK'    { [System.Drawing.Color]::LightGreen }
            'WARN'  { [System.Drawing.Color]::Yellow }
            'ERROR' { [System.Drawing.Color]::Tomato }
            default { [System.Drawing.Color]::LightGray }
        }
        $script:LogBox.SelectionStart  = $script:LogBox.TextLength
        $script:LogBox.SelectionLength = 0
        $script:LogBox.SelectionColor  = $color
        $script:LogBox.AppendText("$line`n")
        $script:LogBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

# ─── Helper: zip met DoEvents per bestand (UI blijft responsief) ─────────────
function Compress-WithProgress {
    param([string]$SourceDir, [string]$ZipPath)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $files = Get-ChildItem $SourceDir -Recurse -File -ErrorAction SilentlyContinue
    if (-not $files) { return }

    $zip = [System.IO.Compression.ZipFile]::Open($ZipPath, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $done = 0
        foreach ($file in $files) {
            $entryName = $file.FullName.Substring($SourceDir.TrimEnd('\').Length + 1)
            try {
                [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                    $zip, $file.FullName, $entryName,
                    [System.IO.Compression.CompressionLevel]::Optimal
                ) | Out-Null
            } catch {
                # Sla vergrendelde bestanden over
            }
            $done++
            if ($done % 20 -eq 0) {
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
    } finally {
        $zip.Dispose()
    }
}

# ─── Helper: voortgang ────────────────────────────────────────────────────────
function Update-Progress {
    param([string]$Status)
    $script:CurrentStep++
    if ($script:ProgressBar -and $script:TotalSteps -gt 0) {
        $pct = [int](($script:CurrentStep / $script:TotalSteps) * 100)
        $script:ProgressBar.Value = [Math]::Min($pct, 100)
        $script:StatusLabel.Text  = $Status
        [System.Windows.Forms.Application]::DoEvents()
    }
}

# ─── Helper: netwerk share ────────────────────────────────────────────────────
function Connect-RemoteShare {
    param([string]$ComputerName, [PSCredential]$Credential)
    $share = "\\$ComputerName\c$"
    Write-Log "Verbinding maken met $share..."

    # Verwijder eventuele bestaande verbinding
    net use $share /delete 2>$null | Out-Null

    if ($Credential) {
        $user   = $Credential.UserName
        $pass   = $Credential.GetNetworkCredential().Password
        $result = net use $share /user:$user $pass 2>&1
    } else {
        $result = net use $share 2>&1
    }

    if ($LASTEXITCODE -eq 0 -or (Test-Path $share)) {
        Write-Log "Verbonden met $share" -Level OK
        $script:TempMount = $share
        return $share
    } else {
        Write-Log "Verbinding mislukt: $result" -Level ERROR
        return $null
    }
}

function Disconnect-RemoteShare {
    if ($script:TempMount) {
        net use $script:TempMount /delete 2>$null | Out-Null
        Write-Log "Verbinding $($script:TempMount) verbroken"
        $script:TempMount = $null
    }
}

# ─── Migratie: Printers ───────────────────────────────────────────────────────
function Backup-Printers {
    param([string]$BackupPath, [string]$SourceComputer)
    Update-Progress "Printers exporteren..."

    $printbrmPaths = @(
        "$env:SystemRoot\System32\spool\tools\PrintBrm.exe",
        "$env:SystemRoot\SysWOW64\spool\tools\PrintBrm.exe"
    )
    $printbrm = $printbrmPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($printbrm -and $SourceComputer -eq $env:COMPUTERNAME) {
        $outFile = Join-Path $BackupPath "printers.printerExport"
        try {
            & $printbrm -B -F $outFile 2>&1 | ForEach-Object { Write-Log $_ }
            if (Test-Path $outFile) {
                Write-Log "Printers geexporteerd: $outFile" -Level OK
                Write-Log "Herstel op nieuwe PC: printbrm.exe -R -F `"$outFile`"" -Level INFO
            }
        } catch {
            Write-Log "PrintBrm fout: $_" -Level ERROR
        }
    } else {
        if (-not $printbrm) { Write-Log "PrintBrm.exe niet gevonden, CSV fallback gebruikt" -Level WARN }

        try {
            $wmiArgs = @{ Class = 'Win32_Printer'; ComputerName = $SourceComputer; ErrorAction = 'Stop' }
            if ($script:SourceCred -and $SourceComputer -ne $env:COMPUTERNAME) {
                $wmiArgs['Credential'] = $script:SourceCred
            }
            $printers = Get-WmiObject @wmiArgs |
                Where-Object { $_.Name -notmatch '^(Microsoft|Fax|OneNote|PDF)' } |
                Select-Object Name, PortName, DriverName, Default, Shared, ShareName

            $printers | Export-Csv "$BackupPath\printer_list.csv" -NoTypeInformation -Encoding UTF8

            # Genereer herstel script
            $restoreLines = @("# Printer herstel script`n# Gegenereerd: $(Get-Date)`n")
            foreach ($p in $printers) {
                $restoreLines += "# $($p.Name) -> $($p.PortName)"
                $restoreLines += "Add-Printer -Name '$($p.Name)' -PortName '$($p.PortName)' -DriverName '$($p.DriverName)'"
            }
            $restoreLines | Out-File "$BackupPath\Restore-Printers.ps1" -Encoding UTF8

            Write-Log "Printerlijst geexporteerd ($($printers.Count) printers) als CSV + herstel script" -Level OK
            if ($SourceComputer -ne $env:COMPUTERNAME) {
                Write-Log "Tip: voer PrintBrm lokaal op broncomputer uit voor volledige driver export" -Level WARN
            }
        } catch {
            Write-Log "Fout bij printerlijst exporteren: $_" -Level ERROR
        }
    }
}

# ─── Migratie: Chrome ─────────────────────────────────────────────────────────
function Backup-ChromeProfile {
    param([string]$BackupPath, [string]$SourceUserFolder)
    Update-Progress "Chrome profiel backuppen..."

    # Wachtwoord sync check
    $syncCheck = [System.Windows.Forms.MessageBox]::Show(
        "Zijn de Chrome wachtwoorden gesynchroniseerd met een Google account?`n`n" +
        "• Ja  → backup gaat direct door`n" +
        "• Nee → je krijgt instructies voor manuele export`n" +
        "• Annuleer → Chrome backup overslaan",
        "Chrome Wachtwoorden",
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($syncCheck -eq [System.Windows.Forms.DialogResult]::Cancel) {
        Write-Log "Chrome backup overgeslagen door gebruiker" -Level WARN
        return
    }

    if ($syncCheck -eq [System.Windows.Forms.DialogResult]::No) {
        [System.Windows.Forms.MessageBox]::Show(
            "Exporteer wachtwoorden manueel in Chrome:`n`n" +
            "1. Open Chrome op de BRONCOMPUTER`n" +
            "2. Ga naar: chrome://password-manager/passwords`n" +
            "3. Klik op het tandwiel-icoon (Instellingen)`n" +
            "4. Kies 'Wachtwoorden exporteren'`n" +
            "5. Sla het CSV bestand op in:`n   $BackupPath`n`n" +
            "Klik OK als je klaar bent, dan gaat de backup verder.",
            "Exporteer Chrome Wachtwoorden",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }

    $chromeSrc = Join-Path $SourceUserFolder "AppData\Local\Google\Chrome\User Data"
    if (-not (Test-Path $chromeSrc)) {
        Write-Log "Chrome profiel niet gevonden: $chromeSrc" -Level WARN
        return
    }

    $zipDest = Join-Path $BackupPath "Chrome_UserData.zip"
    $tempDir = Join-Path $env:TEMP "ChromeBackup_$(Get-Random)"
    try {
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

        # Kopieer zonder grote cache mappen voor kleinere backup
        $excludeDirs = 'Cache', 'Code Cache', 'GPUCache', 'ShaderCache', 'DawnCache',
                       'blob_storage', 'CacheStorage', 'Service Worker'
        $robocopyArgs = @($chromeSrc, $tempDir, '/E', '/XD') + $excludeDirs +
                        @('/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')
        & robocopy @robocopyArgs | Out-Null

        Write-Log "Chrome bestanden zippen..."
        Compress-WithProgress -SourceDir $tempDir -ZipPath $zipDest
        $size = [Math]::Round((Get-Item $zipDest).Length / 1MB, 1)
        Write-Log "Chrome profiel gezipt: $zipDest ($size MB)" -Level OK
    } catch {
        Write-Log "Chrome backup fout (probeer directe zip): $_" -Level WARN
        try {
            Compress-WithProgress -SourceDir $chromeSrc -ZipPath $zipDest
            Write-Log "Chrome profiel gezipt (met cache)" -Level OK
        } catch {
            Write-Log "Chrome backup mislukt: $_" -Level ERROR
        }
    } finally {
        if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

# ─── Migratie: Edge ───────────────────────────────────────────────────────────
function Backup-EdgeProfile {
    param([string]$BackupPath, [string]$SourceUserFolder)
    Update-Progress "Edge profiel backuppen..."

    $edgeSrc = Join-Path $SourceUserFolder "AppData\Local\Microsoft\Edge\User Data"
    if (-not (Test-Path $edgeSrc)) { Write-Log "Edge profiel niet gevonden" -Level WARN; return }

    $zipDest = Join-Path $BackupPath "Edge_UserData.zip"
    $tempDir = Join-Path $env:TEMP "EdgeBackup_$(Get-Random)"
    try {
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        $excludeDirs = 'Cache', 'Code Cache', 'GPUCache', 'Service Worker'
        $robocopyArgs = @($edgeSrc, $tempDir, '/E', '/XD') + $excludeDirs +
                        @('/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')
        & robocopy @robocopyArgs | Out-Null
        Write-Log "Edge bestanden zippen..."
        Compress-WithProgress -SourceDir $tempDir -ZipPath $zipDest
        if (Test-Path $zipDest) {
            $size = [Math]::Round((Get-Item $zipDest).Length / 1MB, 1)
            Write-Log "Edge profiel gezipt: $zipDest ($size MB)" -Level OK
        } else {
            Write-Log "Edge zip mislukt - sluit Edge en probeer opnieuw" -Level ERROR
        }
    } catch {
        Write-Log "Edge backup fout: $_" -Level ERROR
    } finally {
        if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

# ─── Migratie: Firefox ────────────────────────────────────────────────────────
function Backup-FirefoxProfile {
    param([string]$BackupPath, [string]$SourceUserFolder)
    Update-Progress "Firefox profiel backuppen..."

    $ffSrc = Join-Path $SourceUserFolder "AppData\Roaming\Mozilla\Firefox\Profiles"
    if (-not (Test-Path $ffSrc)) { Write-Log "Firefox profiel niet gevonden" -Level WARN; return }

    $zipDest = Join-Path $BackupPath "Firefox_Profiles.zip"
    try {
        Write-Log "Firefox profielen zippen..."
        Compress-WithProgress -SourceDir $ffSrc -ZipPath $zipDest
        $size = [Math]::Round((Get-Item $zipDest).Length / 1MB, 1)
        Write-Log "Firefox profielen gezipt: $zipDest ($size MB)" -Level OK
    } catch {
        Write-Log "Firefox backup fout: $_" -Level ERROR
    }
}

# ─── Migratie: Gebruikersmappen ───────────────────────────────────────────────
function Backup-UserFolders {
    param([string]$BackupPath, [string]$SourceUserFolder, [string[]]$Folders)

    foreach ($folder in $Folders) {
        $src = Join-Path $SourceUserFolder $folder
        if (-not (Test-Path $src)) {
            Write-Log "$folder niet gevonden, overgeslagen" -Level WARN
            continue
        }
        Update-Progress "$folder kopieren..."
        $dest = Join-Path $BackupPath "UserData\$folder"
        New-Item -ItemType Directory -Path $dest -Force | Out-Null

        # Detecteer OneDrive-redirectie (reparse point of lege map met cloud-only bestanden)
        $oneDriveRedirected = $false
        try {
            $dirInfo = New-Object System.IO.DirectoryInfo($src)
            if ($dirInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                $oneDriveRedirected = $true
            }
        } catch {}

        # Tel cloud-only OneDrive bestanden (hebben offline attribuut NIET gezet)
        $cloudOnlyCount = 0
        try {
            $allFiles = Get-ChildItem $src -Recurse -File -ErrorAction SilentlyContinue
            $cloudOnlyCount = ($allFiles | Where-Object {
                $_.Attributes -band [System.IO.FileAttributes]::Offline
            }).Count
        } catch {}

        try {
            & robocopy $src $dest /E /COPYALL /XA:ST /XJ /R:1 /W:1 /NP /MT:8 | Out-Null
            $exitCode = $LASTEXITCODE
            if ($exitCode -le 7) {
                $items = (Get-ChildItem $dest -Recurse -File -ErrorAction SilentlyContinue).Count
                if ($cloudOnlyCount -gt 0 -or ($items -eq 0 -and $oneDriveRedirected)) {
                    Write-Log "$folder gekopieerd ($items lokale bestanden) - WAARSCHUWING: $cloudOnlyCount cloud-only OneDrive bestanden NIET gekopieerd. Download eerst via OneDrive." -Level WARN
                } elseif ($items -eq 0) {
                    Write-Log "$folder gekopieerd (map leeg of alles in OneDrive cloud)" -Level WARN
                } else {
                    Write-Log "$folder gekopieerd ($items bestanden)" -Level OK
                }
            } else {
                Write-Log "$folder gedeeltelijk gekopieerd (robocopy exit: $exitCode)" -Level WARN
            }
        } catch {
            Write-Log "Fout bij kopieren $folder`: $_" -Level ERROR
        }
    }
}

# ─── Migratie: Outlook handtekeningen ────────────────────────────────────────
function Backup-OutlookSignatures {
    param([string]$BackupPath, [string]$SourceUserFolder)
    Update-Progress "Outlook handtekeningen backuppen..."

    $sigSrc = Join-Path $SourceUserFolder "AppData\Roaming\Microsoft\Signatures"
    if (-not (Test-Path $sigSrc)) { Write-Log "Outlook handtekeningen niet gevonden" -Level WARN; return }

    $dest = Join-Path $BackupPath "Outlook_Signatures"
    try {
        Copy-Item -Path $sigSrc -Destination $dest -Recurse -Force
        $count = (Get-ChildItem $dest -Recurse -File).Count
        Write-Log "Outlook handtekeningen gekopieerd ($count bestanden)" -Level OK
        Write-Log "Herstel: kopieer naar AppData\Roaming\Microsoft\Signatures\" -Level INFO
    } catch {
        Write-Log "Outlook handtekeningen fout: $_" -Level ERROR
    }
}

# ─── Migratie: Outlook PST ────────────────────────────────────────────────────
function Backup-OutlookPST {
    param([string]$BackupPath, [string]$SourceUserFolder)
    Update-Progress "Outlook PST bestanden zoeken..."

    $searchPaths = @(
        (Join-Path $SourceUserFolder "Documents\Outlook Files"),
        (Join-Path $SourceUserFolder "AppData\Local\Microsoft\Outlook"),
        (Join-Path $SourceUserFolder "Documents")
    )

    $pstFiles = @()
    foreach ($path in $searchPaths) {
        if (Test-Path $path) {
            $pstFiles += Get-ChildItem $path -Filter "*.pst" -Recurse -ErrorAction SilentlyContinue
        }
    }

    if ($pstFiles.Count -eq 0) {
        Write-Log "Geen Outlook PST bestanden gevonden" -Level INFO
        return
    }

    $totalMB   = [Math]::Round(($pstFiles | Measure-Object -Property Length -Sum).Sum / 1MB, 0)
    $fileList  = ($pstFiles | ForEach-Object { "  - $($_.Name) ($([Math]::Round($_.Length/1MB,0)) MB)" }) -join "`n"
    $msg       = "Gevonden PST bestanden ($totalMB MB totaal):`n$fileList`n`nKopieer naar backup locatie?"

    $result = [System.Windows.Forms.MessageBox]::Show(
        $msg, "Outlook PST bestanden",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $pstDest = Join-Path $BackupPath "Outlook_PST"
    New-Item -ItemType Directory -Path $pstDest -Force | Out-Null
    foreach ($pst in $pstFiles) {
        try {
            Write-Log "Kopieren PST: $($pst.Name) ($([Math]::Round($pst.Length/1MB,0)) MB)..."
            Copy-Item $pst.FullName -Destination $pstDest -Force
            Write-Log "PST gekopieerd: $($pst.Name)" -Level OK
        } catch {
            Write-Log "PST fout ($($pst.Name)): $_" -Level ERROR
        }
    }
}

# ─── Migratie: Mapped drives ──────────────────────────────────────────────────
function Backup-MappedDrives {
    param([string]$BackupPath, [string]$SourceComputer)
    Update-Progress "Mapped drives exporteren..."

    try {
        $wmiArgs = @{ Class = 'Win32_MappedLogicalDisk'; ComputerName = $SourceComputer; ErrorAction = 'Stop' }
        if ($script:SourceCred -and $SourceComputer -ne $env:COMPUTERNAME) {
            $wmiArgs['Credential'] = $script:SourceCred
        }
        $drives = Get-WmiObject @wmiArgs | Select-Object Name, ProviderName

        $drives | Export-Csv "$BackupPath\mapped_drives.csv" -NoTypeInformation -Encoding UTF8

        $restoreScript = @("# Mapped Drives herstel script", "# Gegenereerd: $(Get-Date)", "# Voer uit als de gebruiker (NIET als admin)", "", '$ErrorActionPreference = "Continue"', "")
        foreach ($d in $drives) {
            $letter = $d.Name.TrimEnd(':')
            $path   = $d.ProviderName
            $restoreScript += "net use ${letter}: `"$path`" /persistent:yes"
        }
        $restoreScript | Out-File "$BackupPath\Restore-MappedDrives.ps1" -Encoding UTF8

        Write-Log "Mapped drives geexporteerd ($($drives.Count) drives)" -Level OK
    } catch {
        Write-Log "Fout bij mapped drives export: $_" -Level ERROR
    }
}

# ─── Migratie: WiFi profielen ─────────────────────────────────────────────────
function Backup-WiFiProfiles {
    param([string]$BackupPath, [string]$SourceComputer)
    Update-Progress "WiFi profielen exporteren..."

    if ($SourceComputer -ne $env:COMPUTERNAME) {
        Write-Log "WiFi export alleen mogelijk op lokale computer (overgeslagen)" -Level WARN
        return
    }

    $wifiDir = Join-Path $BackupPath "WiFi_Profiles"
    New-Item -ItemType Directory -Path $wifiDir -Force | Out-Null

    try {
        $profileOutput = netsh wlan show profiles
        $profiles = $profileOutput |
            Select-String "(All User Profile|Alle gebruikersprofiel)\s*:\s*(.+)" |
            ForEach-Object { $_.Matches[0].Groups[2].Value.Trim() }

        $count = 0
        foreach ($wlanProfile in $profiles) {
            netsh wlan export profile name="$wlanProfile" folder="$wifiDir" key=clear 2>&1 | Out-Null
            $count++
        }

        # Herstel script
        $restoreScript = @(
            "# WiFi herstel script",
            "# Voer uit als administrator",
            "# Gegenereerd: $(Get-Date)",
            ""
        )
        Get-ChildItem $wifiDir -Filter "*.xml" | ForEach-Object {
            $restoreScript += "netsh wlan add profile filename=`"$($_.FullName)`" user=all"
        }
        $restoreScript | Out-File "$BackupPath\Restore-WiFiProfiles.ps1" -Encoding UTF8

        Write-Log "WiFi profielen geexporteerd ($count profielen)" -Level OK
    } catch {
        Write-Log "WiFi export fout: $_" -Level ERROR
    }
}

# ─── Migratie: Sticky Notes ───────────────────────────────────────────────────
function Backup-StickyNotes {
    param([string]$BackupPath, [string]$SourceUserFolder)
    Update-Progress "Sticky Notes backuppen..."

    $candidates = @(
        (Join-Path $SourceUserFolder "AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState"),
        (Join-Path $SourceUserFolder "AppData\Roaming\Microsoft\Sticky Notes")
    )

    foreach ($path in $candidates) {
        if (Test-Path $path) {
            $dest = Join-Path $BackupPath "StickyNotes"
            try {
                Copy-Item -Path $path -Destination $dest -Recurse -Force
                Write-Log "Sticky Notes gekopieerd" -Level OK
                Write-Log "Herstel: kopieer naar corresponderende AppData locatie" -Level INFO
                return
            } catch {
                Write-Log "Sticky Notes fout: $_" -Level ERROR
            }
        }
    }
    Write-Log "Sticky Notes niet gevonden" -Level WARN
}

# ─── Migratie: Achtergrond + Thema ───────────────────────────────────────────
function Backup-Wallpaper {
    param([string]$BackupPath, [string]$SourceUserFolder, [string]$SourceComputer)
    Update-Progress "Achtergrond en thema backuppen..."

    try {
        if ($SourceComputer -eq $env:COMPUTERNAME) {
            $wallpaperPath = (Get-ItemProperty "HKCU:\Control Panel\Desktop" -Name Wallpaper -ErrorAction SilentlyContinue).Wallpaper
            if ($wallpaperPath -and (Test-Path $wallpaperPath)) {
                $ext = [System.IO.Path]::GetExtension($wallpaperPath)
                Copy-Item $wallpaperPath "$BackupPath\wallpaper$ext" -Force
                Write-Log "Achtergrond gekopieerd: $([System.IO.Path]::GetFileName($wallpaperPath))" -Level OK
            }
        }

        # Thema map
        $themesSrc = Join-Path $SourceUserFolder "AppData\Roaming\Microsoft\Windows\Themes"
        if (Test-Path $themesSrc) {
            Copy-Item $themesSrc -Destination "$BackupPath\Themes" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Thema map gekopieerd" -Level OK
        }
    } catch {
        Write-Log "Achtergrond/thema fout: $_" -Level ERROR
    }
}

# ─── Migratie: Omgevingsvariabelen ───────────────────────────────────────────
function Export-UserEnvVars {
    param([string]$BackupPath, [string]$SourceComputer)
    Update-Progress "Omgevingsvariabelen exporteren..."

    if ($SourceComputer -ne $env:COMPUTERNAME) {
        Write-Log "Omgevingsvariabelen export alleen mogelijk lokaal (overgeslagen)" -Level WARN
        return
    }

    try {
        $envVars = [System.Environment]::GetEnvironmentVariables('User')
        $lines   = @(
            "# User omgevingsvariabelen herstel script",
            "# Gegenereerd: $(Get-Date)",
            "# Voer uit als de gebruiker",
            ""
        )

        foreach ($key in ($envVars.Keys | Sort-Object)) {
            $val = $envVars[$key] -replace "'", "''"
            if ($key -eq 'PATH') {
                $lines += "# PATH entries (voeg toe indien ontbrekend):"
                $envVars[$key].Split(';') | Where-Object { $_ } | ForEach-Object {
                    $lines += "#   $_"
                }
            } else {
                $lines += "[System.Environment]::SetEnvironmentVariable('$key', '$val', 'User')"
            }
        }

        $lines | Out-File "$BackupPath\Restore-UserEnvVars.ps1" -Encoding UTF8
        Write-Log "Omgevingsvariabelen geexporteerd ($($envVars.Count) variabelen)" -Level OK
    } catch {
        Write-Log "Omgevingsvariabelen fout: $_" -Level ERROR
    }
}

# ─── Rapport genereren ────────────────────────────────────────────────────────
function New-MigrationReport {
    param([string]$BackupPath)
    $reportPath = Join-Path $BackupPath "MIGRATIE_RAPPORT.txt"

    $errors   = ($script:LogEntries | Where-Object { $_ -match '\[ERROR\]' }).Count
    $warnings = ($script:LogEntries | Where-Object { $_ -match '\[WARN\]' }).Count

    $content = @"
========================================================
  MIGRATIE RAPPORT
  Gegenereerd : $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')
  Backup map  : $BackupPath
  Fouten      : $errors
  Waarschuwingen : $warnings
========================================================

LOGBOEK:
$($script:LogEntries -join "`n")

========================================================
HERSTEL INSTRUCTIES NIEUWE COMPUTER:

1. PRINTERS
   Als printers.printerExport aanwezig:
     printbrm.exe -R -F "$BackupPath\printers.printerExport"
   Als printer_list.csv aanwezig:
     Voer Restore-Printers.ps1 uit als administrator
     OF voeg printers manueel toe via Instellingen > Bluetooth en apparaten > Printers

2. CHROME
   - Zorg dat Chrome GESLOTEN is
   - Pak Chrome_UserData.zip uit naar:
     %LOCALAPPDATA%\Google\Chrome\User Data\
   - Overschrijf de bestaande map (of verwijder die eerst)

3. EDGE
   - Zorg dat Edge GESLOTEN is
   - Pak Edge_UserData.zip uit naar:
     %LOCALAPPDATA%\Microsoft\Edge\User Data\

4. FIREFOX
   - Zorg dat Firefox GESLOTEN is
   - Pak Firefox_Profiles.zip uit naar:
     %APPDATA%\Mozilla\Firefox\Profiles\

5. GEBRUIKERSBESTANDEN
   - Kopieer inhoud van UserData\ naar C:\Users\[gebruiker]\
   - Desktop, Documents, Downloads, Pictures, Music, Videos

6. OUTLOOK HANDTEKENINGEN
   - Kopieer Outlook_Signatures\ naar:
     %APPDATA%\Microsoft\Signatures\

7. OUTLOOK PST
   - Kopieer *.pst bestanden uit Outlook_PST\ naar gewenste locatie
   - Open Outlook > Bestand > Account Instellingen > Gegevensbestanden > Toevoegen

8. MAPPED DRIVES
   - Voer Restore-MappedDrives.ps1 uit als de GEBRUIKER (niet als admin)
   - Of voeg netwerk locaties manueel toe via Verkenner

9. WIFI PROFIELEN
   - Voer als administrator uit:
     netsh wlan add profile filename="[bestand].xml" user=all
   - Of voer Restore-WiFiProfiles.ps1 uit als administrator

10. STICKY NOTES
    - Zorg dat Sticky Notes GESLOTEN is
    - Kopieer StickyNotes\ inhoud naar corresponderende AppData locatie:
      %LOCALAPPDATA%\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\

11. ACHTERGROND
    - Kopieer wallpaper.* naar gewenste locatie
    - Rechtermuisknop > Als bureaubladachtergrond instellen

12. OMGEVINGSVARIABELEN
    - Voer Restore-UserEnvVars.ps1 uit als de GEBRUIKER

========================================================
"@

    $content | Out-File $reportPath -Encoding UTF8
    Write-Log "Migratierapport opgeslagen: $reportPath" -Level OK
    return $reportPath
}

# ─── Helper: unzip zonder UI blokkering ──────────────────────────────────────
function Expand-WithProgress {
    param([string]$ZipPath, [string]$DestDir)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $done = 0
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName.EndsWith('/') -or $entry.FullName.EndsWith('\')) { continue }
            $destPath  = Join-Path $DestDir $entry.FullName
            $destParent = [System.IO.Path]::GetDirectoryName($destPath)
            if (-not (Test-Path $destParent)) {
                New-Item -ItemType Directory -Path $destParent -Force | Out-Null
            }
            try {
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $destPath, $true)
            } catch { }
            $done++
            if ($done % 25 -eq 0) { [System.Windows.Forms.Application]::DoEvents() }
        }
    } finally {
        $zip.Dispose()
    }
}

# ─── Backup metadata opslaan / lezen ─────────────────────────────────────────
function Save-BackupMetadata {
    param([string]$BackupPath, [string]$SourceComputer, [string]$SourceUser,
          [string]$DestType, [string]$NetworkBase)
    @{
        Version        = '1.0'
        Date           = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        SourceComputer = $SourceComputer
        SourceUser     = $SourceUser
        BackupPath     = $BackupPath
        DestType       = $DestType      # 'network' of 'local'
        NetworkBase    = $NetworkBase
    } | ConvertTo-Json | Out-File "$BackupPath\backup_metadata.json" -Encoding UTF8
}

function Read-BackupMetadata {
    param([string]$BackupPath)
    $f = Join-Path $BackupPath "backup_metadata.json"
    if (Test-Path $f) {
        return (Get-Content $f -Raw | ConvertFrom-Json)
    }
    return $null
}

# ─── Restore: Printers ───────────────────────────────────────────────────────
function Restore-Printers {
    param([string]$BackupPath)
    Update-Progress "Printers herstellen..."
    $exportFile = Join-Path $BackupPath "printers.printerExport"
    $csvFile    = Join-Path $BackupPath "printer_list.csv"
    $printbrm   = @("$env:SystemRoot\System32\spool\tools\PrintBrm.exe",
                    "$env:SystemRoot\SysWOW64\spool\tools\PrintBrm.exe") |
                  Where-Object { Test-Path $_ } | Select-Object -First 1
    if ((Test-Path $exportFile) -and $printbrm) {
        & $printbrm -R -F $exportFile 2>&1 | ForEach-Object { Write-Log $_ }
        Write-Log "Printers hersteld via PrintBrm" -Level OK
    } elseif (Test-Path $csvFile) {
        Write-Log "printer_list.csv gevonden - voeg printers manueel toe via Instellingen" -Level WARN
        Start-Process explorer.exe $BackupPath
    } else {
        Write-Log "Geen printer backup gevonden" -Level WARN
    }
}

# ─── Restore: Browser profiel ────────────────────────────────────────────────
function Restore-BrowserProfile {
    param([string]$BackupPath, [string]$Browser)
    $cfg = switch ($Browser) {
        'Chrome'  { @{ Zip='Chrome_UserData.zip';  Dest="$env:LOCALAPPDATA\Google\Chrome\User Data";   Proc='chrome'  } }
        'Edge'    { @{ Zip='Edge_UserData.zip';    Dest="$env:LOCALAPPDATA\Microsoft\Edge\User Data"; Proc='msedge'  } }
        'Firefox' { @{ Zip='Firefox_Profiles.zip'; Dest="$env:APPDATA\Mozilla\Firefox\Profiles";      Proc='firefox' } }
    }
    Update-Progress "$Browser profiel herstellen..."
    $zipPath = Join-Path $BackupPath $cfg.Zip
    if (-not (Test-Path $zipPath)) { Write-Log "$Browser backup niet gevonden ($($cfg.Zip))" -Level WARN; return }
    if (Get-Process -Name $cfg.Proc -ErrorAction SilentlyContinue) {
        Write-Log "$Browser staat nog open - sluit $Browser volledig af" -Level ERROR
        [System.Windows.Forms.MessageBox]::Show(
            "Sluit $Browser af voor het herstellen van het profiel.",
            "$Browser is open", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    New-Item -ItemType Directory -Path $cfg.Dest -Force | Out-Null
    Write-Log "$Browser profiel uitpakken naar $($cfg.Dest)..."
    Expand-WithProgress -ZipPath $zipPath -DestDir $cfg.Dest
    Write-Log "$Browser profiel hersteld" -Level OK
}

# ─── Restore: Gebruikersmappen ───────────────────────────────────────────────
function Restore-UserFolders {
    param([string]$BackupPath, [string[]]$Folders)
    $userDataPath = Join-Path $BackupPath "UserData"
    foreach ($folder in $Folders) {
        $src = Join-Path $userDataPath $folder
        if (-not (Test-Path $src)) { Write-Log "$folder backup niet gevonden" -Level WARN; continue }
        Update-Progress "$folder herstellen..."
        $dest = Join-Path $env:USERPROFILE $folder
        New-Item -ItemType Directory -Path $dest -Force | Out-Null
        & robocopy $src $dest /E /XA:ST /XJ /R:1 /W:1 /NP /MT:8 | Out-Null
        $exitCode = $LASTEXITCODE
        if ($exitCode -le 7) {
            $items = (Get-ChildItem $dest -Recurse -File -ErrorAction SilentlyContinue).Count
            Write-Log "$folder hersteld ($items bestanden)" -Level OK
        } else {
            Write-Log "$folder gedeeltelijk hersteld (robocopy exit: $exitCode)" -Level WARN
        }
    }
}

# ─── Restore: Outlook handtekeningen ─────────────────────────────────────────
function Restore-OutlookSignatures {
    param([string]$BackupPath)
    Update-Progress "Outlook handtekeningen herstellen..."
    $src = Join-Path $BackupPath "Outlook_Signatures"
    if (-not (Test-Path $src)) { Write-Log "Geen Outlook handtekeningen backup" -Level WARN; return }
    $dest = "$env:APPDATA\Microsoft\Signatures"
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    Copy-Item "$src\*" $dest -Recurse -Force
    Write-Log "Outlook handtekeningen hersteld" -Level OK
}

# ─── Restore: Outlook PST ────────────────────────────────────────────────────
function Restore-OutlookPST {
    param([string]$BackupPath)
    Update-Progress "Outlook PST herstellen..."
    $pstSrc = Join-Path $BackupPath "Outlook_PST"
    if (-not (Test-Path $pstSrc)) { Write-Log "Geen Outlook PST backup" -Level INFO; return }
    $pstFiles = Get-ChildItem $pstSrc -Filter "*.pst"
    if ($pstFiles.Count -eq 0) { Write-Log "Geen PST bestanden in backup" -Level INFO; return }
    $destDir = "$env:USERPROFILE\Documents\Outlook Files"
    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    foreach ($pst in $pstFiles) {
        Write-Log "PST kopieren: $($pst.Name)..."
        Copy-Item $pst.FullName $destDir -Force
        Write-Log "PST gekopieerd: $($pst.Name)" -Level OK
    }
    [System.Windows.Forms.MessageBox]::Show(
        "PST bestanden gekopieerd naar:`n$destDir`n`nVoeg toe in Outlook via:`nBestand > Account Instellingen > Gegevensbestanden > Toevoegen",
        "Outlook PST - handmatige stap",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

# ─── Restore: Mapped drives ──────────────────────────────────────────────────
function Restore-MappedDrives {
    param([string]$BackupPath)
    Update-Progress "Mapped drives herstellen..."
    $script = Join-Path $BackupPath "Restore-MappedDrives.ps1"
    if (-not (Test-Path $script)) { Write-Log "Geen mapped drives backup" -Level WARN; return }
    try {
        & $script 2>&1 | ForEach-Object { Write-Log $_ }
        Write-Log "Mapped drives hersteld" -Level OK
    } catch { Write-Log "Mapped drives fout: $_" -Level ERROR }
}

# ─── Restore: WiFi profielen ─────────────────────────────────────────────────
function Restore-WiFiProfiles {
    param([string]$BackupPath)
    Update-Progress "WiFi profielen herstellen..."
    $wifiDir = Join-Path $BackupPath "WiFi_Profiles"
    if (-not (Test-Path $wifiDir)) { Write-Log "Geen WiFi backup" -Level WARN; return }
    $xmlFiles = Get-ChildItem $wifiDir -Filter "*.xml"
    $count = 0
    foreach ($xml in $xmlFiles) {
        netsh wlan add profile filename="$($xml.FullName)" user=all 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { $count++ }
    }
    Write-Log "WiFi profielen hersteld ($count van $($xmlFiles.Count))" -Level OK
}

# ─── Restore: Sticky Notes ───────────────────────────────────────────────────
function Restore-StickyNotes {
    param([string]$BackupPath)
    Update-Progress "Sticky Notes herstellen..."
    $src = Join-Path $BackupPath "StickyNotes"
    if (-not (Test-Path $src)) { Write-Log "Geen Sticky Notes backup" -Level WARN; return }
    $dest = "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState"
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    Copy-Item "$src\*" $dest -Recurse -Force
    Write-Log "Sticky Notes hersteld" -Level OK
}

# ─── Restore: Achtergrond ────────────────────────────────────────────────────
function Restore-Wallpaper {
    param([string]$BackupPath)
    Update-Progress "Achtergrond herstellen..."
    $wallpaper = Get-ChildItem $BackupPath -Filter "wallpaper.*" -ErrorAction SilentlyContinue |
                 Select-Object -First 1
    if (-not $wallpaper) { Write-Log "Geen achtergrond backup" -Level WARN; return }
    $destPath = "$env:APPDATA\Microsoft\Windows\Themes\$($wallpaper.Name)"
    Copy-Item $wallpaper.FullName $destPath -Force
    Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public class WallpaperHelper {
    [DllImport("user32.dll",CharSet=CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction,int uParam,string lpvParam,int fuWinIni);
}
'@
    [WallpaperHelper]::SystemParametersInfo(20, 0, $destPath, 3) | Out-Null
    Write-Log "Achtergrond hersteld: $($wallpaper.Name)" -Level OK
}

# ─── Restore: Omgevingsvariabelen ────────────────────────────────────────────
function Restore-EnvVars {
    param([string]$BackupPath)
    Update-Progress "Omgevingsvariabelen herstellen..."
    $script = Join-Path $BackupPath "Restore-UserEnvVars.ps1"
    if (-not (Test-Path $script)) { Write-Log "Geen omgevingsvariabelen backup" -Level WARN; return }
    try { & $script; Write-Log "Omgevingsvariabelen hersteld" -Level OK }
    catch { Write-Log "Omgevingsvariabelen fout: $_" -Level ERROR }
}

# ══════════════════════════════════════════════════════════════════════════════
#  GUI opbouw
# ══════════════════════════════════════════════════════════════════════════════

$form                  = New-Object System.Windows.Forms.Form
$form.Text             = "Gebruikersprofiel Migratie Tool"
$form.Size             = New-Object System.Drawing.Size(760, 1000)
$form.StartPosition    = "CenterScreen"
$form.FormBorderStyle  = "FixedDialog"
$form.MaximizeBox      = $false
$form.Font             = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor        = [System.Drawing.Color]::FromArgb(240, 242, 245)

# Titel
$lblTitle              = New-Object System.Windows.Forms.Label
$lblTitle.Text         = "Gebruikersprofiel Migratie Tool"
$lblTitle.Font         = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor    = [System.Drawing.Color]::FromArgb(0, 99, 177)
$lblTitle.Location     = New-Object System.Drawing.Point(15, 12)
$lblTitle.Size         = New-Object System.Drawing.Size(720, 30)
$form.Controls.Add($lblTitle)

$lblSubtitle           = New-Object System.Windows.Forms.Label
$lblSubtitle.Text      = "Migreert gebruikersprofiel, browsers, bestanden en instellingen van oude naar nieuwe computer"
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$lblSubtitle.Location  = New-Object System.Drawing.Point(15, 44)
$lblSubtitle.Size      = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblSubtitle)

# ─── Modus selector ───────────────────────────────────────────────────────────
$grpMode               = New-Object System.Windows.Forms.GroupBox
$grpMode.Text          = "Modus"
$grpMode.Location      = New-Object System.Drawing.Point(10, 68)
$grpMode.Size          = New-Object System.Drawing.Size(724, 48)
$form.Controls.Add($grpMode)

$rbBackup              = New-Object System.Windows.Forms.RadioButton
$rbBackup.Text         = "Backup  (kopieer van oude computer)"
$rbBackup.Location     = New-Object System.Drawing.Point(12, 18)
$rbBackup.Size         = New-Object System.Drawing.Size(290, 22)
$rbBackup.Checked      = $true
$grpMode.Controls.Add($rbBackup)

$rbRestore             = New-Object System.Windows.Forms.RadioButton
$rbRestore.Text        = "Herstel  (zet terug op nieuwe computer)"
$rbRestore.Location    = New-Object System.Drawing.Point(320, 18)
$rbRestore.Size        = New-Object System.Drawing.Size(290, 22)
$grpMode.Controls.Add($rbRestore)

# ─── Bron sectie ──────────────────────────────────────────────────────────────
$grpSource             = New-Object System.Windows.Forms.GroupBox
$grpSource.Text        = "Bron (Oude Computer)"
$grpSource.Location    = New-Object System.Drawing.Point(10, 125)
$grpSource.Size        = New-Object System.Drawing.Size(724, 110)
$form.Controls.Add($grpSource)

$lblComputer           = New-Object System.Windows.Forms.Label
$lblComputer.Text      = "Computer naam:"
$lblComputer.Location  = New-Object System.Drawing.Point(10, 26)
$lblComputer.Size      = New-Object System.Drawing.Size(120, 20)
$grpSource.Controls.Add($lblComputer)

$txtComputer           = New-Object System.Windows.Forms.TextBox
$txtComputer.Text      = $env:COMPUTERNAME
$txtComputer.Location  = New-Object System.Drawing.Point(135, 23)
$txtComputer.Size      = New-Object System.Drawing.Size(200, 23)
$grpSource.Controls.Add($txtComputer)

$btnLocalPC            = New-Object System.Windows.Forms.Button
$btnLocalPC.Text       = "Deze PC"
$btnLocalPC.Location   = New-Object System.Drawing.Point(345, 21)
$btnLocalPC.Size       = New-Object System.Drawing.Size(75, 26)
$btnLocalPC.Add_Click({ $txtComputer.Text = $env:COMPUTERNAME })
$grpSource.Controls.Add($btnLocalPC)

$lblCredHint           = New-Object System.Windows.Forms.Label
$lblCredHint.Text      = "Remote PC vereist admin credentials voor C$ share"
$lblCredHint.ForeColor = [System.Drawing.Color]::DimGray
$lblCredHint.Location  = New-Object System.Drawing.Point(430, 26)
$lblCredHint.Size      = New-Object System.Drawing.Size(280, 20)
$grpSource.Controls.Add($lblCredHint)

$lblUser               = New-Object System.Windows.Forms.Label
$lblUser.Text          = "Gebruikersnaam:"
$lblUser.Location      = New-Object System.Drawing.Point(10, 56)
$lblUser.Size          = New-Object System.Drawing.Size(120, 20)
$grpSource.Controls.Add($lblUser)

$txtUser               = New-Object System.Windows.Forms.TextBox
$txtUser.Text          = [System.IO.Path]::GetFileName($env:USERPROFILE)
$txtUser.Location      = New-Object System.Drawing.Point(135, 53)
$txtUser.Size          = New-Object System.Drawing.Size(200, 23)
$grpSource.Controls.Add($txtUser)

$btnGetCred            = New-Object System.Windows.Forms.Button
$btnGetCred.Text       = "Credentials instellen"
$btnGetCred.Location   = New-Object System.Drawing.Point(135, 80)
$btnGetCred.Size       = New-Object System.Drawing.Size(160, 26)
$btnGetCred.Add_Click({
    $cred = Get-Credential -Message "Admin credentials voor \\$($txtComputer.Text)\c$"
    if ($cred) {
        $script:SourceCred = $cred
        $btnGetCred.Text       = "OK: $($cred.UserName)"
        $btnGetCred.BackColor  = [System.Drawing.Color]::FromArgb(200, 240, 200)
    }
})
$grpSource.Controls.Add($btnGetCred)

$btnTestConn           = New-Object System.Windows.Forms.Button
$btnTestConn.Text      = "Test verbinding"
$btnTestConn.Location  = New-Object System.Drawing.Point(305, 80)
$btnTestConn.Size      = New-Object System.Drawing.Size(120, 26)
$btnTestConn.Add_Click({
    $comp = $txtComputer.Text.Trim()
    if ($comp -eq $env:COMPUTERNAME) {
        [System.Windows.Forms.MessageBox]::Show("Lokale computer - geen netwerkverbinding nodig.","Test") | Out-Null
        return
    }
    if (Test-Connection -ComputerName $comp -Count 1 -Quiet) {
        $share = "\\$comp\c$"
        if (Test-Path $share) {
            [System.Windows.Forms.MessageBox]::Show("Verbinding OK: $share bereikbaar.","Test",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show("PC bereikbaar maar C`$ share niet toegankelijk.`nStel credentials in.","Test",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("$comp niet bereikbaar. Controleer naam/netwerk.","Test",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})
$grpSource.Controls.Add($btnTestConn)

# ─── Herstel bron sectie (zichtbaar in Herstel-modus) ────────────────────────
$grpRestoreSource          = New-Object System.Windows.Forms.GroupBox
$grpRestoreSource.Text     = "Backup bron"
$grpRestoreSource.Location = New-Object System.Drawing.Point(10, 125)
$grpRestoreSource.Size     = New-Object System.Drawing.Size(724, 110)
$grpRestoreSource.Visible  = $false
$form.Controls.Add($grpRestoreSource)

$rbRestoreNetwork          = New-Object System.Windows.Forms.RadioButton
$rbRestoreNetwork.Text     = "Netwerk share (backup stond op server):"
$rbRestoreNetwork.Location = New-Object System.Drawing.Point(10, 22)
$rbRestoreNetwork.Size     = New-Object System.Drawing.Size(250, 22)
$grpRestoreSource.Controls.Add($rbRestoreNetwork)

$txtRestoreNetwork         = New-Object System.Windows.Forms.TextBox
$txtRestoreNetwork.Text    = "\\server\migratie"
$txtRestoreNetwork.Location= New-Object System.Drawing.Point(268, 19)
$txtRestoreNetwork.Size    = New-Object System.Drawing.Size(230, 23)
$txtRestoreNetwork.Enabled = $false
$grpRestoreSource.Controls.Add($txtRestoreNetwork)

$btnScanBackups            = New-Object System.Windows.Forms.Button
$btnScanBackups.Text       = "Zoek backups"
$btnScanBackups.Location   = New-Object System.Drawing.Point(508, 18)
$btnScanBackups.Size       = New-Object System.Drawing.Size(100, 26)
$btnScanBackups.Enabled    = $false
$btnScanBackups.Add_Click({
    $basePath = $txtRestoreNetwork.Text.Trim()
    $cboBackupFolders.Items.Clear()
    $found = Get-ChildItem $basePath -Directory -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -match '^Migratie_' } |
             Sort-Object LastWriteTime -Descending
    foreach ($f in $found) { $cboBackupFolders.Items.Add($f.FullName) | Out-Null }
    if ($cboBackupFolders.Items.Count -gt 0) {
        $cboBackupFolders.SelectedIndex = 0
        Write-Log "$($cboBackupFolders.Items.Count) backup(s) gevonden op $basePath" -Level OK
    } else {
        Write-Log "Geen backups gevonden op $basePath" -Level WARN
    }
})
$grpRestoreSource.Controls.Add($btnScanBackups)

$cboBackupFolders          = New-Object System.Windows.Forms.ComboBox
$cboBackupFolders.Location = New-Object System.Drawing.Point(268, 49)
$cboBackupFolders.Size     = New-Object System.Drawing.Size(340, 23)
$cboBackupFolders.DropDownStyle = 'DropDownList'
$cboBackupFolders.Enabled  = $false
$grpRestoreSource.Controls.Add($cboBackupFolders)

$rbRestoreLocal            = New-Object System.Windows.Forms.RadioButton
$rbRestoreLocal.Text       = "Lokale map / USB:"
$rbRestoreLocal.Checked    = $true
$rbRestoreLocal.Location   = New-Object System.Drawing.Point(10, 79)
$rbRestoreLocal.Size       = New-Object System.Drawing.Size(150, 22)
$grpRestoreSource.Controls.Add($rbRestoreLocal)

$txtRestorePath            = New-Object System.Windows.Forms.TextBox
$txtRestorePath.Text       = "C:\Temp\Migratie\Migratie_..."
$txtRestorePath.Location   = New-Object System.Drawing.Point(168, 76)
$txtRestorePath.Size       = New-Object System.Drawing.Size(370, 23)
$grpRestoreSource.Controls.Add($txtRestorePath)

$btnBrowseRestore          = New-Object System.Windows.Forms.Button
$btnBrowseRestore.Text     = "..."
$btnBrowseRestore.Location = New-Object System.Drawing.Point(548, 75)
$btnBrowseRestore.Size     = New-Object System.Drawing.Size(35, 26)
$btnBrowseRestore.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Kies backup map (Migratie_... map)"
    if ($fbd.ShowDialog() -eq 'OK') {
        $txtRestorePath.Text = $fbd.SelectedPath
        # Lees metadata en vul info in log
        $meta = Read-BackupMetadata -BackupPath $fbd.SelectedPath
        if ($meta) {
            Write-Log "Backup van $($meta.SourceComputer) / $($meta.SourceUser) op $($meta.Date)" -Level OK
        }
    }
})
$grpRestoreSource.Controls.Add($btnBrowseRestore)

# Netwerk/lokaal toggle voor restore
$rbRestoreNetwork.Add_CheckedChanged({
    $isNet = $rbRestoreNetwork.Checked
    $txtRestoreNetwork.Enabled = $isNet
    $btnScanBackups.Enabled    = $isNet
    $cboBackupFolders.Enabled  = $isNet
    $txtRestorePath.Enabled    = -not $isNet
    $btnBrowseRestore.Enabled  = -not $isNet
})
$rbRestoreLocal.Add_CheckedChanged({
    $isNet = $rbRestoreNetwork.Checked
    $txtRestoreNetwork.Enabled = $isNet
    $btnScanBackups.Enabled    = $isNet
    $cboBackupFolders.Enabled  = $isNet
    $txtRestorePath.Enabled    = -not $isNet
    $btnBrowseRestore.Enabled  = -not $isNet
})

# ─── Te migreren items ────────────────────────────────────────────────────────
$grpItems              = New-Object System.Windows.Forms.GroupBox
$grpItems.Text         = "Te migreren items"
$grpItems.Location     = New-Object System.Drawing.Point(10, 243)
$grpItems.Size         = New-Object System.Drawing.Size(724, 295)
$form.Controls.Add($grpItems)

# Kolom 1: browsers + office (links)
# Kolom 2: bestanden + overige (rechts)
$checkDefs = @(
    # Name                 Text                                             Left  Top  Col
    @{ n='chkPrinters';    t='Printers (PrintBrm / WMI export)';           x=10; y=24 }
    @{ n='chkChrome';      t='Chrome profiel (+ wachtwoord check)';        x=10; y=48 }
    @{ n='chkEdge';        t='Edge profiel';                               x=10; y=72 }
    @{ n='chkFirefox';     t='Firefox profiel';                            x=10; y=96 }
    @{ n='chkOutlookSig';  t='Outlook handtekeningen';                     x=10; y=120 }
    @{ n='chkOutlookPST';  t='Outlook PST bestanden (bevestiging gevraagd)';x=10; y=144 }
    @{ n='chkMappedDrives';t='Mapped drives (herstelscript genereren)';    x=10; y=168 }
    @{ n='chkWiFi';        t='WiFi profielen (alleen lokaal)';             x=10; y=192 }
    @{ n='chkStickyNotes'; t='Sticky Notes';                               x=10; y=216 }
    @{ n='chkWallpaper';   t='Achtergrond + Thema';                        x=10; y=240 }
    @{ n='chkEnvVars';     t='Omgevingsvariabelen (user)';                 x=10; y=264 }

    @{ n='chkDesktop';     t='Bureaublad (Desktop)';                       x=380; y=24 }
    @{ n='chkDocuments';   t='Documenten (Documents)';                     x=380; y=48 }
    @{ n='chkDownloads';   t='Downloads';                                  x=380; y=72 }
    @{ n='chkPictures';    t='Afbeeldingen (Pictures)';                    x=380; y=96 }
    @{ n='chkMusic';       t='Muziek (Music)';                             x=380; y=120 }
    @{ n='chkVideos';      t='Videos';                                     x=380; y=144 }
)

$checkboxes = @{}
foreach ($cd in $checkDefs) {
    $cb           = New-Object System.Windows.Forms.CheckBox
    $cb.Text      = $cd.t
    $cb.Location  = New-Object System.Drawing.Point($cd.x, $cd.y)
    $cb.Size      = New-Object System.Drawing.Size(355, 22)
    $cb.Checked   = $true
    $grpItems.Controls.Add($cb)
    $checkboxes[$cd.n] = $cb
}

$btnAll  = New-Object System.Windows.Forms.Button
$btnAll.Text     = "Alles selecteren"
$btnAll.Location = New-Object System.Drawing.Point(10, 265)
$btnAll.Size     = New-Object System.Drawing.Size(130, 26)
$btnAll.Add_Click({ $checkboxes.Values | ForEach-Object { $_.Checked = $true } })
$grpItems.Controls.Add($btnAll)

$btnNone = New-Object System.Windows.Forms.Button
$btnNone.Text     = "Niets selecteren"
$btnNone.Location = New-Object System.Drawing.Point(150, 265)
$btnNone.Size     = New-Object System.Drawing.Size(130, 26)
$btnNone.Add_Click({ $checkboxes.Values | ForEach-Object { $_.Checked = $false } })
$grpItems.Controls.Add($btnNone)

# ─── Bestemming ───────────────────────────────────────────────────────────────
$grpDest           = New-Object System.Windows.Forms.GroupBox
$grpDest.Text      = "Backup bestemming"
$grpDest.Location  = New-Object System.Drawing.Point(10, 546)
$grpDest.Size      = New-Object System.Drawing.Size(724, 112)
$form.Controls.Add($grpDest)

$rbNetwork         = New-Object System.Windows.Forms.RadioButton
$rbNetwork.Text    = "Netwerk share:"
$rbNetwork.Location= New-Object System.Drawing.Point(10, 22)
$rbNetwork.Size    = New-Object System.Drawing.Size(120, 22)
$grpDest.Controls.Add($rbNetwork)

$txtNetworkPath        = New-Object System.Windows.Forms.TextBox
$txtNetworkPath.Text   = "\\server\migratie"
$txtNetworkPath.Location = New-Object System.Drawing.Point(140, 19)
$txtNetworkPath.Size   = New-Object System.Drawing.Size(370, 23)
$txtNetworkPath.Enabled= $false
$grpDest.Controls.Add($txtNetworkPath)

$rbLocal           = New-Object System.Windows.Forms.RadioButton
$rbLocal.Text      = "Lokale map:"
$rbLocal.Checked   = $true
$rbLocal.Location  = New-Object System.Drawing.Point(10, 52)
$rbLocal.Size      = New-Object System.Drawing.Size(120, 22)
$grpDest.Controls.Add($rbLocal)

$txtLocalPath        = New-Object System.Windows.Forms.TextBox
$txtLocalPath.Text   = "C:\Temp\Migratie"
$txtLocalPath.Location = New-Object System.Drawing.Point(140, 49)
$txtLocalPath.Size   = New-Object System.Drawing.Size(310, 23)
$grpDest.Controls.Add($txtLocalPath)

$btnBrowse         = New-Object System.Windows.Forms.Button
$btnBrowse.Text    = "..."
$btnBrowse.Location= New-Object System.Drawing.Point(460, 48)
$btnBrowse.Size    = New-Object System.Drawing.Size(35, 26)
$btnBrowse.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Kies backup bestemming"
    if ($fbd.ShowDialog() -eq 'OK') { $txtLocalPath.Text = $fbd.SelectedPath }
})
$grpDest.Controls.Add($btnBrowse)

$rbNetwork.Add_CheckedChanged({ $txtNetworkPath.Enabled = $rbNetwork.Checked })
$rbLocal.Add_CheckedChanged({ $txtNetworkPath.Enabled = $rbNetwork.Checked })

# ─── Modus toggle handlers ────────────────────────────────────────────────────
$rbBackup.Add_CheckedChanged({
    if ($rbBackup.Checked) {
        $grpSource.Visible        = $true
        $grpDest.Visible          = $true
        $grpRestoreSource.Visible = $false
        $grpItems.Text            = "Te migreren items"
        $btnStart.Text            = "Start Backup"
    }
})
$rbRestore.Add_CheckedChanged({
    if ($rbRestore.Checked) {
        $grpSource.Visible        = $false
        $grpDest.Visible          = $false
        $grpRestoreSource.Visible = $true
        $grpItems.Text            = "Te herstellen items"
        $btnStart.Text            = "Start Herstel"
    }
})

$chkZip            = New-Object System.Windows.Forms.CheckBox
$chkZip.Text       = "Alles zippen na backup (handig voor transport via USB)"
$chkZip.Location   = New-Object System.Drawing.Point(10, 80)
$chkZip.Size       = New-Object System.Drawing.Size(400, 22)
$chkZip.Checked    = $false
$grpDest.Controls.Add($chkZip)

# ─── Status + voortgang ───────────────────────────────────────────────────────
$script:StatusLabel        = New-Object System.Windows.Forms.Label
$script:StatusLabel.Text   = "Klaar om te starten..."
$script:StatusLabel.Location = New-Object System.Drawing.Point(10, 665)
$script:StatusLabel.Size   = New-Object System.Drawing.Size(724, 20)
$form.Controls.Add($script:StatusLabel)

$script:ProgressBar        = New-Object System.Windows.Forms.ProgressBar
$script:ProgressBar.Location = New-Object System.Drawing.Point(10, 687)
$script:ProgressBar.Size   = New-Object System.Drawing.Size(724, 22)
$script:ProgressBar.Style  = "Continuous"
$form.Controls.Add($script:ProgressBar)

$script:LogBox             = New-Object System.Windows.Forms.RichTextBox
$script:LogBox.Location    = New-Object System.Drawing.Point(10, 715)
$script:LogBox.Size        = New-Object System.Drawing.Size(724, 182)
$script:LogBox.ReadOnly    = $true
$script:LogBox.BackColor   = [System.Drawing.Color]::FromArgb(18, 18, 18)
$script:LogBox.ForeColor   = [System.Drawing.Color]::LightGray
$script:LogBox.Font        = New-Object System.Drawing.Font("Consolas", 8)
$script:LogBox.ScrollBars  = "Vertical"
$form.Controls.Add($script:LogBox)

# ─── Knoppen ──────────────────────────────────────────────────────────────────
$btnStart             = New-Object System.Windows.Forms.Button
$btnStart.Text        = "Start Backup"
$btnStart.Location    = New-Object System.Drawing.Point(10, 917)
$btnStart.Size        = New-Object System.Drawing.Size(150, 40)
$btnStart.BackColor   = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnStart.ForeColor   = [System.Drawing.Color]::White
$btnStart.Font        = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnStart.FlatStyle   = "Flat"
$form.Controls.Add($btnStart)

$btnOpenFolder        = New-Object System.Windows.Forms.Button
$btnOpenFolder.Text   = "Open backup map"
$btnOpenFolder.Location = New-Object System.Drawing.Point(170, 917)
$btnOpenFolder.Size   = New-Object System.Drawing.Size(140, 40)
$btnOpenFolder.Enabled= $false
$form.Controls.Add($btnOpenFolder)

$btnClearLog          = New-Object System.Windows.Forms.Button
$btnClearLog.Text     = "Log wissen"
$btnClearLog.Location = New-Object System.Drawing.Point(320, 917)
$btnClearLog.Size     = New-Object System.Drawing.Size(100, 40)
$btnClearLog.Add_Click({ $script:LogBox.Clear(); $script:LogEntries.Clear() })
$form.Controls.Add($btnClearLog)

$btnClose             = New-Object System.Windows.Forms.Button
$btnClose.Text        = "Sluiten"
$btnClose.Location    = New-Object System.Drawing.Point(584, 917)
$btnClose.Size        = New-Object System.Drawing.Size(150, 40)
$btnClose.Add_Click({ Disconnect-RemoteShare; $form.Close() })
$form.Controls.Add($btnClose)

# ─── Start migratie handler ───────────────────────────────────────────────────
$btnStart.Add_Click({
    $btnStart.Enabled        = $false
    $btnOpenFolder.Enabled   = $false
    $script:ProgressBar.Value = 0
    $script:LogEntries.Clear()
    $script:LogBox.Clear()

    # ── HERSTEL MODUS ─────────────────────────────────────────────────────────
    if ($rbRestore.Checked) {
        # Bepaal backup pad
        $backupPath = if ($rbRestoreNetwork.Checked) {
            if ($cboBackupFolders.SelectedItem) { $cboBackupFolders.SelectedItem.ToString() } else { $null }
        } else {
            $txtRestorePath.Text.Trim()
        }

        if (-not $backupPath -or -not (Test-Path $backupPath)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Backup map niet gevonden:`n$backupPath`n`nKies een geldige backup map.",
                "Backup map niet gevonden",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            $btnStart.Enabled = $true
            return
        }

        # Lees metadata (informatief)
        $meta = Read-BackupMetadata -BackupPath $backupPath
        if ($meta) {
            Write-Log "Backup van $($meta.SourceComputer) / $($meta.SourceUser) ($($meta.Date))" -Level OK
        }
        Write-Log "Herstel gestart vanuit: $backupPath"

        # Pre-flight browser check (herstel mode)
        $browserDefs = @(
            @{ Name='Chrome';         Process='chrome';  Selected=$checkboxes['chkChrome'].Checked }
            @{ Name='Microsoft Edge'; Process='msedge';  Selected=$checkboxes['chkEdge'].Checked }
            @{ Name='Firefox';        Process='firefox'; Selected=$checkboxes['chkFirefox'].Checked }
        )
        $openBrowsers = $browserDefs | Where-Object { $_.Selected -and (Get-Process -Name $_.Process -ErrorAction SilentlyContinue) }
        if ($openBrowsers) {
            $lijst = ($openBrowsers | ForEach-Object { "  - $($_.Name)" }) -join "`n"
            $ans = [System.Windows.Forms.MessageBox]::Show(
                "Open browsers gevonden:`n$lijst`n`nAutomatisch sluiten?",
                "Open browsers", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($ans -eq [System.Windows.Forms.DialogResult]::No) { $btnStart.Enabled = $true; return }
            if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
                foreach ($b in $openBrowsers) {
                    Get-Process -Name $b.Process -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Milliseconds 600
                    [System.Windows.Forms.Application]::DoEvents()
                }
            }
        }

        $selectedChecks     = $checkboxes.Values | Where-Object { $_.Checked }
        $script:TotalSteps  = $selectedChecks.Count + 1
        $script:CurrentStep = 0

        try {
            if ($checkboxes['chkPrinters'].Checked)    { Restore-Printers          -BackupPath $backupPath }
            if ($checkboxes['chkChrome'].Checked)      { Restore-BrowserProfile    -BackupPath $backupPath -Browser 'Chrome' }
            if ($checkboxes['chkEdge'].Checked)        { Restore-BrowserProfile    -BackupPath $backupPath -Browser 'Edge' }
            if ($checkboxes['chkFirefox'].Checked)     { Restore-BrowserProfile    -BackupPath $backupPath -Browser 'Firefox' }
            if ($checkboxes['chkOutlookSig'].Checked)  { Restore-OutlookSignatures -BackupPath $backupPath }
            if ($checkboxes['chkOutlookPST'].Checked)  { Restore-OutlookPST        -BackupPath $backupPath }

            $folders = @()
            if ($checkboxes['chkDesktop'].Checked)   { $folders += 'Desktop' }
            if ($checkboxes['chkDocuments'].Checked) { $folders += 'Documents' }
            if ($checkboxes['chkDownloads'].Checked) { $folders += 'Downloads' }
            if ($checkboxes['chkPictures'].Checked)  { $folders += 'Pictures' }
            if ($checkboxes['chkMusic'].Checked)     { $folders += 'Music' }
            if ($checkboxes['chkVideos'].Checked)    { $folders += 'Videos' }
            if ($folders.Count -gt 0) { Restore-UserFolders -BackupPath $backupPath -Folders $folders }

            if ($checkboxes['chkMappedDrives'].Checked) { Restore-MappedDrives -BackupPath $backupPath }
            if ($checkboxes['chkWiFi'].Checked)          { Restore-WiFiProfiles -BackupPath $backupPath }
            if ($checkboxes['chkStickyNotes'].Checked)   { Restore-StickyNotes  -BackupPath $backupPath }
            if ($checkboxes['chkWallpaper'].Checked)     { Restore-Wallpaper    -BackupPath $backupPath }
            if ($checkboxes['chkEnvVars'].Checked)       { Restore-EnvVars      -BackupPath $backupPath }

            $script:ProgressBar.Value = 100
            $script:StatusLabel.Text  = "Herstel voltooid!"
            $errors   = ($script:LogEntries | Where-Object { $_ -match '\[ERROR\]' }).Count
            $warnings = ($script:LogEntries | Where-Object { $_ -match '\[WARN\]'  }).Count
            Write-Log "=== HERSTEL VOLTOOID (fouten: $errors, waarschuwingen: $warnings) ===" -Level OK
            [System.Windows.Forms.MessageBox]::Show(
                "Herstel voltooid!`n`nFouten: $errors`nWaarschuwingen: $warnings",
                "Herstel Voltooid",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } catch {
            Write-Log "Onverwachte fout bij herstel: $_" -Level ERROR
        } finally {
            $btnStart.Enabled = $true
        }
        return
    }

    # ── BACKUP MODUS ──────────────────────────────────────────────────────────
    $sourceComputer = $txtComputer.Text.Trim()
    $sourceUser     = $txtUser.Text.Trim()
    $isLocal        = ($sourceComputer -ieq $env:COMPUTERNAME)

    # ── Pre-flight: controleer open browsers ──────────────────────────────────
    $browserDefs = @(
        @{ Name = 'Chrome';          Process = 'chrome';   Selected = $checkboxes['chkChrome'].Checked }
        @{ Name = 'Microsoft Edge';  Process = 'msedge';   Selected = $checkboxes['chkEdge'].Checked }
        @{ Name = 'Firefox';         Process = 'firefox';  Selected = $checkboxes['chkFirefox'].Checked }
    )

    $openBrowsers = $browserDefs | Where-Object {
        (Get-Process -Name $_.Process -ErrorAction SilentlyContinue) -ne $null
    }

    if ($openBrowsers) {
        $lijst = ($openBrowsers | ForEach-Object { "  - $($_.Name)" }) -join "`n"
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "De volgende browsers zijn nog geopend:`n$lijst`n`n" +
            "Open browsers kunnen bestanden vergrendelen waardoor de backup mislukt.`n`n" +
            "Klik 'Ja' om ze automatisch te sluiten en door te gaan.`n" +
            "Klik 'Nee' om ze manueel te sluiten (migratie stopt).`n" +
            "Klik 'Annuleer' om toch door te gaan zonder sluiten (risico op fouten).",
            "Open browsers gedetecteerd",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($ans -eq [System.Windows.Forms.DialogResult]::No) {
            $btnStart.Enabled = $true
            return
        }

        if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
            foreach ($b in $openBrowsers) {
                Write-Log "$($b.Name) afsluiten..."
                Get-Process -Name $b.Process -ErrorAction SilentlyContinue |
                    Stop-Process -Force -ErrorAction SilentlyContinue
                Start-Sleep -Milliseconds 800
                [System.Windows.Forms.Application]::DoEvents()
                if (Get-Process -Name $b.Process -ErrorAction SilentlyContinue) {
                    Write-Log "$($b.Name) kon niet afgesloten worden" -Level WARN
                } else {
                    Write-Log "$($b.Name) afgesloten" -Level OK
                }
            }
        }
        # Annuleer = doorgaan zonder sluiten
    }

    # Bepaal backup map
    $timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupName = "Migratie_${sourceUser}_${timestamp}"

    if ($rbNetwork.Checked) {
        $script:BackupRoot = Join-Path $txtNetworkPath.Text.Trim() $backupName
    } else {
        $script:BackupRoot = Join-Path $txtLocalPath.Text.Trim() $backupName
    }

    Write-Log "Migratie gestart voor gebruiker: $sourceUser op $sourceComputer"
    Write-Log "Backup locatie: $($script:BackupRoot)"

    # Maak backup map aan
    try {
        New-Item -ItemType Directory -Path $script:BackupRoot -Force | Out-Null
    } catch {
        Write-Log "Kan backup map niet aanmaken: $_" -Level ERROR
        [System.Windows.Forms.MessageBox]::Show("Backup map aanmaken mislukt:`n$_","Fout",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $btnStart.Enabled = $true
        return
    }

    # Sla metadata op zodat herstel de backup locatie kan herkennen
    $destType   = if ($rbNetwork.Checked) { 'network' } else { 'local' }
    $netBase    = if ($rbNetwork.Checked) { $txtNetworkPath.Text.Trim() } else { '' }
    Save-BackupMetadata -BackupPath $script:BackupRoot -SourceComputer $sourceComputer `
                        -SourceUser $sourceUser -DestType $destType -NetworkBase $netBase

    # Verbind remote share indien nodig
    $sourceUserFolder = $null
    if (-not $isLocal) {
        $sharePath = Connect-RemoteShare -ComputerName $sourceComputer -Credential $script:SourceCred
        if (-not $sharePath) {
            $ans = [System.Windows.Forms.MessageBox]::Show(
                "Verbinding met \\$sourceComputer\c$ mislukt.`n`nWilt u doorgaan met alleen lokale functies`n(WiFi, omgevingsvariabelen, etc.)?",
                "Verbinding mislukt",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) {
                $btnStart.Enabled = $true
                return
            }
            $sourceUserFolder = $null
        } else {
            $sourceUserFolder = "\\$sourceComputer\c$\Users\$sourceUser"
        }
    } else {
        # Lokaal: gebruik USERPROFILE voor actieve gebruiker, anders zoek via registry (domein suffix zoals .UVION)
        if ($sourceUser -ieq $env:USERNAME) {
            $sourceUserFolder = $env:USERPROFILE
        } else {
            # Zoek profielpad via registry (ProfileList) voor andere gebruikers
            $profileFolder = $null
            $profileList = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" -ErrorAction SilentlyContinue
            foreach ($key in $profileList) {
                $profilePath = (Get-ItemProperty $key.PSPath -Name ProfileImagePath -ErrorAction SilentlyContinue).ProfileImagePath
                if ($profilePath -and ($profilePath -like "*\$sourceUser" -or $profilePath -like "*\$sourceUser.*")) {
                    $profileFolder = $profilePath
                    break
                }
            }
            $sourceUserFolder = if ($profileFolder) { $profileFolder } else { "C:\Users\$sourceUser" }
        }
        Write-Log "Profielmap: $sourceUserFolder"
    }

    if ($sourceUserFolder -and -not (Test-Path $sourceUserFolder)) {
        Write-Log "Gebruikersmap niet gevonden: $sourceUserFolder" -Level ERROR
        [System.Windows.Forms.MessageBox]::Show("Gebruikersmap niet gevonden:`n$sourceUserFolder","Fout",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $btnStart.Enabled = $true
        Disconnect-RemoteShare
        return
    }

    # Tel stappen voor voortgangsbalk
    $selectedChecks       = $checkboxes.Values | Where-Object { $_.Checked }
    $script:TotalSteps    = $selectedChecks.Count + 2  # +2 voor zip + rapport
    $script:CurrentStep   = 0

    try {
        # Browsers + Office
        if ($checkboxes['chkPrinters'].Checked)    { Backup-Printers       -BackupPath $script:BackupRoot -SourceComputer $sourceComputer }
        if ($checkboxes['chkChrome'].Checked  -and $sourceUserFolder) { Backup-ChromeProfile  -BackupPath $script:BackupRoot -SourceUserFolder $sourceUserFolder }
        if ($checkboxes['chkEdge'].Checked    -and $sourceUserFolder) { Backup-EdgeProfile    -BackupPath $script:BackupRoot -SourceUserFolder $sourceUserFolder }
        if ($checkboxes['chkFirefox'].Checked -and $sourceUserFolder) { Backup-FirefoxProfile -BackupPath $script:BackupRoot -SourceUserFolder $sourceUserFolder }
        if ($checkboxes['chkOutlookSig'].Checked -and $sourceUserFolder) { Backup-OutlookSignatures -BackupPath $script:BackupRoot -SourceUserFolder $sourceUserFolder }
        if ($checkboxes['chkOutlookPST'].Checked -and $sourceUserFolder) { Backup-OutlookPST  -BackupPath $script:BackupRoot -SourceUserFolder $sourceUserFolder }

        # Gebruikersbestanden
        $folders = @()
        if ($checkboxes['chkDesktop'].Checked)   { $folders += 'Desktop' }
        if ($checkboxes['chkDocuments'].Checked) { $folders += 'Documents' }
        if ($checkboxes['chkDownloads'].Checked) { $folders += 'Downloads' }
        if ($checkboxes['chkPictures'].Checked)  { $folders += 'Pictures' }
        if ($checkboxes['chkMusic'].Checked)     { $folders += 'Music' }
        if ($checkboxes['chkVideos'].Checked)    { $folders += 'Videos' }
        if ($folders.Count -gt 0 -and $sourceUserFolder) {
            Backup-UserFolders -BackupPath $script:BackupRoot -SourceUserFolder $sourceUserFolder -Folders $folders
        }

        # Systeem/netwerk
        if ($checkboxes['chkMappedDrives'].Checked) { Backup-MappedDrives -BackupPath $script:BackupRoot -SourceComputer $sourceComputer }
        if ($checkboxes['chkWiFi'].Checked)          { Backup-WiFiProfiles -BackupPath $script:BackupRoot -SourceComputer $sourceComputer }
        if ($checkboxes['chkStickyNotes'].Checked -and $sourceUserFolder) { Backup-StickyNotes -BackupPath $script:BackupRoot -SourceUserFolder $sourceUserFolder }
        if ($checkboxes['chkWallpaper'].Checked   -and $sourceUserFolder) { Backup-Wallpaper   -BackupPath $script:BackupRoot -SourceUserFolder $sourceUserFolder -SourceComputer $sourceComputer }
        if ($checkboxes['chkEnvVars'].Checked)       { Export-UserEnvVars  -BackupPath $script:BackupRoot -SourceComputer $sourceComputer }

        # Optioneel alles zippen
        if ($chkZip.Checked) {
            Update-Progress "Alles zippen naar ZIP bestand..."
            $zipPath = "$($script:BackupRoot).zip"
            Compress-WithProgress -SourceDir $script:BackupRoot -ZipPath $zipPath
            $zipMB = [Math]::Round((Get-Item $zipPath).Length / 1MB, 0)
            Write-Log "Alles gezipt: $zipPath ($zipMB MB)" -Level OK
        }

        # Rapport genereren
        Update-Progress "Migratierapport genereren..."
        New-MigrationReport -BackupPath $script:BackupRoot | Out-Null

        $script:ProgressBar.Value = 100
        $script:StatusLabel.Text  = "Migratie voltooid!"

        $errors   = ($script:LogEntries | Where-Object { $_ -match '\[ERROR\]' }).Count
        $warnings = ($script:LogEntries | Where-Object { $_ -match '\[WARN\]'  }).Count

        Write-Log "=== MIGRATIE VOLTOOID (fouten: $errors, waarschuwingen: $warnings) ===" -Level OK

        $btnOpenFolder.Enabled = $true
        $btnOpenFolder.Add_Click({
            if (Test-Path $script:BackupRoot) { Start-Process explorer.exe $script:BackupRoot }
        })

        [System.Windows.Forms.MessageBox]::Show(
            "Migratie voltooid!`n`n" +
            "Fouten      : $errors`n" +
            "Waarschuwingen: $warnings`n`n" +
            "Backup opgeslagen in:`n$($script:BackupRoot)`n`n" +
            "Rapport: MIGRATIE_RAPPORT.txt",
            "Migratie Voltooid",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

    } catch {
        Write-Log "Onverwachte fout: $_" -Level ERROR
        [System.Windows.Forms.MessageBox]::Show("Onverwachte fout:`n$_","Fout",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } finally {
        Disconnect-RemoteShare
        $btnStart.Enabled = $true
    }
})

# ─── Toon formulier ───────────────────────────────────────────────────────────
Write-Log "Tool geladen. Kies modus (Backup of Herstel) en klik 'Start Backup' of 'Start Herstel'."
$form.ShowDialog() | Out-Null
