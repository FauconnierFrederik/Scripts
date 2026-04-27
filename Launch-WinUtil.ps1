# ============================================================
#  Launch-WinUtil.ps1
#  Start Chris Titus Tech's Windows Utility via internet.
#  Bron: https://github.com/christitustech/winutil
# ============================================================

Add-Type -AssemblyName System.Windows.Forms

$bevestiging = [System.Windows.Forms.MessageBox]::Show(
    "Dit start een extern script van Chris Titus Tech (WinUtil).`n`nWinUtil biedt tools voor:`n  - Software installeren`n  - Windows tweaks en optimalisaties`n  - Systeem reparaties`n`nBron: https://github.com/christitustech/winutil`n`nDoorgaan?",
    "WinUtil – Bevestiging",
    "YesNo",
    "Question"
)

if ($bevestiging -ne "Yes") { exit }

Write-Host "`n[*] WinUtil downloaden en starten..." -ForegroundColor Cyan

try {
    Invoke-RestMethod -Uri "https://christitus.com/win" | Invoke-Expression
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "WinUtil kon niet worden gestart.`n`nFout: $_`n`nControleer je internetverbinding.",
        "WinUtil – Fout", "OK", "Error"
    )
}
