# Legt auf dem Desktop eine Verknüpfung "Gesprächsnotiz-AI (App)" an,
# die das Programm als eigenes App-Fenster startet (ohne Konsolenfenster).
# Dieses Skript gehört in denselben Ordner wie START_HIER.txt.

$script = Join-Path $PSScriptRoot "App-Fenster-starten.ps1"
$exe = Join-Path $PSScriptRoot "GespraechsnotizAI\GespraechsnotizAI.exe"

if (!(Test-Path $script) -or !(Test-Path $exe)) {
    Write-Host "FEHLER: App-Fenster-starten.ps1 oder GespraechsnotizAI.exe wurde nicht gefunden." -ForegroundColor Red
    Write-Host "Beide Skripte muessen im selben Ordner liegen wie START_HIER.txt."
    Read-Host "Eingabetaste zum Beenden"
    exit 1
}

$desktop = [Environment]::GetFolderPath("Desktop")
$linkPath = Join-Path $desktop "Gespraechsnotiz-AI (App).lnk"

$shell = New-Object -ComObject WScript.Shell
$link = $shell.CreateShortcut($linkPath)
$link.TargetPath = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$link.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`""
$link.WorkingDirectory = $PSScriptRoot
$link.IconLocation = "$exe,0"
$link.WindowStyle = 7  # minimiert starten
$link.Description = "Gesprächsnotiz-AI als eigenes App-Fenster"
$link.Save()

Write-Host "Desktop-Verknuepfung wurde erstellt:" -ForegroundColor Green
Write-Host $linkPath
Write-Host ""
Write-Host "Die normale Browser-Version (GespraechsnotizAI.exe) funktioniert weiterhin wie bisher."
Read-Host "Eingabetaste zum Beenden"
