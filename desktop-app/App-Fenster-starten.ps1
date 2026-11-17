# Gesprächsnotiz-AI als eigenes App-Fenster starten (Demo, portabel)
#
# So funktioniert es:
#  1. GespraechsnotizAI.exe wird unsichtbar gestartet (kein schwarzes Konsolenfenster),
#     auf einem eigenen Port (8517), damit es die normale Browser-Version (Port 8501) nicht stört.
#  2. Microsoft Edge (oder Chrome) öffnet die Oberfläche im "App-Modus":
#     eigenes Fenster, keine Adressleiste, keine Tabs.
#  3. Wird das App-Fenster geschlossen, beendet das Skript das Programm samt KI-Server.
#
# Dieses Skript gehört in denselben Ordner wie START_HIER.txt.

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Title      = "Gesprächsnotiz-AI"
$Root       = $PSScriptRoot
$AppDir     = Join-Path $Root "GespraechsnotizAI"
$Exe        = Join-Path $AppDir "GespraechsnotizAI.exe"
$Port       = 8517
$Url        = "http://localhost:$Port"
$DataDir    = Join-Path $env:LOCALAPPDATA "GespraechsnotizAI"
$ProfileDir = Join-Path $DataDir "AppFenster"
$LogFile    = Join-Path $DataDir "app-fenster.log"

function Show-Message([string]$Text, [string]$Icon = "Information") {
    [System.Windows.Forms.MessageBox]::Show($Text, $Title, "OK", $Icon) | Out-Null
}

function Test-Server {
    try {
        $response = Invoke-WebRequest "$Url/_stcore/health" -UseBasicParsing -TimeoutSec 2
        return $response.StatusCode -eq 200
    } catch {
        return $false
    }
}

function Find-Browser {
    $candidates = @()
    foreach ($name in "msedge.exe", "chrome.exe") {
        foreach ($hive in "HKLM:", "HKCU:") {
            $key = "$hive\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$name"
            try { $candidates += (Get-ItemProperty $key -ErrorAction Stop).'(default)' } catch { }
        }
    }
    $candidates += "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
    $candidates += "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
    $candidates += "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
    return $candidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
}

function Get-AppWindowProcesses {
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe' OR Name='chrome.exe'" |
        Where-Object { $_.CommandLine -and $_.CommandLine.Contains($ProfileDir) }
}

function Stop-ProcessTree([int]$ProcessId) {
    # erst die Kinder (KI-Server, Streamlit), dann den Prozess selbst
    Get-CimInstance Win32_Process -Filter "ParentProcessId=$ProcessId" |
        ForEach-Object { Stop-ProcessTree $_.ProcessId }
    Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
}

# --- Vorbereitungen ---------------------------------------------------------
if (-not (Test-Path $Exe)) {
    Show-Message "GespraechsnotizAI.exe wurde nicht gefunden:`n$Exe`n`nDieses Skript muss im selben Ordner liegen wie START_HIER.txt." "Error"
    exit 1
}
New-Item -ItemType Directory -Force -Path $ProfileDir | Out-Null

$Browser = Find-Browser
$ServerProcess = $null

# --- Programm unsichtbar starten (falls das App-Fenster nicht schon läuft) --
if (-not (Test-Server)) {
    $info = New-Object System.Diagnostics.ProcessStartInfo
    $info.FileName = $env:ComSpec
    # Ausgaben landen in einer Log-Datei; "< NUL" verhindert, dass etwas auf Eingaben wartet
    $info.Arguments = "/d /c `"`"$Exe`" < NUL > `"$LogFile`" 2>&1`""
    $info.WorkingDirectory = $AppDir
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.EnvironmentVariables["STREAMLIT_SERVER_PORT"] = "$Port"
    $info.EnvironmentVariables["STREAMLIT_SERVER_SHOW_EMAIL_PROMPT"] = "false"
    $info.EnvironmentVariables["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    $info.EnvironmentVariables["STREAMLIT_CLIENT_TOOLBAR_MODE"] = "minimal"
    # Streamlit soll keinen eigenen Browser-Tab öffnen: BROWSER zeigt auf eine Datei, die nichts tut.
    # (Python erwartet hier genau einen Programmpfad, keine Befehle mit Argumenten.)
    $noBrowser = Join-Path $DataDir "kein-browser.cmd"
    Set-Content -Path $noBrowser -Value "@exit /b 0" -Encoding ASCII
    $info.EnvironmentVariables["BROWSER"] = $noBrowser
    $info.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8"
    $ServerProcess = [System.Diagnostics.Process]::Start($info)

    # kleines Startfenster, solange KI und Oberfläche laden
    $splash = New-Object System.Windows.Forms.Form
    $splash.Text = $Title
    $splash.Size = New-Object System.Drawing.Size(420, 140)
    $splash.StartPosition = "CenterScreen"
    $splash.FormBorderStyle = "FixedDialog"
    $splash.ControlBox = $false
    $splash.TopMost = $true
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Gesprächsnotiz-AI wird gestartet …`nDie lokale KI wird geladen, das dauert einen Moment."
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $label.Dock = "Fill"
    $label.TextAlign = "MiddleCenter"
    $splash.Controls.Add($label)
    $splash.Show()

    $deadline = (Get-Date).AddMinutes(4)
    while (-not (Test-Server)) {
        [System.Windows.Forms.Application]::DoEvents()
        if ($ServerProcess.HasExited -or (Get-Date) -gt $deadline) {
            $splash.Close()
            $tail = if (Test-Path $LogFile) { (Get-Content $LogFile -Tail 15) -join "`n" } else { "" }
            Show-Message "Das Programm konnte nicht gestartet werden.`n`nProtokoll: $LogFile`n`n$tail" "Error"
            if (-not $ServerProcess.HasExited) { Stop-ProcessTree $ServerProcess.Id }
            exit 1
        }
        Start-Sleep -Milliseconds 300
    }
    $splash.Close()
}

# --- App-Fenster öffnen -----------------------------------------------------
if (-not $Browser) {
    Start-Process $Url
    Show-Message "Microsoft Edge wurde nicht gefunden, deshalb läuft die App im normalen Browser.`n`nKlicke auf OK, um Gesprächsnotiz-AI zu beenden."
} else {
    $arguments = @(
        "--app=$Url",
        "--user-data-dir=`"$ProfileDir`"",
        "--window-size=1400,900",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-background-mode"
    )
    Start-Process -FilePath $Browser -ArgumentList $arguments | Out-Null
    Start-Sleep -Seconds 3
    # warten, bis das App-Fenster geschlossen wurde
    while (Get-AppWindowProcesses) {
        Start-Sleep -Seconds 2
    }
}

# --- Aufräumen ----------------------------------------------------------------
if ($ServerProcess -and -not $ServerProcess.HasExited) {
    Stop-ProcessTree $ServerProcess.Id
}
