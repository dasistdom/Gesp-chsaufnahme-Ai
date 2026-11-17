@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY goto nopython

if exist ".venv\Scripts\pythonw.exe" goto start
echo Erster Start: Das Programm wird einmalig eingerichtet, das dauert 1 bis 2 Minuten ...
%PY% -m venv .venv
if errorlevel 1 goto fehler
".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto fehler

:start
start "" ".venv\Scripts\pythonw.exe" main.py
exit /b 0

:nopython
echo Python wurde nicht gefunden.
echo Bitte Python 3.11 oder neuer von https://www.python.org/downloads/ installieren
echo und dabei den Haken "Add python.exe to PATH" setzen. Danach diese Datei erneut starten.
pause
exit /b 1

:fehler
echo Beim Einrichten ist ein Fehler aufgetreten. Bitte die Meldungen oben an das Team schicken.
pause
exit /b 1
