@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY goto nopython

if exist ".venv\Scripts\python.exe" goto check
echo Erster Start: Das Programm wird eingerichtet ...
%PY% -m venv .venv
if errorlevel 1 goto fehler

:check
rem Neue oder geaenderte Bausteine nachinstallieren
fc /b requirements.txt ".venv\installiert.txt" >nul 2>nul
if not errorlevel 1 goto start
echo.
echo Bausteine werden installiert bzw. aktualisiert.
echo Beim ersten Mal mit Spracherkennung dauert das etwa 5 bis 10 Minuten (ca. 1 GB Download).
echo Bitte das Fenster offen lassen ...
echo.
".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto fehler
copy /y requirements.txt ".venv\installiert.txt" >nul

:start
start "" ".venv\Scripts\pythonw.exe" main.py
exit /b 0

:nopython
echo Python wurde nicht gefunden.
echo Bitte Python 3.10 oder neuer von https://www.python.org/downloads/ installieren
echo und dabei den Haken "Add python.exe to PATH" setzen. Danach diese Datei erneut starten.
pause
exit /b 1

:fehler
echo Beim Einrichten ist ein Fehler aufgetreten. Bitte die Meldungen oben an das Team schicken.
pause
exit /b 1
