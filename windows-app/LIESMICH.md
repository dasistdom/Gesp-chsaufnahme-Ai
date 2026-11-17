# Gesprächsnotiz-AI – Windows-Programm (Demo)

Eigenständiges Windows-Programm mit eigener Oberfläche (Qt / PySide6) – kein Browser.
Die bisherige Website-Version (`app.py`) bleibt unverändert und funktioniert weiter.

## Was die Demo kann
| Bereich | Funktion |
|---|---|
| **Aufnehmen** | Systemton (z. B. LearnSpace3D) und optional Mikrofon aufnehmen, Pause/Weiter, automatisch neue Teile alle X Minuten, Pegelanzeige, Liste der gespeicherten Teile |
| **Notizen** | Alle erstellten Word-Notizen auflisten, durchsuchen, ansehen und in Word öffnen |
| **Einstellungen** | Ordner für Aufnahmen und Notizen öffnen |

Noch **nicht** enthalten: Spracherkennung und KI-Zusammenfassung. Dafür vorerst die Website-Version nutzen
(„Ältere Aufnahme auswerten“) – beide verwenden dieselben Ordner unter `Dokumente\GespraechsnotizAI`.

## Starten (Entwicklung)
1. Einmalig **Python 3.11+** von python.org installieren (Haken „Add python.exe to PATH“).
2. Doppelklick auf **`Starten.bat`**. Beim ersten Mal wird automatisch eingerichtet (1–2 Minuten).

## Aufbau
- `main.py` – Oberfläche (Fenster, Seiten, Knöpfe)
- `aufnahme.py` – Aufnahme-Technik, 1:1 aus `app.py` übernommen (Windows WASAPI, nur `ctypes` + `numpy`)
- `notizen.py` – liest `.docx`-Dateien für Vorschau und Suche

## Nächste Schritte
1. Gemeinsamen Kern bauen: Aufnahme, Spracherkennung, KI und Word-Export als Module, die Website **und** Programm nutzen.
2. Auswertung direkt im Programm (Fortschritt pro Teil, wie in der Website).
3. Portable `.exe` automatisch über GitHub Actions bauen lassen (Windows-Rechner bei GitHub) – dann braucht niemand mehr Python.
