# Gesprächsnotiz-AI – Windows-Programm

Eigenständiges Windows-Programm mit eigener Oberfläche (Qt / PySide6) – kein Browser.
Die bisherige Website-Version (`app.py`) bleibt unverändert und funktioniert weiter.

## Was das Programm kann
| Bereich | Funktion |
|---|---|
| **Aufnehmen** | Systemton (z. B. LearnSpace3D) und optional Mikrofon aufnehmen, Pause/Weiter, automatisch neue Teile alle X Minuten, Pegelanzeige |
| **Auswertung** | Jeder fertige Teil (bei Pause oder nach X Minuten) wird **sofort im Hintergrund** ausgewertet: Spracherkennung (Whisper) + Lernnotizen (lokale KI). Nach dem Stoppen entsteht automatisch das Word-Dokument. Auch vorhandene Aufnahmen (.wav) lassen sich auswerten. |
| **Teilen** | Notizen automatisch in euren **Discord-Kanal** schicken – jeder fertige Teil schon während der Vorlesung, am Ende Zusammenfassung + Word-Dokument. Keine Aufnahmen, keine Transkripte. |
| **Notizen** | Alle Word-Notizen auflisten, durchsuchen, ansehen und in Word öffnen |
| **Einstellungen** | KI-Modell (z. B. Qwen 3 8B), Kontextgröße, Spracherkennung, Programmordner; werden gespeichert |

Die KI, die Modelle und ffmpeg kommen aus der **vorhandenen GespraechsnotizAI-Installation**. Liegt dieser
App-Ordner darin (z. B. `C:\Test\GespraechsnotizAI\Windows-App`), wird sie automatisch gefunden – sonst unter
Einstellungen den Ordner wählen, in dem `GespraechsnotizAI.exe` liegt. Notizen, Aufnahmen und gespeicherte
Auswertungen liegen wie bei der Website unter `Dokumente\GespraechsnotizAI`.

## Discord einrichten (einmalig)
1. In Discord: beim Kanal auf das **Zahnrad** („Kanal bearbeiten“) → **Integrationen** → **Webhooks** → **Neuer Webhook**.
2. Namen geben, z. B. „Gesprächsnotiz-AI“ → **Webhook-URL kopieren**.
3. Im Programm: **Einstellungen** → „Mit der Gruppe teilen“ → Adresse einfügen → **Test senden**.
4. Haken bei „Notizen automatisch … schicken“ setzen.

Die Webhook-Adresse ist wie ein Schlüssel zum Kanal: nicht öffentlich teilen und **nicht** in GitHub hochladen
(sie wird nur lokal in `Dokumente\GespraechsnotizAI\windows-app-einstellungen.json` gespeichert).

## Starten (Entwicklung)
1. Einmalig **Python 3.10+** von python.org installieren (Haken „Add python.exe to PATH“).
2. Doppelklick auf **`Starten.bat`**. Beim ersten Mal (und nach Updates) werden die Bausteine automatisch installiert –
   mit Spracherkennung ca. 1 GB, 5–10 Minuten.

## Aufbau
- `main.py` – Oberfläche (Fenster, Seiten, Knöpfe)
- `aufnahme.py` – Aufnahme-Technik, 1:1 aus `app.py` übernommen (Windows WASAPI, nur `ctypes` + `numpy`)
- `auswertung.py` – Spracherkennung, KI und Word-Export, aus `app.py` übernommen (gleiche Qualität wie die Website)
- `notizen.py` – liest `.docx`-Dateien für Vorschau und Suche
- `teilen.py` – schickt Notizen über einen Discord-Webhook

## Nächste Schritte (siehe Fahrplan)
1. Lehrkraft und eigene Beiträge im Transkript unterscheiden
2. ⭐ „Wichtig“-Knopf während der Vorlesung
3. Quiz und Lernkarten
4. Unterrichtsmaterial (PowerPoint) hochladen und Notizen darauf abstimmen
5. Lern-Präsentation erstellen
6. Portable `.exe` automatisch über GitHub Actions bauen – dann braucht niemand mehr Python
