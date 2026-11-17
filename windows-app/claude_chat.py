# -*- coding: utf-8 -*-
"""Zusammenarbeit mit Claude über claude.ai – kostenlos mit einem Gratis-Konto, aber von Hand.

1. build_prompt(): Anweisung + Transkript für die Zwischenablage, zum Einfügen auf claude.ai
2. save_answer(): Claudes Antwort (Markdown aus der Zwischenablage) als Word-Dokument speichern
"""
from __future__ import annotations

import datetime as dt
import re
from pathlib import Path

from docx import Document
from docx.shared import Pt

CLAUDE_URL = "https://claude.ai/new"
LARGE_PROMPT = 250_000  # Zeichen; sehr lange Texte können die Grenze eines Gratis-Kontos sprengen

PROMPT = """Du bist ein Lernassistent für Teilnehmende einer Umschulung. Unten steht das automatisch erstellte \
Transkript einer Vorlesung{topic}{date}. Es enthält Hörfehler der Spracherkennung – deute falsch erkannte Wörter \
sinnvoll aus dem Zusammenhang und lass Unverständliches weg.{speakers}

Erstelle daraus Lernnotizen auf Deutsch, neutral in der dritten Person, in genau diesem Format (Markdown):

# {title}

## Worum ging es?
2 bis 4 Sätze.

## Die wichtigsten Punkte
Nach Themen gegliedert (### Zwischenüberschriften), darunter Stichpunkte als vollständige Sätze, mit Beispielen \
aus der Vorlesung.

## Fachbegriffe
- Begriff – kurze Erklärung

## Aufgaben und Termine
Nur was ausdrücklich genannt wurde, sonst „Keine genannt“.
{questions}
## Prüfungsfragen zum Üben
5 Fragen zum Stoff, jeweils mit kurzer Lösung.

Erfinde keine Inhalte, die nicht im Transkript stehen.

TRANSKRIPT:
{transcript}
"""

SPEAKER_HINT = (" Zeilen mit „Du:“ stammen von der Person, die mitschreibt (eigene Fragen und Beiträge), "
                "„Lehrkraft:“ ist der Ton aus der Vorlesung.")
QUESTIONS_SECTION = "\n## Fragen und Antworten\nEigene Fragen aus dem Kurs mit der Antwort der Lehrkraft.\n"


def clean_title(title: str) -> str:
    return re.sub(r"^(Vorlesungsnotizen|Notizen nach Folien):\s*", "", title).strip() or "Vorlesung"


def build_prompt(title: str, date: dt.datetime | None, transcript: str, has_speakers: bool) -> str:
    name = clean_title(title)
    return PROMPT.format(
        topic=f" zum Thema „{name}“" if name != "Vorlesung" else "",
        date=f" vom {date:%d.%m.%Y}" if date else "",
        speakers=SPEAKER_HINT if has_speakers else "",
        questions=QUESTIONS_SECTION if has_speakers else "",
        title=f"Lernnotizen: {name}",
        transcript=transcript.strip(),
    )


def _add_markdown(doc: Document, text: str) -> None:
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped in ("---", "***"):
            continue
        clean = re.sub(r"\*\*(.+?)\*\*", r"\1", stripped)
        clean = re.sub(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)", r"\1", clean)
        heading = re.match(r"^(#{1,6})\s+(.*)", clean)
        if heading:
            level = min(len(heading.group(1)), 3)
            doc.add_heading(heading.group(2).strip(), level=level)
        elif re.match(r"^[-*•]\s+", clean):
            indent = len(line) - len(line.lstrip())
            style = "List Bullet 2" if indent >= 2 else "List Bullet"
            doc.add_paragraph(re.sub(r"^[-*•]\s+", "", clean), style=style)
        elif re.match(r"^\d+[.)]\s+", clean):
            doc.add_paragraph(re.sub(r"^\d+[.)]\s+", "", clean), style="List Number")
        else:
            doc.add_paragraph(clean)


def looks_like_answer(text: str) -> bool:
    """Grobe Prüfung, ob die Zwischenablage eine Claude-Antwort und nicht z. B. den Prompt enthält."""
    return len(text.strip()) > 200 and "TRANSKRIPT:" not in text and "##" in text


def save_answer(text: str, original: Path, notes_dir: Path) -> tuple[Path, str]:
    """Speichert Claudes Antwort als Word-Dokument neben den Notizen. Gibt (Pfad, Titel) zurück."""
    lines = text.strip().splitlines()
    title = "Lernnotizen (Claude)"
    if lines and lines[0].lstrip().startswith("# "):
        title = lines.pop(0).lstrip()[2:].strip()
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    doc.add_heading(title, level=0)
    note = doc.add_paragraph(f"Erstellt mit Claude (claude.ai) aus dem Transkript von „{original.name}“ "
                             f"am {dt.datetime.now():%d.%m.%Y um %H:%M} Uhr. Bitte wichtige Stellen nachprüfen.")
    note.runs[0].italic = True
    _add_markdown(doc, "\n".join(lines))
    notes_dir.mkdir(parents=True, exist_ok=True)
    path = notes_dir / f"{original.stem}_Claude.docx"
    counter = 2
    while path.exists():
        path = notes_dir / f"{original.stem}_Claude_{counter}.docx"
        counter += 1
    doc.save(path)
    return path, title
