# -*- coding: utf-8 -*-
"""Teilen mit der Gruppe über einen Discord-Webhook.

Geteilt werden nur Notizen und Zusammenfassungen – keine Aufnahmen und keine Transkripte.
Ein Webhook ist eine Adresse, über die ein Programm Nachrichten in genau einen Discord-Kanal
schreiben darf. Wer die Adresse kennt, kann dort posten: also nicht öffentlich teilen.
"""
from __future__ import annotations

import json
import queue
import threading
import time
import urllib.error
import urllib.request
import uuid
from pathlib import Path
from typing import Callable

WEBHOOK_PREFIXES = (
    "https://discord.com/api/webhooks/",
    "https://discordapp.com/api/webhooks/",
    "https://ptb.discord.com/api/webhooks/",
    "https://canary.discord.com/api/webhooks/",
)
MESSAGE_LIMIT = 1900  # Discord erlaubt 2000 Zeichen pro Nachricht
USER_AGENT = "GespraechsnotizAI-Windows-Programm/1.0"


def is_webhook(url: str) -> bool:
    return url.strip().startswith(WEBHOOK_PREFIXES)


def split_message(text: str, limit: int = MESSAGE_LIMIT) -> list[str]:
    """Teilt langen Text an Zeilengrenzen in Nachrichten, die Discord annimmt."""
    chunks, current = [], ""
    for line in text.splitlines():
        while len(line) > limit:
            chunks.append(line[:limit])
            line = line[limit:]
        if len(current) + len(line) + 1 > limit and current:
            chunks.append(current.rstrip())
            current = ""
        current += line + "\n"
    if current.strip():
        chunks.append(current.rstrip())
    return chunks


def _post(url: str, body: bytes, content_type: str) -> None:
    target = url.strip() + ("&" if "?" in url else "?") + "wait=true"
    for attempt in range(4):
        request = urllib.request.Request(target, data=body, method="POST",
                                         headers={"Content-Type": content_type, "User-Agent": USER_AGENT})
        try:
            with urllib.request.urlopen(request, timeout=30):
                return
        except urllib.error.HTTPError as exc:
            if exc.code == 429 and attempt < 3:  # zu viele Nachrichten auf einmal: kurz warten
                try:
                    wait = float(json.loads(exc.read().decode("utf-8")).get("retry_after", 1.0))
                except (ValueError, AttributeError):
                    wait = 1.0
                time.sleep(min(wait, 10.0) + 0.2)
                continue
            if exc.code in (401, 404):
                raise RuntimeError("Die Webhook-Adresse ist ungültig oder wurde gelöscht.") from exc
            raise RuntimeError(f"Discord hat die Nachricht abgelehnt (Fehler {exc.code}).") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError("Discord ist nicht erreichbar. Besteht eine Internetverbindung?") from exc


def send_text(url: str, text: str) -> None:
    for index, chunk in enumerate(split_message(text)):
        if index:
            time.sleep(0.6)  # Discord begrenzt, wie schnell nacheinander gepostet werden darf
        _post(url, json.dumps({"content": chunk, "allowed_mentions": {"parse": []}}).encode("utf-8"),
              "application/json")


def send_file(url: str, path: Path, text: str = "") -> None:
    boundary = uuid.uuid4().hex
    payload = {"content": text[:MESSAGE_LIMIT], "allowed_mentions": {"parse": []},
               "attachments": [{"id": 0, "filename": path.name}]}
    body = (
        f"--{boundary}\r\nContent-Disposition: form-data; name=\"payload_json\"\r\n"
        f"Content-Type: application/json\r\n\r\n{json.dumps(payload)}\r\n"
        f"--{boundary}\r\nContent-Disposition: form-data; name=\"files[0]\"; filename=\"{path.name}\"\r\n"
        f"Content-Type: application/octet-stream\r\n\r\n"
    ).encode("utf-8") + path.read_bytes() + f"\r\n--{boundary}--\r\n".encode("utf-8")
    _post(url, body, f"multipart/form-data; boundary={boundary}")


def _fmt(seconds: float) -> str:
    seconds = int(seconds)
    h, rest = divmod(seconds, 3600)
    m, s = divmod(rest, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


class DiscordShare:
    """Schickt Notizen im Hintergrund an Discord, damit Aufnahme und Auswertung nie warten müssen."""

    def __init__(self, get_url: Callable[[], str], is_enabled: Callable[[], bool]) -> None:
        self.get_url = get_url
        self.is_enabled = is_enabled
        self.status = ""
        self._queue: queue.Queue = queue.Queue()
        threading.Thread(target=self._work, name="discord", daemon=True).start()

    @property
    def active(self) -> bool:
        return self.is_enabled() and is_webhook(self.get_url())

    def part_done(self, title: str, part: dict) -> None:
        notes = [c for c in part.get("chunks", []) if c.get("notes", "").strip()]
        if not self.active or not notes:
            return
        head = f"📝 **{title or 'Vorlesung'}** · Teil {part['number']} (ab {_fmt(part['offset'])})"
        body = "\n\n".join(f"**{c['label']}**\n{c['notes'].strip()}" for c in notes)
        self._queue.put(("text", f"{head}\n\n{body}", None))

    def finished(self, title: str, summary: str, docx_path: Path) -> None:
        if not self.active:
            return
        self._queue.put(("text", f"✅ **{title or 'Vorlesung'}** – Zusammenfassung\n\n{summary.strip()}", None))
        self._queue.put(("file", "📎 Alle Notizen als Word-Dokument", Path(docx_path)))

    def test(self, url: str) -> None:
        self._queue.put(("test", "👋 Gesprächsnotiz-AI ist verbunden. Notizen erscheinen ab jetzt hier.", url))

    def _work(self) -> None:
        while True:
            kind, text, extra = self._queue.get()
            url = extra if kind == "test" else self.get_url()
            try:
                if kind == "file":
                    send_file(url, extra, text)
                else:
                    send_text(url, text)
                self.status = f"Discord: zuletzt gesendet um {time.strftime('%H:%M')} Uhr"
            except Exception as exc:
                self.status = f"Discord: {exc}"
