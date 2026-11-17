# -*- coding: utf-8 -*-
"""Auswertung: Spracherkennung (Whisper), lokale KI (llama-server mit Qwen) und Word-Export.

Übernommen aus app.py (Website-Version), damit beide Versionen gleich gute Notizen schreiben.
Nutzt die vorhandene Installation von GespraechsnotizAI mit (KI-Server, Modelle, ffmpeg).
"""
from __future__ import annotations

import atexit
import ctypes
import datetime as dt
import io
import json
import os
import queue
import re
import socket
import subprocess
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Callable

from docx import Document
from docx.shared import Pt

DATA_DIR = Path.home() / "Documents" / "GespraechsnotizAI"
NOTES_DIR = DATA_DIR / "Notizen"
SESSIONS_DIR = DATA_DIR / "Sitzungen"
BUNDLED_MODEL = "qwen2.5-0.5b-instruct-q4_k_m.gguf"
PREFERRED_MODEL = "qwen3-8b"
BUNDLED_AI_URL = "http://127.0.0.1:9931/v1/chat/completions"

# Wird von configure_program_dir() gesetzt: der Ordner der vorhandenen Installation
PROGRAM_DIR: Path | None = None
LLAMA_DIR = Path()
MODELS_DIR = Path()
WHISPER_MODEL_DIR = Path()
FFMPEG_DIR = Path()


def _runtime_of(folder: Path) -> Path | None:
    for candidate in (folder / "_internal" / "runtime", folder / "GespraechsnotizAI" / "_internal" / "runtime"):
        if (candidate / "llama" / "Programm").is_dir():
            return candidate
    return None


def find_program_dir(start: Path) -> Path | None:
    """Sucht die vorhandene GespraechsnotizAI-Installation im Ordner selbst und darüber."""
    for folder in [start, *start.parents]:
        runtime = _runtime_of(folder)
        if runtime:
            return runtime.parent.parent
    return None


def configure_program_dir(folder: Path) -> bool:
    global PROGRAM_DIR, LLAMA_DIR, MODELS_DIR, WHISPER_MODEL_DIR, FFMPEG_DIR
    runtime = _runtime_of(folder)
    if runtime is None:
        return False
    PROGRAM_DIR = runtime.parent.parent
    LLAMA_DIR = runtime / "llama"
    MODELS_DIR = LLAMA_DIR / "Modelle"
    WHISPER_MODEL_DIR = runtime / "whisper"
    FFMPEG_DIR = runtime / "ffmpeg"
    if FFMPEG_DIR.is_dir() and str(FFMPEG_DIR) not in os.environ.get("PATH", ""):
        os.environ["PATH"] = str(FFMPEG_DIR) + os.pathsep + os.environ.get("PATH", "")  # für Whisper
    return True


def available_models() -> list[Path]:
    if not MODELS_DIR.is_dir():
        return []
    return sorted(MODELS_DIR.glob("*.gguf"), key=lambda p: (p.name != BUNDLED_MODEL, p.name.lower()))


def preferred_model(models: list[Path]) -> Path | None:
    return next((m for m in models if PREFERRED_MODEL in m.name.lower()), None) \
        or max((m for m in models if m.name != BUNDLED_MODEL), key=lambda m: m.stat().st_size, default=None) \
        or (models[0] if models else None)


class AIConfig:
    """Aktuelle KI-Einstellungen. Die Abschnittsgröße richtet sich nach dem Kontextfenster des Modells."""

    def __init__(self) -> None:
        self.configure(BUNDLED_AI_URL, 2048, "Qwen 2.5 0.5B (mitgeliefert)")

    def configure(self, url: str, context: int, label: str) -> None:
        self.url, self.context, self.label = url, context, label
        if context <= 2048:  # mitgeliefertes Modell: sehr kleine Abschnitte
            self.notes_tokens, self.summary_tokens, self.chunk_chars = 450, 600, 2400
        else:
            self.notes_tokens, self.summary_tokens = 700, 1200
            self.chunk_chars = min(20_000, int((context - self.notes_tokens - 400) * 2.2))

_HALLUCINATIONS = re.compile(r"untertitel|amara\.org|copyright|www\.", re.IGNORECASE)


def transcribe_with(model, path: str) -> list[dict]:
    """Gibt eine Liste von Abschnitten mit Start-/Endzeit und Text zurück."""
    result = model.transcribe(
        path, language="de", fp16=False, verbose=False,
        condition_on_previous_text=False,  # verhindert Endlosschleifen bei Pausen
    )
    segments = []
    for seg in result.get("segments", []):
        text = (seg.get("text") or "").strip()
        if not text:
            continue
        if seg.get("no_speech_prob", 0) > 0.6 and seg.get("avg_logprob", 0) < -1.0:
            continue
        if _HALLUCINATIONS.search(text) and len(text) < 80:
            continue
        segments.append({"start": float(seg["start"]), "end": float(seg["end"]), "text": text})
    return segments

class ModelServer:
    """Startet den mitgelieferten llama-server zusätzlich mit einem größeren Modell und mehr Kontext."""

    def __init__(self) -> None:
        self.process: subprocess.Popen | None = None
        self.model: Path | None = None
        self.context = 0
        self.port = 0
        atexit.register(self.stop)

    @property
    def url(self) -> str:
        return f"http://127.0.0.1:{self.port}/v1/chat/completions"

    def running_with(self, model: Path, context: int) -> bool:
        return (self.process is not None and self.process.poll() is None
                and self.model == model and self.context == context)

    def ensure(self, model: Path, context: int) -> str:
        if self.running_with(model, context):
            return self.url
        self.stop()
        exe = LLAMA_DIR / "Programm" / "llama-server.exe"
        if not exe.is_file():
            raise FileNotFoundError(f"KI-Server nicht gefunden: {exe}")
        self.port = _free_port(9932)
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        log_path = DATA_DIR / "ki-server.log"
        threads = max(1, (os.cpu_count() or 2) // 2)  # ≈ Anzahl echter Prozessorkerne
        command = [str(exe), "-m", str(model), "--host", "127.0.0.1", "--port", str(self.port),
                   "-c", str(context), "-np", "1", "-t", str(threads), "--alias", model.name]
        with open(log_path, "w", encoding="utf-8") as log:
            # Kein eigenes Fenster, aber an der Konsole des Programms: schließt sich mit ihr.
            self.process = subprocess.Popen(command, cwd=str(exe.parent), stdout=log,
                                            stderr=subprocess.STDOUT)
        self.model, self.context = model, context
        deadline = time.time() + 300  # große Modelle brauchen eine Weile zum Laden
        while time.time() < deadline:
            if self.process.poll() is not None:
                self.process = None
                tail = log_path.read_text(encoding="utf-8", errors="replace")[-600:]
                raise RuntimeError(f"Der KI-Server ist beim Laden abgebrochen. Zu wenig Arbeitsspeicher? "
                                   f"Auszug aus {log_path}:\n{tail}")
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{self.port}/health", timeout=2) as resp:
                    if resp.status == 200:
                        return self.url
            except Exception:
                pass
            time.sleep(1)
        self.stop()
        raise TimeoutError("Das Modell wurde nicht innerhalb von 5 Minuten geladen.")

    def stop(self) -> None:
        if self.process is not None and self.process.poll() is None:
            self.process.terminate()
            try:
                self.process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                self.process.kill()
        self.process = None
        self.model = None


def _free_port(start: int) -> int:
    for port in range(start, start + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            if sock.connect_ex(("127.0.0.1", port)) != 0:
                return port
    raise RuntimeError("Kein freier Netzwerk-Port für den KI-Server gefunden.")


def total_ram_gb() -> float:
    try:
        class _MemStatus(ctypes.Structure):
            _fields_ = [("dwLength", ctypes.c_uint32), ("dwMemoryLoad", ctypes.c_uint32),
                        ("ullTotalPhys", ctypes.c_uint64), ("ullAvailPhys", ctypes.c_uint64),
                        ("ullTotalPageFile", ctypes.c_uint64), ("ullAvailPageFile", ctypes.c_uint64),
                        ("ullTotalVirtual", ctypes.c_uint64), ("ullAvailVirtual", ctypes.c_uint64),
                        ("ullAvailExtendedVirtual", ctypes.c_uint64)]
        status = _MemStatus()
        status.dwLength = ctypes.sizeof(status)
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status))
        return status.ullTotalPhys / 1024 ** 3
    except Exception:
        return 0.0

def model_label(path: Path) -> str:
    if path.name == BUNDLED_MODEL:
        return "Qwen 2.5 0.5B – mitgeliefert, schnell, aber schwache Notizen"
    if PREFERRED_MODEL in path.name.lower():
        return f"Qwen 3 8B – empfohlen ({path.stat().st_size / 1024 ** 3:.1f} GB)"
    return f"{path.stem} ({path.stat().st_size / 1024 ** 3:.1f} GB)"

SYSTEM_PROMPT = (
    "Du bist ein Lernassistent für Teilnehmende einer Umschulung. Du schreibst klare, "
    "gut lesbare Lernnotizen auf Deutsch, neutral in der dritten Person (z. B. „Die Lehrkraft "
    "erklärt …“, „Es wird besprochen …“) – niemals in der Ich-Form. Du hältst dich streng an "
    "den gegebenen Text, erfindest nichts und wiederholst dich nie. Die Transkripte wurden "
    "automatisch erstellt und enthalten Hörfehler: Deute falsch erkannte Wörter sinnvoll aus dem "
    "Zusammenhang oder lass sie weg. Unverständliche Stellen lässt du weg."
)

NOTES_INSTRUCTIONS = """Schreibe Lernnotizen zu diesem Ausschnitt. Halte dich genau an dieses Format:

- 3 bis 8 Stichpunkte mit den wichtigsten Aussagen, jeder als vollständiger Satz.

Fachbegriffe:
- Begriff – kurze Erklärung (nur echte Fachbegriffe, die im Text vorkommen)

Aufgaben/Termine:
- nur wenn im Text eine Aufgabe, Frist oder ein Termin genannt wird

Lass „Fachbegriffe“ oder „Aufgaben/Termine“ komplett weg, wenn es dazu nichts gibt.
Schreibe die Anweisungen selbst nicht ab und wiederhole keinen Punkt."""

SUMMARY_INSTRUCTIONS = """Erstelle daraus eine Zusammenfassung mit genau diesen Überschriften:

## Worum ging es?
2 bis 4 Sätze Fließtext.

## Die wichtigsten Punkte
5 bis 10 Stichpunkte, jeder als vollständiger Satz.

## Fachbegriffe
Höchstens 10 echte Fachbegriffe, jeweils „Begriff – Erklärung“. Keine Wiederholungen.

## Aufgaben und Termine
Nur was ausdrücklich genannt wurde, sonst „Keine genannt“."""


def ask_ai(prompt: str, max_tokens: int) -> str:
    payload = {
        "model": "local-model",
        "temperature": 0.3,
        "repeat_penalty": 1.15,      # gegen Endlos-Wiederholungen
        "repeat_last_n": 256,
        "max_tokens": max_tokens,
        "chat_template_kwargs": {"enable_thinking": False},  # Qwen3: kein langes "Nachdenken"
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ],
    }
    request = urllib.request.Request(
        AI.url,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=900) as response:
        result = json.loads(response.read().decode("utf-8"))
    content = result["choices"][0]["message"].get("content") or ""
    content = re.sub(r"<think>.*?(</think>|$)", "", content, flags=re.DOTALL)
    return clean_output(content)


_ECHO_LINES = re.compile(
    r"^(lernnotizen:?|die wichtigsten inhalte:?|fachbegriffe mit kurzer erklärung:?|"
    r"aufgaben, termine oder hinweise.*|.*keine (genannten )?(termine|hinweise|aufgaben)( genannt)?\.?)$",
    re.IGNORECASE,
)


def clean_output(text: str) -> str:
    """Entfernt Wiederholungen, abgeschriebene Anweisungen und leere Platzhalter."""
    seen: set[str] = set()
    lines: list[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        key = re.sub(r"[^\w]+", " ", stripped.lower()).strip()
        bare = re.sub(r"^[-*•\d.)\s]+", "", stripped).strip("*: ").strip()
        if key and key in seen:
            continue  # wörtliche Wiederholung
        if bare and _ECHO_LINES.match(bare) and bare.lower().rstrip(".") != "keine genannt":
            continue
        if key:
            seen.add(key)
        lines.append(line.rstrip())
    # leere Zwischenüberschriften (z. B. „Fachbegriffe:“ ohne Punkte darunter) entfernen
    result: list[str] = []
    for index, line in enumerate(lines):
        if line.strip().endswith(":") and not line.lstrip().startswith(("-", "*", "•")):
            following = next((l for l in lines[index + 1:] if l.strip()), "")
            if not following.lstrip().startswith(("-", "*", "•")):
                continue
        result.append(line)
    return re.sub(r"\n{3,}", "\n\n", "\n".join(result)).strip()


def ask_ai_safe(build_prompt: Callable[[str], str], text: str, max_tokens: int) -> str:
    """Fragt die KI; ist der Text zu lang für den Kontext, wird er halbiert."""
    try:
        return ask_ai(build_prompt(text), max_tokens)
    except urllib.error.HTTPError as exc:
        if exc.code == 400 and len(text) > 600:
            middle = len(text) // 2
            cut = text.rfind(". ", 0, middle)
            cut = cut + 1 if cut > middle // 2 else middle
            return (ask_ai_safe(build_prompt, text[:cut], max_tokens) + "\n"
                    + ask_ai_safe(build_prompt, text[cut:], max_tokens))
        raise
    except urllib.error.URLError as exc:
        raise RuntimeError(
            "Der lokale KI-Server ist nicht erreichbar. Bitte das Programm über "
            "GespraechsnotizAI.exe neu starten."
        ) from exc


def fmt_time(seconds: float) -> str:
    seconds = int(seconds)
    h, rest = divmod(seconds, 3600)
    m, s = divmod(rest, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


def split_text(text: str, limit: int | None = None) -> list[str]:
    limit = limit or AI.chunk_chars
    sentences = re.split(r"(?<=[.!?])\s+", text)
    chunks, current = [], ""
    for sentence in sentences:
        while len(sentence) > limit:
            chunks.append(sentence[:limit])
            sentence = sentence[limit:]
        if len(current) + len(sentence) + 1 > limit and current:
            chunks.append(current.strip())
            current = ""
        current += sentence + " "
    if current.strip():
        chunks.append(current.strip())
    return chunks


def chunk_segments(segments: list[dict], limit: int | None = None) -> list[dict]:
    limit = limit or AI.chunk_chars
    chunks: list[dict] = []
    current: list[dict] = []
    length = 0
    for seg in segments:
        if current and length + len(seg["text"]) > limit:
            chunks.append(current)
            current, length = [], 0
        current.append(seg)
        length += len(seg["text"]) + 1
    if current:
        chunks.append(current)
    return [
        {
            "label": f"{fmt_time(c[0]['start'])} – {fmt_time(c[-1]['end'])}",
            "text": " ".join(s["text"] for s in c),
        }
        for c in chunks
    ]


def write_notes(parts: list[dict], title: str, progress: Callable[[float, str], None]) -> None:
    """Schreibt Lernnotizen zu jedem Abschnitt (ergänzt part["notes"])."""
    topic = f" zum Thema „{title}“" if title else ""
    for index, part in enumerate(parts, start=1):
        progress((index - 1) / max(len(parts), 1), f"Abschnitt {index} von {len(parts)} wird zusammengefasst …")
        source = "Ausschnitt aus einem Begleitdokument" if part["label"] == "Begleitdokument" \
            else "Ausschnitt aus dem automatischen Transkript einer Vorlesung bzw. eines Gesprächs"
        part["notes"] = ask_ai_safe(
            lambda text, source=source: (
                f"{source}{topic}:\n\n<<<\n{text}\n>>>\n\n{NOTES_INSTRUCTIONS}"
            ),
            part["text"], AI.notes_tokens,
        ) or "(Keine Notizen erzeugt.)"
    progress(1.0, "Abschnitte fertig.")


def summarize_parts(parts: list[dict], title: str, progress: Callable[[float, str], None]) -> str:
    """Verdichtet die Abschnittsnotizen und erstellt die Kurzfassung der ganzen Vorlesung."""
    topic = f" zum Thema „{title}“" if title else ""
    combined = "\n\n".join(p.get("notes", "") for p in parts)
    if not combined.strip():
        return ""
    rounds = 0
    while len(combined) > AI.chunk_chars and rounds < 4:
        rounds += 1
        progress(rounds / 6, f"Notizen werden verdichtet (Runde {rounds}) …")
        combined = "\n\n".join(
            ask_ai_safe(
                lambda text: (
                    f"Lernnotizen zu Teilen einer Vorlesung{topic}:\n\n<<<\n{text}\n>>>\n\n"
                    "Fasse diese Notizen in 5 bis 10 Stichpunkten kürzer zusammen. Behalte wichtige "
                    "Fachbegriffe, Aufgaben und Termine. Wiederhole nichts."
                ),
                block, AI.notes_tokens,
            )
            for block in split_text(combined)
        )

    progress(5 / 6, "Kurzfassung der ganzen Vorlesung wird erstellt …")
    summary = ask_ai_safe(
        lambda text: (
            f"Lernnotizen zu einer kompletten Vorlesung{topic}:\n\n<<<\n{text}\n>>>\n\n"
            f"{SUMMARY_INSTRUCTIONS}"
        ),
        combined[: AI.chunk_chars + 200], AI.summary_tokens,
    )
    progress(1.0, "Fertig.")
    return summary


def create_notes(
    segments: list[dict], document_text: str, title: str,
    progress: Callable[[float, str], None],
) -> tuple[str, list[dict]]:
    """Fasst abschnittsweise zusammen und erstellt anschließend eine Kurzfassung."""
    parts = chunk_segments(segments)
    for doc_chunk in split_text(document_text) if document_text else []:
        parts.append({"label": "Begleitdokument", "text": doc_chunk})
    if not parts:
        return "", []
    share = len(parts) / (len(parts) + 2)
    write_notes(parts, title, lambda v, t: progress(v * share, t))
    summary = summarize_parts(parts, title, lambda v, t: progress(share + v * (1 - share), t))
    return summary, parts

def _add_markdown(doc: Document, text: str) -> None:
    for line in text.splitlines():
        line = line.rstrip()
        if not line.strip():
            continue
        clean = re.sub(r"\*\*(.+?)\*\*", r"\1", line).strip()
        if clean.startswith("#"):
            doc.add_heading(clean.lstrip("#").strip(), level=3)
        elif re.match(r"^[-*•]\s+", clean):
            doc.add_paragraph(re.sub(r"^[-*•]\s+", "", clean), style="List Bullet")
        elif re.match(r"^\d+[.)]\s+", clean):
            doc.add_paragraph(re.sub(r"^\d+[.)]\s+", "", clean), style="List Number")
        else:
            doc.add_paragraph(clean)


def build_docx(title: str, summary: str, parts: list[dict], segments: list[dict],
               created: dt.datetime) -> bytes:
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    doc.add_heading(f"Vorlesungsnotizen: {title}" if title else "Vorlesungsnotizen", level=0)
    info = f"Erstellt am {created:%d.%m.%Y um %H:%M} Uhr"
    if segments:
        info += f" · Aufnahmelänge ca. {fmt_time(segments[-1]['end'])}"
    doc.add_paragraph(info)
    note = doc.add_paragraph(
        "Automatisch mit lokaler KI erstellt. Die Notizen können Fehler enthalten – "
        "wichtige Stellen bitte im Transkript oder in den Unterlagen nachprüfen."
    )
    note.runs[0].italic = True

    doc.add_heading("Zusammenfassung", level=1)
    _add_markdown(doc, summary or "Keine Zusammenfassung erstellt.")

    doc.add_heading("Notizen nach Abschnitten", level=1)
    for index, part in enumerate(parts, start=1):
        doc.add_heading(f"Abschnitt {index} ({part['label']})", level=2)
        _add_markdown(doc, part.get("notes", ""))

    if segments:
        doc.add_page_break()
        doc.add_heading("Vollständiges Transkript", level=1)
        paragraph_start, buffer = segments[0]["start"], []
        for seg in segments:
            buffer.append(seg["text"])
            if seg["end"] - paragraph_start >= 60:
                doc.add_paragraph(f"[{fmt_time(paragraph_start)}] " + " ".join(buffer))
                paragraph_start, buffer = seg["end"], []
        if buffer:
            doc.add_paragraph(f"[{fmt_time(paragraph_start)}] " + " ".join(buffer))

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def safe_filename(text: str) -> str:
    text = re.sub(r"[^\w\- ]+", "", text, flags=re.UNICODE).strip().replace(" ", "_")
    return text[:60] or "Vorlesung"


AI = AIConfig()
_SERVER = ModelServer()


_START_LOCK = threading.Lock()


def start_ai(model: Path, context: int) -> None:
    """Startet den KI-Server mit dem gewählten Modell (kann einige Minuten dauern).

    Läuft er schon mit diesem Modell, kehrt der Aufruf sofort zurück; lädt er gerade, wird gewartet.
    """
    if model.name == BUNDLED_MODEL:
        context = 2048
    with _START_LOCK:
        url = _SERVER.ensure(model, context)
        AI.configure(url, context, model_label(model))


def stop_ai() -> None:
    _SERVER.stop()


def ai_running() -> bool:
    return _SERVER.process is not None and _SERVER.process.poll() is None


def save_session(title: str, created: dt.datetime, segments: list[dict], parts: list[dict],
                 summary: str) -> Path | None:
    """Speichert Transkript und Notizen, damit später z. B. Unterrichtsmaterial dazukommen kann."""
    if not segments:
        return None
    try:
        SESSIONS_DIR.mkdir(parents=True, exist_ok=True)
        path = SESSIONS_DIR / f"{created:%Y-%m-%d_%H%M%S}_{safe_filename(title)}.json"
        data = {"title": title, "created": created.isoformat(timespec="seconds"),
                "date": created.date().isoformat(), "duration": segments[-1]["end"],
                "segments": segments, "summary": summary,
                "parts": [{k: p.get(k, "") for k in ("label", "text", "notes")} for p in parts]}
        path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        return path
    except OSError:
        return None


class Auswerter:
    """Wertet fertige Aufnahmeteile nacheinander im Hintergrund aus und erstellt am Ende das Word-Dokument.

    Der Zustand (parts, status, result, error) wird von der Oberfläche regelmäßig gelesen.
    """

    def __init__(self) -> None:
        self._queue: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._whisper = {}
        self.whisper_model = "base"
        self.ai_model: Path | None = None
        self.ai_context = 8192
        self.limit_threads = lambda: False  # True, solange eine Aufnahme läuft
        self.on_part_done: Callable[[str, dict], None] | None = None   # z. B. in Discord teilen
        self.on_finished: Callable[[str, str, Path], None] | None = None
        self.reset("")

    def reset(self, title: str) -> None:
        self.session_id = time.time_ns()
        self.title = title
        self.parts: list[dict] = []
        self.closed = False
        self.status = ""
        self.error = ""
        self.result: dict | None = None

    @property
    def busy(self) -> bool:
        if any(p["state"] in ("wartet", "läuft") for p in self.parts):
            return True
        return self.closed and self.result is None and not self.error

    def submit(self, path: Path, offset: float, number: int) -> None:
        part = {"number": number, "path": Path(path), "offset": offset, "state": "wartet",
                "status": "Wartet auf Auswertung …", "segments": [], "chunks": []}
        self.parts.append(part)
        self._queue.put(("part", self.session_id, part))
        self._ensure_thread()

    def finish(self) -> None:
        self.closed = True
        self._queue.put(("final", self.session_id, None))
        self._ensure_thread()

    def _ensure_thread(self) -> None:
        if self._thread is None or not self._thread.is_alive():
            self._thread = threading.Thread(target=self._work, name="auswerter", daemon=True)
            self._thread.start()

    def _model(self):
        if self.whisper_model not in self._whisper:
            import whisper  # erst hier laden: dauert einige Sekunden
            self._whisper[self.whisper_model] = whisper.load_model(
                self.whisper_model, download_root=str(WHISPER_MODEL_DIR))
        return self._whisper[self.whisper_model]

    def _limit_cpu(self) -> None:
        """Während der Vorlesung nur die Hälfte der Kerne nutzen, damit LearnSpace3D flüssig bleibt."""
        try:
            import torch
            cores = os.cpu_count() or 2
            torch.set_num_threads(max(1, cores // 2) if self.limit_threads() else cores)
        except Exception:
            pass

    def _ensure_ai(self, report: Callable[[str], None]) -> None:
        if self.ai_model is None:
            raise RuntimeError("Kein KI-Modell eingestellt (Einstellungen → KI-Modell).")
        if not _SERVER.running_with(self.ai_model, 2048 if self.ai_model.name == BUNDLED_MODEL else self.ai_context):
            report("Wartet auf die KI (wird geladen) …")
        start_ai(self.ai_model, self.ai_context)

    def _work(self) -> None:
        while True:
            kind, session_id, part = self._queue.get()
            if session_id != self.session_id:
                continue
            try:
                if kind == "part":
                    self._do_part(part)
                else:
                    self._do_final()
            except Exception as exc:
                if part is not None:
                    part["state"], part["status"] = "fehler", f"Fehler: {exc}"
                else:
                    self.error = f"Die Endfassung konnte nicht erstellt werden: {exc}"

    def _do_part(self, part: dict) -> None:
        part["state"], part["status"] = "läuft", "Spracherkennung läuft …"
        self._limit_cpu()
        segments = transcribe_with(self._model(), str(part["path"]))
        for seg in segments:
            seg["start"] += part["offset"]
            seg["end"] += part["offset"]
        part["segments"] = segments
        self._ensure_ai(lambda t: part.__setitem__("status", t))
        chunks = chunk_segments(segments)
        write_notes(chunks, self.title, lambda v, t: part.__setitem__("status", f"KI: {t}"))
        part["chunks"] = chunks
        part["state"] = "fertig"
        part["status"] = "Fertig" if segments else "Fertig (keine Sprache erkannt)"
        if self.on_part_done:
            self.on_part_done(self.title, part)

    def _do_final(self) -> None:
        segments = [s for p in self.parts for s in p["segments"]]
        chunks = [c for p in self.parts for c in p["chunks"]]
        if not chunks:
            self.error = "In der Aufnahme wurde keine Sprache erkannt."
            return
        self._ensure_ai(lambda t: setattr(self, "status", t))
        summary = summarize_parts(chunks, self.title, lambda v, t: setattr(self, "status", t))
        created = dt.datetime.now()
        docx_bytes = build_docx(self.title, summary, chunks, segments, created)
        NOTES_DIR.mkdir(parents=True, exist_ok=True)
        path = NOTES_DIR / f"{created:%Y-%m-%d_%H%M}_{safe_filename(self.title)}.docx"
        path.write_bytes(docx_bytes)
        save_session(self.title, created, segments, chunks, summary)
        self.status = "Fertig."
        self.result = {"path": path, "summary": summary}
        if self.on_finished:
            self.on_finished(self.title, summary, path)
