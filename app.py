# -*- coding: utf-8 -*-
"""Gesprächsnotiz-AI – Vorlesungen (z. B. LearnSpace3D) aufnehmen und als Lernnotizen zusammenfassen.

Diese Datei ersetzt  GespraechsnotizAI\\_internal\\app.py  im entpackten Programmordner.
Sie benötigt keine zusätzlichen Bibliotheken: Der Systemton wird direkt über die
Windows-Audio-Schnittstelle (WASAPI-Loopback) aufgenommen.
"""
from __future__ import annotations

import ctypes
import datetime as dt
import io
import json
import os
import re
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
import uuid
import wave
from pathlib import Path
from typing import Callable, Iterable

import numpy as np
import streamlit as st
import whisper
from docx import Document
from docx.shared import Pt
from pypdf import PdfReader

# ---------------------------------------------------------------------------
# Einstellungen
# ---------------------------------------------------------------------------
APP_TITLE = "Gesprächsnotiz-AI – Vorlesungsnotizen"
WHISPER_MODEL_DIR = Path(__file__).resolve().parent / "runtime" / "whisper"
LOCAL_AI_URL = "http://127.0.0.1:9931/v1/chat/completions"
SUPPORTED_AUDIO = ("wav", "mp3", "m4a", "mpeg", "mpga", "webm", "ogg")
MAX_DOCUMENT_CHARS = 120_000

# Der mitgelieferte KI-Server läuft mit nur 2048 Tokens Kontext.
# Deshalb werden lange Texte in kleine Abschnitte zerlegt.
CHUNK_CHARS = 2400
NOTES_MAX_TOKENS = 450
SUMMARY_MAX_TOKENS = 600

RECORD_RATE = 16_000  # Whisper arbeitet intern mit 16 kHz mono

DATA_DIR = Path.home() / "Documents" / "GespraechsnotizAI"
RECORDINGS_DIR = DATA_DIR / "Aufnahmen"
NOTES_DIR = DATA_DIR / "Notizen"

st.set_page_config(page_title="Gesprächsnotiz-AI", page_icon="🎓", layout="wide")


# ---------------------------------------------------------------------------
# Systemton-Aufnahme (Windows WASAPI-Loopback, nur ctypes)
# ---------------------------------------------------------------------------
class _GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_uint32),
        ("Data2", ctypes.c_uint16),
        ("Data3", ctypes.c_uint16),
        ("Data4", ctypes.c_ubyte * 8),
    ]

    @classmethod
    def parse(cls, text: str) -> "_GUID":
        guid = cls()
        ctypes.memmove(ctypes.byref(guid), uuid.UUID(text).bytes_le, 16)
        return guid


class _WAVEFORMATEX(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ("wFormatTag", ctypes.c_uint16),
        ("nChannels", ctypes.c_uint16),
        ("nSamplesPerSec", ctypes.c_uint32),
        ("nAvgBytesPerSec", ctypes.c_uint32),
        ("nBlockAlign", ctypes.c_uint16),
        ("wBitsPerSample", ctypes.c_uint16),
        ("cbSize", ctypes.c_uint16),
    ]


class _WAVEFORMATEXTENSIBLE(ctypes.Structure):
    _pack_ = 1
    _fields_ = [
        ("Format", _WAVEFORMATEX),
        ("wValidBitsPerSample", ctypes.c_uint16),
        ("dwChannelMask", ctypes.c_uint32),
        ("SubFormat", _GUID),
    ]


_CLSID_MMDeviceEnumerator = "BCDE0395-E52F-467C-8E3D-C4579291692E"
_IID_IMMDeviceEnumerator = "A95664D2-9614-4F35-A746-DE8DB63617E6"
_IID_IAudioClient = "1CB9AD4C-DBFA-4C32-B178-C2F568A703B2"
_IID_IAudioCaptureClient = "C8ADBD64-E71E-48A0-A4DE-185C395CD317"
_CLSCTX_ALL = 23
_AUDCLNT_STREAMFLAGS_LOOPBACK = 0x00020000
_AUDCLNT_BUFFERFLAGS_SILENT = 0x2
_WAVE_FORMAT_PCM = 1
_WAVE_FORMAT_IEEE_FLOAT = 3
_WAVE_FORMAT_EXTENSIBLE = 0xFFFE


def _com_method(obj: ctypes.c_void_p, index: int, *argtypes):
    vtable = ctypes.cast(obj, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
    prototype = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, *argtypes)
    return prototype(vtable[index])


def _check(hr: int, what: str) -> None:
    if hr < 0:
        raise OSError(f"{what} fehlgeschlagen (Fehlercode 0x{hr & 0xFFFFFFFF:08X})")


def _release(obj: ctypes.c_void_p) -> None:
    if obj and obj.value:
        _com_method(obj, 2)(obj)


class _Resampler:
    """Wandelt einen beliebigen Mono-Datenstrom fortlaufend in 16 kHz um."""

    def __init__(self, source_rate: int):
        self.step = source_rate / RECORD_RATE
        self.box = max(1, int(round(self.step)))
        self.carry = np.zeros(0, dtype=np.float32)
        self.pos = 0.0

    def __call__(self, samples: np.ndarray) -> np.ndarray:
        if self.box > 1:  # einfacher Tiefpass gegen Verzerrungen
            samples = np.convolve(samples, np.ones(self.box) / self.box, mode="same")
        buf = np.concatenate([self.carry, samples.astype(np.float32)])
        if len(buf) < 2:
            self.carry = buf
            return np.zeros(0, dtype=np.float32)
        positions = np.arange(self.pos, len(buf) - 1, self.step)
        out = np.interp(positions, np.arange(len(buf)), buf)
        next_pos = (positions[-1] + self.step) if len(positions) else self.pos
        consumed = int(next_pos)
        self.carry = buf[consumed:]
        self.pos = next_pos - consumed
        return out.astype(np.float32)


class LoopbackRecorder:
    """Nimmt alles auf, was am Standard-Ausgabegerät (Lautsprecher/Headset) abgespielt wird."""

    def __init__(self) -> None:
        self._thread: threading.Thread | None = None
        self._stop = threading.Event()
        self.paused = False
        self.path: Path | None = None
        self.started_at = 0.0
        self.samples_written = 0
        self.level = 0.0
        self.error = ""
        self.device_format = ""

    # -- Status -------------------------------------------------------------
    @property
    def running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    @property
    def recorded_seconds(self) -> float:
        return self.samples_written / RECORD_RATE

    @property
    def wall_seconds(self) -> float:
        return time.time() - self.started_at if self.started_at else 0.0

    # -- Steuerung ----------------------------------------------------------
    def start(self, path: Path) -> None:
        if self.running:
            return
        path.parent.mkdir(parents=True, exist_ok=True)
        self.path = path
        self.samples_written = 0
        self.level = 0.0
        self.error = ""
        self.paused = False
        self.started_at = time.time()
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, name="loopback-recorder", daemon=True)
        self._thread.start()

    def stop(self) -> Path | None:
        self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=10)
        self._thread = None
        self.started_at = 0.0
        return self.path

    # -- Aufnahme-Thread ----------------------------------------------------
    def _run(self) -> None:
        ole32 = ctypes.windll.ole32
        ole32.CoInitializeEx(None, 0)
        try:
            with wave.open(str(self.path), "wb") as wav:
                wav.setnchannels(1)
                wav.setsampwidth(2)
                wav.setframerate(RECORD_RATE)
                while not self._stop.is_set():
                    try:
                        self._capture(ole32, wav)
                    except Exception as exc:  # z. B. Headset ab-/angesteckt
                        self.error = f"{exc} – neuer Versuch …"
                        self._stop.wait(1.0)
        except Exception as exc:
            self.error = f"Aufnahme konnte nicht gespeichert werden: {exc}"
        finally:
            ole32.CoUninitialize()

    def _capture(self, ole32, wav: wave.Wave_write) -> None:
        enumerator, device, client, capture = (ctypes.c_void_p() for _ in range(4))
        mix_format = ctypes.POINTER(_WAVEFORMATEX)()
        try:
            clsid = _GUID.parse(_CLSID_MMDeviceEnumerator)
            iid_enum = _GUID.parse(_IID_IMMDeviceEnumerator)
            _check(
                ole32.CoCreateInstance(ctypes.byref(clsid), None, _CLSCTX_ALL,
                                       ctypes.byref(iid_enum), ctypes.byref(enumerator)),
                "Audiogeräte öffnen",
            )
            # IMMDeviceEnumerator::GetDefaultAudioEndpoint(eRender, eConsole)
            _check(_com_method(enumerator, 4, ctypes.c_int, ctypes.c_int, ctypes.POINTER(ctypes.c_void_p))(
                enumerator, 0, 0, ctypes.byref(device)), "Standard-Lautsprecher finden")
            iid_client = _GUID.parse(_IID_IAudioClient)
            # IMMDevice::Activate
            _check(_com_method(device, 3, ctypes.POINTER(_GUID), ctypes.c_uint32, ctypes.c_void_p,
                               ctypes.POINTER(ctypes.c_void_p))(
                device, ctypes.byref(iid_client), _CLSCTX_ALL, None, ctypes.byref(client)),
                "Audio-Client starten")
            # IAudioClient::GetMixFormat
            _check(_com_method(client, 8, ctypes.POINTER(ctypes.POINTER(_WAVEFORMATEX)))(
                client, ctypes.byref(mix_format)), "Audioformat lesen")
            fmt = mix_format.contents
            channels, rate = fmt.nChannels, fmt.nSamplesPerSec
            bits, block_align = fmt.wBitsPerSample, fmt.nBlockAlign
            tag = fmt.wFormatTag
            if tag == _WAVE_FORMAT_EXTENSIBLE:
                ext = ctypes.cast(mix_format, ctypes.POINTER(_WAVEFORMATEXTENSIBLE)).contents
                tag = ext.SubFormat.Data1
            if tag not in (_WAVE_FORMAT_PCM, _WAVE_FORMAT_IEEE_FLOAT):
                raise OSError(f"Nicht unterstütztes Audioformat ({tag})")
            self.device_format = f"{rate} Hz, {channels} Kanäle, {bits} Bit"

            # IAudioClient::Initialize (geteilt, Loopback, 1 s Puffer)
            _check(_com_method(client, 3, ctypes.c_int, ctypes.c_uint32, ctypes.c_longlong,
                               ctypes.c_longlong, ctypes.POINTER(_WAVEFORMATEX), ctypes.c_void_p)(
                client, 0, _AUDCLNT_STREAMFLAGS_LOOPBACK, 10_000_000, 0, mix_format, None),
                "Loopback-Aufnahme vorbereiten")
            iid_capture = _GUID.parse(_IID_IAudioCaptureClient)
            _check(_com_method(client, 14, ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p))(
                client, ctypes.byref(iid_capture), ctypes.byref(capture)), "Aufnahmedienst holen")
            _check(_com_method(client, 10)(client), "Aufnahme starten")

            get_next = _com_method(capture, 5, ctypes.POINTER(ctypes.c_uint32))
            get_buffer = _com_method(capture, 3, ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),
                                     ctypes.POINTER(ctypes.c_uint32), ctypes.POINTER(ctypes.c_uint32),
                                     ctypes.c_void_p, ctypes.c_void_p)
            release_buffer = _com_method(capture, 4, ctypes.c_uint32)

            resample = _Resampler(rate)
            pending: list[np.ndarray] = []
            pending_len = 0
            self.error = ""

            while not self._stop.is_set():
                self._stop.wait(0.05)
                packet = ctypes.c_uint32()
                _check(get_next(capture, ctypes.byref(packet)), "Audiodaten abfragen")
                while packet.value:
                    data = ctypes.POINTER(ctypes.c_ubyte)()
                    frames = ctypes.c_uint32()
                    flags = ctypes.c_uint32()
                    _check(get_buffer(capture, ctypes.byref(data), ctypes.byref(frames),
                                      ctypes.byref(flags), None, None), "Audiodaten lesen")
                    count = frames.value
                    if flags.value & _AUDCLNT_BUFFERFLAGS_SILENT or not data:
                        raw = None
                    else:
                        raw = ctypes.string_at(data, count * block_align)
                    release_buffer(capture, count)

                    if raw is None:
                        samples = np.zeros(count, dtype=np.float32)
                    else:
                        samples = _to_mono_float(raw, tag, bits, channels)
                    if not self.paused and len(samples):
                        out = resample(samples)
                        self.level = float(np.sqrt(np.mean(samples ** 2))) if len(samples) else 0.0
                        pending.append(out)
                        pending_len += len(out)
                    elif self.paused:
                        self.level = 0.0
                    _check(get_next(capture, ctypes.byref(packet)), "Audiodaten abfragen")

                if pending_len >= RECORD_RATE // 2:
                    self._write(wav, pending)
                    pending, pending_len = [], 0

            if pending:
                self._write(wav, pending)
            _com_method(client, 11)(client)  # IAudioClient::Stop
        finally:
            if mix_format:
                ole32.CoTaskMemFree(mix_format)
            for obj in (capture, client, device, enumerator):
                _release(obj)

    def _write(self, wav: wave.Wave_write, blocks: list[np.ndarray]) -> None:
        audio = np.concatenate(blocks)
        pcm = (np.clip(audio, -1.0, 1.0) * 32767).astype("<i2")
        wav.writeframes(pcm.tobytes())  # Dateikopf wird dabei laufend aktualisiert
        self.samples_written += len(pcm)


def _to_mono_float(raw: bytes, tag: int, bits: int, channels: int) -> np.ndarray:
    if tag == _WAVE_FORMAT_IEEE_FLOAT and bits == 32:
        data = np.frombuffer(raw, dtype="<f4")
    elif bits == 16:
        data = np.frombuffer(raw, dtype="<i2").astype(np.float32) / 32768.0
    elif bits == 32:
        data = np.frombuffer(raw, dtype="<i4").astype(np.float32) / 2147483648.0
    elif bits == 24:
        b = np.frombuffer(raw, dtype=np.uint8).reshape(-1, 3).astype(np.int32)
        ints = (b[:, 0] | (b[:, 1] << 8) | (b[:, 2] << 16))
        ints = np.where(ints >= 1 << 23, ints - (1 << 24), ints)
        data = ints.astype(np.float32) / 8388608.0
    else:
        raise OSError(f"Nicht unterstützte Bittiefe: {bits}")
    usable = len(data) - len(data) % channels
    return data[:usable].reshape(-1, channels).mean(axis=1)


@st.cache_resource
def get_recorder() -> LoopbackRecorder:
    """Ein Rekorder pro laufendem Programm – bleibt auch beim Neuladen der Seite erhalten."""
    return LoopbackRecorder()


# ---------------------------------------------------------------------------
# Spracherkennung
# ---------------------------------------------------------------------------
@st.cache_resource(show_spinner="Whisper-Modell wird geladen …")
def load_whisper_model(model_name: str):
    return whisper.load_model(model_name, download_root=str(WHISPER_MODEL_DIR))


_HALLUCINATIONS = re.compile(r"untertitel|amara\.org|copyright|www\.", re.IGNORECASE)


def transcribe_file(path: str, model_name: str) -> list[dict]:
    """Gibt eine Liste von Abschnitten mit Start-/Endzeit und Text zurück."""
    model = load_whisper_model(model_name)
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


def transcribe_bytes(audio_bytes: bytes, filename: str, model_name: str) -> list[dict]:
    suffix = Path(filename).suffix.lower()
    if suffix.lstrip(".") not in SUPPORTED_AUDIO:
        suffix = ".wav"
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(audio_bytes)
            temp_path = tmp.name
        return transcribe_file(temp_path, model_name)
    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)


# ---------------------------------------------------------------------------
# Dokumente
# ---------------------------------------------------------------------------
def extract_pdf(file) -> str:
    reader = PdfReader(file)
    pages = []
    for number, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if text.strip():
            pages.append(f"[Seite {number}]\n{text}")
    return "\n\n".join(pages)


def extract_docx(file) -> str:
    doc = Document(file)
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for index, table in enumerate(doc.tables, start=1):
        rows = [" | ".join(cell.text.strip() for cell in row.cells) for row in table.rows]
        if rows:
            parts.append(f"[Tabelle {index}]\n" + "\n".join(rows))
    return "\n\n".join(parts)


def extract_documents(files: Iterable) -> tuple[str, list[str]]:
    contents: list[str] = []
    errors: list[str] = []
    total = 0
    for uploaded in files or []:
        try:
            suffix = Path(uploaded.name).suffix.lower()
            if suffix == ".pdf":
                text = extract_pdf(uploaded)
            elif suffix == ".docx":
                text = extract_docx(uploaded)
            else:
                errors.append(f"{uploaded.name}: Dateityp wird nicht unterstützt.")
                continue
            if not text.strip():
                errors.append(f"{uploaded.name}: Kein Text gefunden. Scan-PDFs brauchen zuerst eine Texterkennung (OCR).")
                continue
            remaining = MAX_DOCUMENT_CHARS - total
            if remaining <= 0:
                errors.append("Weitere Dokumente wurden wegen des Textlimits ausgelassen.")
                break
            if len(text) > remaining:
                text = text[:remaining]
                errors.append("Der Dokumenttext wurde auf 120.000 Zeichen begrenzt.")
            contents.append(f"{uploaded.name}:\n{text}")
            total += len(text)
        except Exception as exc:
            errors.append(f"{uploaded.name}: {exc}")
    return "\n\n".join(contents), errors


# ---------------------------------------------------------------------------
# Lokale KI (llama.cpp-Server)
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = (
    "Du bist ein Lernassistent für Teilnehmende einer Umschulung. Du schreibst klare, "
    "gut lesbare Lernnotizen auf Deutsch. Du hältst dich streng an den gegebenen Text "
    "und erfindest nichts dazu. Unverständliche Stellen lässt du weg."
)


def ask_ai(prompt: str, max_tokens: int) -> str:
    payload = {
        "model": "local-model",
        "temperature": 0.2,
        "max_tokens": max_tokens,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ],
    }
    request = urllib.request.Request(
        LOCAL_AI_URL,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=900) as response:
        result = json.loads(response.read().decode("utf-8"))
    return (result["choices"][0]["message"].get("content") or "").strip()


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


def split_text(text: str, limit: int = CHUNK_CHARS) -> list[str]:
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


def chunk_segments(segments: list[dict], limit: int = CHUNK_CHARS) -> list[dict]:
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


def create_notes(
    segments: list[dict], document_text: str, title: str,
    progress: Callable[[float, str], None],
) -> tuple[str, list[dict]]:
    """Fasst abschnittsweise zusammen und erstellt anschließend eine Kurzfassung."""
    topic = f" zum Thema „{title}“" if title else ""
    parts = chunk_segments(segments)
    for doc_chunk in split_text(document_text) if document_text else []:
        parts.append({"label": "Begleitdokument", "text": doc_chunk})
    if not parts:
        return "", []

    total_steps = len(parts) + 1
    for index, part in enumerate(parts, start=1):
        progress((index - 1) / total_steps, f"Abschnitt {index} von {len(parts)} wird zusammengefasst …")
        source = "Ausschnitt aus einem Begleitdokument" if part["label"] == "Begleitdokument" \
            else "Ausschnitt aus dem Transkript einer Vorlesung"
        part["notes"] = ask_ai_safe(
            lambda text, source=source: (
                f"{source}{topic}:\n\n{text}\n\n"
                "Schreibe dazu kurze Lernnotizen als Stichpunkte:\n"
                "- die wichtigsten Inhalte\n"
                "- Fachbegriffe mit kurzer Erklärung\n"
                "- Aufgaben, Termine oder Hinweise der Lehrkraft (nur wenn sie genannt werden)"
            ),
            part["text"], NOTES_MAX_TOKENS,
        ) or "(Keine Notizen erzeugt.)"

    # Notizen so lange verdichten, bis sie in eine einzige Anfrage passen
    combined = "\n\n".join(p["notes"] for p in parts)
    rounds = 0
    while len(combined) > CHUNK_CHARS and rounds < 4:
        rounds += 1
        progress(len(parts) / total_steps, f"Notizen werden verdichtet (Runde {rounds}) …")
        combined = "\n\n".join(
            ask_ai_safe(
                lambda text: (
                    f"Lernnotizen zu Teilen einer Vorlesung{topic}:\n\n{text}\n\n"
                    "Fasse diese Notizen kürzer zusammen. Behalte Fachbegriffe, Aufgaben und Termine."
                ),
                block, NOTES_MAX_TOKENS,
            )
            for block in split_text(combined)
        )

    progress(len(parts) / total_steps, "Kurzfassung der ganzen Vorlesung wird erstellt …")
    summary = ask_ai_safe(
        lambda text: (
            f"Lernnotizen zu einer kompletten Vorlesung{topic}:\n\n{text}\n\n"
            "Erstelle daraus eine Zusammenfassung mit diesen Überschriften:\n"
            "## Worum ging es?\n(2–3 Sätze)\n"
            "## Die wichtigsten Punkte\n(Stichpunkte)\n"
            "## Fachbegriffe\n(Begriff: kurze Erklärung)\n"
            "## Aufgaben und Termine\n(nur wenn genannt, sonst „Keine genannt“)"
        ),
        combined[: CHUNK_CHARS + 200], SUMMARY_MAX_TOKENS,
    )
    progress(1.0, "Fertig.")
    return summary, parts


# ---------------------------------------------------------------------------
# Word-Dokument
# ---------------------------------------------------------------------------
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


# ---------------------------------------------------------------------------
# Verarbeitung
# ---------------------------------------------------------------------------
def run_pipeline(*, audio_path: str | None = None, audio_bytes: bytes | None = None,
                 audio_name: str = "aufnahme.wav", docs=None, title: str = "",
                 model_name: str = "base") -> None:
    segments: list[dict] = []
    document_text = ""
    with st.status("Auswertung läuft – bitte das Fenster offen lassen …", expanded=True) as status:
        if audio_path or audio_bytes:
            st.write("🎧 Audio wird transkribiert. Bei langen Vorlesungen kann das einige Minuten dauern …")
            try:
                if audio_path:
                    segments = transcribe_file(audio_path, model_name)
                else:
                    segments = transcribe_bytes(audio_bytes, audio_name, model_name)
            except Exception as exc:
                status.update(label="Transkription fehlgeschlagen", state="error")
                st.error("Die Aufnahme konnte nicht transkribiert werden.")
                st.exception(exc)
                return
            if segments:
                st.write(f"✅ Transkript fertig ({fmt_time(segments[-1]['end'])} Audio).")
            else:
                st.warning("In der Aufnahme wurde keine Sprache erkannt.")

        if docs:
            st.write("📄 Dokumente werden ausgelesen …")
            document_text, errors = extract_documents(docs)
            for error in errors:
                st.warning(error)

        if not segments and not document_text:
            status.update(label="Nichts zum Auswerten gefunden", state="error")
            return

        st.write("🧠 Die lokale KI schreibt Notizen …")
        bar = st.progress(0.0)
        try:
            summary, parts = create_notes(segments, document_text, title,
                                          lambda value, text: bar.progress(min(value, 1.0), text=text))
        except Exception as exc:
            status.update(label="KI-Auswertung fehlgeschlagen", state="error")
            st.error("Die lokale KI-Auswertung ist fehlgeschlagen. Läuft das Konsolenfenster noch?")
            st.exception(exc)
            return

        created = dt.datetime.now()
        docx_bytes = build_docx(title, summary, parts, segments, created)
        filename = f"{created:%Y-%m-%d_%H%M}_{safe_filename(title)}.docx"
        saved_to = ""
        try:
            NOTES_DIR.mkdir(parents=True, exist_ok=True)
            (NOTES_DIR / filename).write_bytes(docx_bytes)
            saved_to = str(NOTES_DIR / filename)
        except OSError as exc:
            st.warning(f"Konnte die Notizen nicht automatisch speichern: {exc}")
        status.update(label="Fertig! Deine Notizen sind unten.", state="complete", expanded=False)

    st.session_state["result"] = {
        "title": title,
        "summary": summary,
        "parts": parts,
        "transcript": "\n".join(f"[{fmt_time(s['start'])}] {s['text']}" for s in segments),
        "docx": docx_bytes,
        "filename": filename,
        "saved_to": saved_to,
    }


# ---------------------------------------------------------------------------
# Oberfläche
# ---------------------------------------------------------------------------
st.title("🎓 Gesprächsnotiz-AI")
st.write(
    "Nimmt deine Vorlesung im **LearnSpace3D** (oder jeden anderen Ton am PC) auf und "
    "macht daraus automatisch **Lernnotizen als Word-Dokument** – komplett lokal auf deinem Rechner."
)

with st.sidebar:
    st.header("⚙️ Einstellungen")
    st.write("**Lokales KI-Modell:** Qwen 2.5 0.5B")
    whisper_model = st.selectbox(
        "Spracherkennung (Whisper)", ["base", "small"], index=0,
        help="„base“ ist mitgeliefert. „small“ erkennt deutlich genauer, ist aber langsamer "
             "und wird beim ersten Mal einmalig aus dem Internet geladen (ca. 460 MB).",
    )
    st.caption(f"Aufnahmen: `{RECORDINGS_DIR}`")
    st.caption(f"Notizen: `{NOTES_DIR}`")

lecture_tab, file_tab = st.tabs(["🎓 Vorlesung aufnehmen", "📁 Datei auswerten"])

recorder = get_recorder()

with lecture_tab:
    if not sys.platform.startswith("win"):
        st.error("Die Systemton-Aufnahme funktioniert nur unter Windows.")
    else:
        title = st.text_input("Thema / Fach der Vorlesung", key="lecture_title",
                              placeholder="z. B. Netzwerktechnik – Subnetting")
        st.markdown(
            "**So geht's:** 1. LearnSpace3D öffnen und der Vorlesung beitreten · "
            "2. hier auf **Aufnahme starten** klicken · 3. am Ende **Stoppen & Notizen erstellen**.\n\n"
            "Aufgenommen wird alles, was über deine Lautsprecher bzw. dein Headset läuft. "
            "Schalte andere Töne (Musik, Videos) während der Vorlesung aus oder nutze **Pause**."
        )
        consent = st.checkbox(
            "Ich habe die Erlaubnis der Lehrkraft bzw. der WBS, die Vorlesung für meine "
            "persönlichen Lernnotizen aufzunehmen, und gebe die Aufnahme nicht weiter.",
            key="consent",
        )

        @st.fragment(run_every=1.0 if recorder.running else None)
        def recorder_panel() -> None:
            if recorder.running:
                state = "⏸️ Pausiert" if recorder.paused else "🔴 Aufnahme läuft"
                st.subheader(f"{state} – {fmt_time(recorder.wall_seconds)}")
                st.progress(min(recorder.level * 4, 1.0),
                            text="Lautstärke (bewegt sich, wenn Ton ankommt)")
                st.caption(f"Aufgenommener Ton: {fmt_time(recorder.recorded_seconds)} · "
                           f"{recorder.device_format} · Datei: {recorder.path}")
                if recorder.error:
                    st.warning(recorder.error)
                col_pause, col_stop = st.columns(2)
                if col_pause.button("▶️ Weiter" if recorder.paused else "⏸️ Pause", width="stretch"):
                    recorder.paused = not recorder.paused
                    st.rerun()
                if col_stop.button("⏹️ Stoppen & Notizen erstellen", type="primary", width="stretch"):
                    path = recorder.stop()
                    st.session_state["process_recording"] = str(path) if path else None
                    st.rerun()
            else:
                if st.button("🔴 Aufnahme starten", type="primary", width="stretch", disabled=not consent):
                    name = f"{dt.datetime.now():%Y-%m-%d_%H%M}_{safe_filename(st.session_state.get('lecture_title', ''))}.wav"
                    recorder.start(RECORDINGS_DIR / name)
                    st.rerun()
                if not consent:
                    st.caption("Bitte zuerst die Erlaubnis bestätigen.")

        recorder_panel()

        pending = st.session_state.pop("process_recording", None)
        if pending:
            if recorder.error and recorder.samples_written == 0:
                st.error(f"Es wurde kein Ton aufgenommen: {recorder.error.replace(' – neuer Versuch …', '')}")
            elif recorder.samples_written < RECORD_RATE * 2:
                st.warning("Die Aufnahme ist leer oder kürzer als 2 Sekunden. Lief Ton über den Lautsprecher?")
            else:
                run_pipeline(audio_path=pending, title=st.session_state.get("lecture_title", ""),
                             model_name=whisper_model)

        with st.expander("Ältere Aufnahme auswerten (z. B. nach einem Absturz)"):
            recordings = sorted(RECORDINGS_DIR.glob("*.wav"), reverse=True) if RECORDINGS_DIR.exists() else []
            if not recordings:
                st.caption("Noch keine gespeicherten Aufnahmen.")
            else:
                chosen = st.selectbox("Aufnahme", recordings, format_func=lambda p: p.name)
                if st.button("Diese Aufnahme auswerten", disabled=recorder.running):
                    run_pipeline(audio_path=str(chosen), title=st.session_state.get("lecture_title", ""),
                                 model_name=whisper_model)

with file_tab:
    file_title = st.text_input("Thema (optional)", key="file_title")
    audio_col, docs_col = st.columns(2)
    with audio_col:
        st.subheader("Gespräch / Audio")
        mic_audio = st.audio_input("Direkt über das Mikrofon aufnehmen")
        uploaded_audio = st.file_uploader("Oder Audiodatei auswählen", type=list(SUPPORTED_AUDIO))
    with docs_col:
        st.subheader("Begleitende Dokumente")
        docs = st.file_uploader("PDF- oder Word-Dateien (z. B. Folien, Skript)",
                                type=["pdf", "docx"], accept_multiple_files=True)
        st.caption("Mehrere Dokumente sind möglich. Scan-PDFs brauchen eine Texterkennung (OCR).")
    if st.button("🚀 Auswerten", type="primary", width="stretch"):
        selected = mic_audio if mic_audio is not None else uploaded_audio
        if selected is None and not docs:
            st.error("Bitte nimm etwas auf oder lade mindestens eine Datei hoch.")
        else:
            run_pipeline(
                audio_bytes=selected.getvalue() if selected is not None else None,
                audio_name=getattr(selected, "name", "aufnahme.wav") if selected is not None else "",
                docs=docs, title=file_title, model_name=whisper_model,
            )

result = st.session_state.get("result")
if result:
    st.divider()
    st.header("📝 Deine Notizen")
    st.download_button("⬇️ Notizen als Word-Dokument herunterladen", data=result["docx"],
                       file_name=result["filename"], type="primary", width="stretch",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    if result["saved_to"]:
        st.caption(f"Automatisch gespeichert unter: `{result['saved_to']}`")
    st.markdown(result["summary"] or "_Keine Zusammenfassung erstellt._")
    with st.expander("Notizen nach Abschnitten"):
        for index, part in enumerate(result["parts"], start=1):
            st.markdown(f"**Abschnitt {index} ({part['label']})**\n\n{part.get('notes', '')}")
    if result["transcript"]:
        with st.expander("Vollständiges Transkript"):
            st.text(result["transcript"])

with st.expander("Datenschutz und Hinweise"):
    st.markdown(
        "- Aufnahme, Spracherkennung und Zusammenfassung laufen **komplett lokal** auf diesem PC. Es wird nichts ins Internet geschickt.\n"
        "- Nimm Vorlesungen nur mit Erlaubnis auf und gib die Aufnahmen nicht weiter. Das heimliche Aufnehmen nicht öffentlicher Gespräche ist in Deutschland strafbar (§ 201 StGB).\n"
        "- Die Aufnahmen bleiben im Ordner „Aufnahmen“ gespeichert, bis du sie löschst (ca. 115 MB pro Stunde).\n"
        "- Das mitgelieferte KI-Modell ist sehr klein. Die Notizen sind eine Hilfe, ersetzen aber nicht das Nacharbeiten."
    )
