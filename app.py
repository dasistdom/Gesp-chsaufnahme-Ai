# -*- coding: utf-8 -*-
"""Gesprächsnotiz-AI – Vorlesungen (z. B. LearnSpace3D) aufnehmen und als Lernnotizen zusammenfassen.

Diese Datei ersetzt  GespraechsnotizAI\\_internal\\app.py  im entpackten Programmordner.
Sie benötigt keine zusätzlichen Bibliotheken: Der Systemton wird direkt über die
Windows-Audio-Schnittstelle (WASAPI-Loopback) aufgenommen.
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
import streamlit.components.v1 as components
import whisper
from docx import Document
from docx.shared import Pt
from pypdf import PdfReader

# ---------------------------------------------------------------------------
# Einstellungen
# ---------------------------------------------------------------------------
APP_TITLE = "Gesprächsnotiz-AI – Vorlesungsnotizen"
WHISPER_MODEL_DIR = Path(__file__).resolve().parent / "runtime" / "whisper"
LLAMA_DIR = Path(__file__).resolve().parent / "runtime" / "llama"
MODELS_DIR = LLAMA_DIR / "Modelle"
BUNDLED_MODEL = "qwen2.5-0.5b-instruct-q4_k_m.gguf"
PREFERRED_MODEL = "qwen3-8b"  # wird automatisch gewählt, wenn eine passende .gguf-Datei vorhanden ist
BUNDLED_AI_URL = "http://127.0.0.1:9931/v1/chat/completions"  # startet GespraechsnotizAI.exe
SUPPORTED_AUDIO = ("wav", "mp3", "m4a", "mpeg", "mpga", "webm", "ogg")
MAX_DOCUMENT_CHARS = 120_000


RECORD_RATE = 16_000  # Whisper arbeitet intern mit 16 kHz mono

DATA_DIR = Path.home() / "Documents" / "GespraechsnotizAI"
RECORDINGS_DIR = DATA_DIR / "Aufnahmen"
NOTES_DIR = DATA_DIR / "Notizen"

st.set_page_config(page_title="Gesprächsnotiz-AI", page_icon="🎓", layout="wide")


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


@st.cache_resource
def get_ai_config() -> AIConfig:
    return AIConfig()


AI = get_ai_config()


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


class _Timeline:
    """Gemeinsame Zeitachse (16 kHz), auf der Lautsprecher- und Mikrofonton gemischt werden."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self.base = 0
        self.buf = np.zeros(0, dtype=np.float32)

    def add(self, index: int, data: np.ndarray) -> None:
        with self._lock:
            if index < self.base:
                data = data[self.base - index:]
                index = self.base
            if not len(data):
                return
            start = index - self.base
            end = start + len(data)
            if end > len(self.buf):
                self.buf = np.concatenate([self.buf, np.zeros(end - len(self.buf), dtype=np.float32)])
            self.buf[start:end] += data

    def take(self, upto: int) -> np.ndarray:
        """Entnimmt alles bis zum Index `upto`; Lücken werden mit Stille gefüllt."""
        with self._lock:
            count = max(0, upto - self.base)
            if count > len(self.buf):
                self.buf = np.concatenate([self.buf, np.zeros(count - len(self.buf), dtype=np.float32)])
            out, self.buf = self.buf[:count], self.buf[count:]
            self.base += count
            return out

    @property
    def end(self) -> int:
        with self._lock:
            return self.base + len(self.buf)


def _clock_100ns() -> int:
    """Windows-Leistungszähler in 100-ns-Einheiten – dieselbe Uhr, die WASAPI für Zeitstempel nutzt."""
    kernel32 = ctypes.windll.kernel32
    counter, frequency = ctypes.c_int64(), ctypes.c_int64()
    kernel32.QueryPerformanceCounter(ctypes.byref(counter))
    kernel32.QueryPerformanceFrequency(ctypes.byref(frequency))
    return counter.value * 10_000_000 // frequency.value


_SOURCE_LABELS = {"speaker": "Lautsprecher/Headset", "mic": "Mikrofon"}


class LoopbackRecorder:
    """Nimmt den Ton des Standard-Ausgabegeräts auf – auf Wunsch zusammen mit dem Mikrofon.

    Die Aufnahme wird in Teile aufgeteilt (bei jeder Pause und optional alle X Minuten).
    Jeder fertige Teil wird an `on_part(pfad, startzeit_in_sekunden, nummer)` übergeben,
    damit er schon während der laufenden Vorlesung ausgewertet werden kann.
    """

    def __init__(self) -> None:
        self._thread: threading.Thread | None = None
        self._stop = threading.Event()
        self.paused = False
        self.with_mic = False
        self.base_path: Path | None = None
        self.path: Path | None = None
        self.part_number = 0
        self.auto_split_minutes = 0
        self.on_part: Callable[[Path, float, int], None] | None = None
        self.started_at = 0.0
        self.samples_written = 0
        self.part_samples = 0
        self.got_audio = False
        self.levels: dict[str, float] = {}
        self.errors: dict[str, str] = {}
        self.formats: dict[str, str] = {}
        self._timeline = _Timeline()
        self._t0 = 0
        self._skips: list[list] = []      # Zeitbereiche [start, ende) ohne Aufnahme (Pausen)
        self._split_at: int | None = None  # Index, an dem der aktuelle Teil endet

    # -- Status -------------------------------------------------------------
    @property
    def running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    @property
    def recorded_seconds(self) -> float:
        return self.samples_written / RECORD_RATE

    @property
    def part_seconds(self) -> float:
        return self.part_samples / RECORD_RATE

    @property
    def wall_seconds(self) -> float:
        return time.time() - self.started_at if self.started_at else 0.0

    @property
    def error(self) -> str:
        return " · ".join(f"{_SOURCE_LABELS[k]}: {v}" for k, v in self.errors.items() if v)

    # -- Steuerung ----------------------------------------------------------
    def start(self, path: Path, with_mic: bool = False, auto_split_minutes: int = 0,
              on_part: Callable[[Path, float, int], None] | None = None) -> None:
        if self.running:
            return
        path.parent.mkdir(parents=True, exist_ok=True)
        self.base_path = path
        self.with_mic = with_mic
        self.auto_split_minutes = auto_split_minutes
        self.on_part = on_part
        self.part_number = 0
        self.samples_written = 0
        self.part_samples = 0
        self.got_audio = False
        self.levels, self.errors, self.formats = {}, {}, {}
        self.paused = False
        self._skips = []
        self._split_at = None
        self.started_at = time.time()
        self._timeline = _Timeline()
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, name="recorder", daemon=True)
        self._thread.start()

    def pause(self) -> None:
        """Pausiert – der bisherige Teil wird abgeschlossen und kann ausgewertet werden."""
        if self.paused or not self.running:
            return
        now = self._now_index()
        self._skips.append([now, None])
        self._split_at = now
        self.paused = True

    def resume(self) -> None:
        if not self.paused:
            return
        if self._skips and self._skips[-1][1] is None:
            self._skips[-1][1] = self._now_index()
        self.paused = False

    def stop(self) -> Path | None:
        self.resume()
        self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=15)
        self._thread = None
        self.started_at = 0.0
        return self.path

    # -- Schreib-Thread -----------------------------------------------------
    def _open_part(self) -> wave.Wave_write:
        self.part_number += 1
        self.part_samples = 0
        self.path = self.base_path.with_name(f"{self.base_path.stem}_Teil{self.part_number:02d}.wav")
        wav = wave.open(str(self.path), "wb")
        wav.setnchannels(1)
        wav.setsampwidth(2)
        wav.setframerate(RECORD_RATE)
        return wav

    def _close_part(self, wav: wave.Wave_write) -> None:
        wav.close()
        start_seconds = (self.samples_written - self.part_samples) / RECORD_RATE
        if self.part_samples < RECORD_RATE * 2:  # leere/zu kurze Teile verwerfen
            self.samples_written -= self.part_samples
            try:
                self.path.unlink()
            except OSError:
                pass
            self.part_number -= 1
            return
        if self.on_part:
            try:
                self.on_part(self.path, start_seconds, self.part_number)
            except Exception as exc:
                self.errors["speaker"] = f"Teil konnte nicht zur Auswertung übergeben werden: {exc}"

    def _run(self) -> None:
        sources = [threading.Thread(target=self._source_loop, args=("speaker",), daemon=True)]
        if self.with_mic:
            sources.append(threading.Thread(target=self._source_loop, args=("mic",), daemon=True))
        lag = int(RECORD_RATE * 1.5)  # Vorlauf, damit beide Quellen ihre Daten geliefert haben
        wav = None
        try:
            self._t0 = _clock_100ns()
            wav = self._open_part()
            for thread in sources:
                thread.start()
            while True:
                stopping = self._stop.wait(0.5)
                if stopping:
                    for thread in sources:
                        thread.join(timeout=5)
                    target = max(self._timeline.end, self._now_index())
                else:
                    target = self._now_index() - lag

                auto = self.auto_split_minutes * 60 * RECORD_RATE
                if self._split_at is None and auto and self.part_samples >= auto:
                    self._split_at = self._timeline.base
                if self._split_at is not None and target >= self._split_at and not stopping:
                    self._flush(wav, self._split_at)
                    self._split_at = None
                    self._close_part(wav)
                    wav = self._open_part()

                self._flush(wav, target)
                if stopping:
                    break
        except Exception as exc:
            self.errors["speaker"] = f"Aufnahme konnte nicht gespeichert werden: {exc}"
            self._stop.set()
        finally:
            if wav is not None:
                self._close_part(wav)

    def _now_index(self) -> int:
        return (_clock_100ns() - self._t0) * RECORD_RATE // 10_000_000

    def _flush(self, wav: wave.Wave_write, upto: int) -> None:
        start = self._timeline.base
        audio = self._timeline.take(upto)
        if not len(audio):
            return
        keep = np.ones(len(audio), dtype=bool)
        for skip_start, skip_end in self._skips:  # pausierte Zeit verwerfen
            end = upto if skip_end is None else skip_end
            lo, hi = max(skip_start, start) - start, min(end, upto) - start
            if hi > lo:
                keep[lo:hi] = False
        self._skips = [s for s in self._skips if s[1] is None or s[1] > upto]
        audio = audio[keep]
        if not len(audio):
            return
        pcm = (np.clip(audio, -1.0, 1.0) * 32767).astype("<i2")
        wav.writeframes(pcm.tobytes())  # Dateikopf wird dabei laufend aktualisiert
        self.samples_written += len(pcm)
        self.part_samples += len(pcm)

    # -- Aufnahme-Threads (je Quelle einer) ---------------------------------
    def _source_loop(self, source: str) -> None:
        ole32 = ctypes.windll.ole32
        ole32.CoInitializeEx(None, 0)
        try:
            while not self._stop.is_set():
                try:
                    self._capture(ole32, source)
                except Exception as exc:  # z. B. Headset ab-/angesteckt, kein Mikrofon
                    self.errors[source] = f"{exc} – neuer Versuch …"
                    self.levels[source] = 0.0
                    self._stop.wait(1.0 if source == "speaker" else 3.0)
        finally:
            ole32.CoUninitialize()

    def _capture(self, ole32, source: str) -> None:
        loopback = source == "speaker"
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
            # IMMDeviceEnumerator::GetDefaultAudioEndpoint(eRender=0 / eCapture=1, eConsole=0)
            _check(_com_method(enumerator, 4, ctypes.c_int, ctypes.c_int, ctypes.POINTER(ctypes.c_void_p))(
                enumerator, 0 if loopback else 1, 0, ctypes.byref(device)),
                "Standard-Lautsprecher finden" if loopback else "Standard-Mikrofon finden")
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
            self.formats[source] = f"{rate} Hz, {channels} Kanäle"

            # IAudioClient::Initialize (geteilter Modus, 1 s Puffer)
            flags_init = _AUDCLNT_STREAMFLAGS_LOOPBACK if loopback else 0
            _check(_com_method(client, 3, ctypes.c_int, ctypes.c_uint32, ctypes.c_longlong,
                               ctypes.c_longlong, ctypes.POINTER(_WAVEFORMATEX), ctypes.c_void_p)(
                client, 0, flags_init, 10_000_000, 0, mix_format, None),
                "Aufnahme vorbereiten")
            iid_capture = _GUID.parse(_IID_IAudioCaptureClient)
            _check(_com_method(client, 14, ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p))(
                client, ctypes.byref(iid_capture), ctypes.byref(capture)), "Aufnahmedienst holen")
            _check(_com_method(client, 10)(client), "Aufnahme starten")

            get_next = _com_method(capture, 5, ctypes.POINTER(ctypes.c_uint32))
            get_buffer = _com_method(capture, 3, ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),
                                     ctypes.POINTER(ctypes.c_uint32), ctypes.POINTER(ctypes.c_uint32),
                                     ctypes.POINTER(ctypes.c_uint64), ctypes.POINTER(ctypes.c_uint64))
            release_buffer = _com_method(capture, 4, ctypes.c_uint32)

            resample = _Resampler(rate)
            fallback_index: int | None = None
            self.errors[source] = ""

            while not self._stop.is_set():
                self._stop.wait(0.02)
                packet = ctypes.c_uint32()
                _check(get_next(capture, ctypes.byref(packet)), "Audiodaten abfragen")
                while packet.value:
                    data = ctypes.POINTER(ctypes.c_ubyte)()
                    frames, flags = ctypes.c_uint32(), ctypes.c_uint32()
                    position, qpc = ctypes.c_uint64(), ctypes.c_uint64()
                    _check(get_buffer(capture, ctypes.byref(data), ctypes.byref(frames), ctypes.byref(flags),
                                      ctypes.byref(position), ctypes.byref(qpc)), "Audiodaten lesen")
                    count = frames.value
                    silent = bool(flags.value & _AUDCLNT_BUFFERFLAGS_SILENT) or not data
                    raw = None if silent else ctypes.string_at(data, count * block_align)
                    release_buffer(capture, count)

                    samples = np.zeros(count, dtype=np.float32) if raw is None \
                        else _to_mono_float(raw, tag, bits, channels)
                    out = resample(samples)
                    # Position auf der gemeinsamen Zeitachse anhand des Zeitstempels
                    if qpc.value:
                        index = (qpc.value - self._t0) * RECORD_RATE // 10_000_000
                    else:
                        index = fallback_index if fallback_index is not None else self._now_index()
                    fallback_index = index + len(out)
                    if len(out):  # Pausen verwirft der Schreib-Thread anhand der Zeitachse
                        level = float(np.sqrt(np.mean(samples ** 2))) if len(samples) else 0.0
                        self.levels[source] = 0.0 if self.paused else level
                        if level > 1e-4 and not self.paused:
                            self.got_audio = True
                        self._timeline.add(int(index), out)
                    _check(get_next(capture, ctypes.byref(packet)), "Audiodaten abfragen")

            _com_method(client, 11)(client)  # IAudioClient::Stop
        finally:
            if mix_format:
                ole32.CoTaskMemFree(mix_format)
            for obj in (capture, client, device, enumerator):
                _release(obj)


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
    return transcribe_with(load_whisper_model(model_name), path)


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


@st.cache_resource
def get_model_server() -> ModelServer:
    return ModelServer()


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
# Auswertung im Hintergrund – Teil für Teil, während die Vorlesung weiterläuft
# ---------------------------------------------------------------------------
class LiveProcessor:
    def __init__(self) -> None:
        self._queue: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._models: dict = {}
        self.recorder: LoopbackRecorder | None = None
        self.reset("", "base")

    def reset(self, title: str, model_name: str) -> None:
        self.session_id = uuid.uuid4().hex  # alte, noch wartende Aufträge werden ignoriert
        self.title = title
        self.model_name = model_name
        self.parts: list[dict] = []
        self.closed = False
        self.final_status = ""
        self.error = ""
        self.result: dict | None = None

    @property
    def active(self) -> bool:
        return bool(self.parts) or self.closed

    @property
    def busy(self) -> bool:
        if any(p["state"] in ("wartet", "läuft") for p in self.parts):
            return True
        return self.closed and self.result is None and not self.error

    def submit(self, path: Path, offset: float, number: int) -> None:
        part = {"number": number, "path": path, "offset": offset, "state": "wartet",
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
            self._thread = threading.Thread(target=self._work, name="live-processor", daemon=True)
            self._thread.start()

    def _model(self):
        if self.model_name not in self._models:
            self._models[self.model_name] = whisper.load_model(
                self.model_name, download_root=str(WHISPER_MODEL_DIR))
        return self._models[self.model_name]

    def _limit_threads(self) -> None:
        """Während der Vorlesung nur die Hälfte der Prozessorkerne nutzen, damit LearnSpace3D flüssig bleibt."""
        try:
            import torch
            cores = os.cpu_count() or 2
            recording = self.recorder is not None and self.recorder.running
            torch.set_num_threads(max(1, cores // 2) if recording else cores)
        except Exception:
            pass

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
        part["state"], part["status"] = "läuft", "🎧 Sprache wird in Text umgewandelt …"
        self._limit_threads()
        segments = transcribe_with(self._model(), str(part["path"]))
        for seg in segments:
            seg["start"] += part["offset"]
            seg["end"] += part["offset"]
        part["segments"] = segments
        chunks = chunk_segments(segments)
        write_notes(chunks, self.title, lambda v, t: part.__setitem__("status", f"🧠 {t}"))
        part["chunks"] = chunks
        part["state"] = "fertig"
        part["status"] = "✅ Fertig" if segments else "✅ Fertig (keine Sprache erkannt)"

    def _do_final(self) -> None:
        segments = [s for p in self.parts for s in p["segments"]]
        chunks = [c for p in self.parts for c in p["chunks"]]
        if not chunks:
            self.error = "In der Aufnahme wurde keine Sprache erkannt."
            return
        summary = summarize_parts(chunks, self.title, lambda v, t: setattr(self, "final_status", t))
        created = dt.datetime.now()
        docx_bytes = build_docx(self.title, summary, chunks, segments, created)
        filename = f"{created:%Y-%m-%d_%H%M}_{safe_filename(self.title)}.docx"
        saved_to = ""
        try:
            NOTES_DIR.mkdir(parents=True, exist_ok=True)
            (NOTES_DIR / filename).write_bytes(docx_bytes)
            saved_to = str(NOTES_DIR / filename)
        except OSError:
            pass
        self.final_status = "Fertig."
        self.result = {
            "id": self.session_id, "title": self.title, "summary": summary, "parts": chunks,
            "transcript": "\n".join(f"[{fmt_time(s['start'])}] {s['text']}" for s in segments),
            "docx": docx_bytes, "filename": filename, "saved_to": saved_to,
        }


@st.cache_resource
def get_live_processor() -> LiveProcessor:
    return LiveProcessor()


# ---------------------------------------------------------------------------
# Pausen-Spiel „Kachelfusion“ (läuft komplett im Browser)
# ---------------------------------------------------------------------------
GAME_HTML = r"""
<!doctype html><html><head><meta charset="utf-8">
<style>
  :root { --bg:#0f1117; --panel:#1a1d29; --cell:#232736; --text:#e8eaf2; --muted:#8b90a5; --accent:#ff4b6e; }
  * { box-sizing:border-box; }
  body { margin:0; font-family:"Source Sans Pro",system-ui,sans-serif; background:transparent; color:var(--text);
         display:flex; justify-content:center; user-select:none; }
  .wrap { width:min(440px,100%); padding:8px; outline:none; }
  .top { display:flex; align-items:center; gap:8px; margin-bottom:12px; }
  .title { font-size:26px; font-weight:800; letter-spacing:.5px; flex:1;
           background:linear-gradient(90deg,#4fd1c5,#9f7aea,#ff4b6e); -webkit-background-clip:text; color:transparent; }
  .score { background:var(--panel); border-radius:10px; padding:6px 12px; text-align:center; min-width:72px; }
  .score small { display:block; font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:1px; }
  .score b { font-size:18px; }
  button { background:var(--accent); color:#fff; border:0; border-radius:10px; padding:10px 14px; font-weight:700;
           cursor:pointer; font-size:14px; }
  button:hover { filter:brightness(1.1); }
  .board { position:relative; background:var(--panel); border-radius:16px; padding:10px; aspect-ratio:1;
           display:grid; grid-template-columns:repeat(4,1fr); gap:10px; touch-action:none; }
  .cell { background:var(--cell); border-radius:12px; }
  .tile { position:absolute; display:flex; align-items:center; justify-content:center; border-radius:12px;
          font-weight:800; transition:transform 110ms ease-out; will-change:transform; }
  .tile.new { animation:pop 160ms ease-out; }
  .tile.merged .in { animation:bump 180ms ease-out; }
  .in { width:100%; height:100%; display:flex; align-items:center; justify-content:center; border-radius:12px; }
  @keyframes pop { from { opacity:0; scale:.4 } to { opacity:1; scale:1 } }
  @keyframes bump { 50% { scale:1.15 } }
  .hint { color:var(--muted); font-size:13px; text-align:center; margin-top:10px; }
  .overlay { position:absolute; inset:0; border-radius:16px; background:rgba(15,17,23,.82); display:none;
             flex-direction:column; align-items:center; justify-content:center; gap:12px; z-index:5; }
  .overlay.show { display:flex; }
  .overlay h2 { margin:0; font-size:28px; }
</style></head>
<body>
<div class="wrap" id="wrap" tabindex="0">
  <div class="top">
    <div class="title">Kachelfusion</div>
    <div class="score"><small>Punkte</small><b id="score">0</b></div>
    <div class="score"><small>Rekord</small><b id="best">0</b></div>
    <button id="new">Neu</button>
  </div>
  <div class="board" id="board">
    <div class="overlay" id="overlay"><h2 id="msg"></h2><button id="again">Nochmal</button></div>
  </div>
  <div class="hint">Pfeiltasten oder W A S D · auf dem Handy wischen · gleiche Kacheln verschmelzen</div>
</div>
<script>
(() => {
  const N = 4;
  const COLORS = {2:["#2b3a55","#cfe3ff"],4:["#2c4f6b","#d9f0ff"],8:["#167a8a","#fff"],16:["#13a39a","#fff"],
    32:["#3cb371","#fff"],64:["#7bc043","#102010"],128:["#c9b52a","#1b1700"],256:["#e8912d","#fff"],
    512:["#e5603d","#fff"],1024:["#d6336c","#fff"],2048:["#9f43d9","#fff"],4096:["#6741d9","#fff"]};
  const store = { get(k){ try { return localStorage.getItem(k); } catch(e) { return null; } },
                  set(k,v){ try { localStorage.setItem(k,v); } catch(e) {} } };
  const board = document.getElementById("board"), wrap = document.getElementById("wrap");
  const scoreEl = document.getElementById("score"), bestEl = document.getElementById("best");
  const overlay = document.getElementById("overlay"), msg = document.getElementById("msg");
  for (let i = 0; i < N*N; i++) { const c = document.createElement("div"); c.className = "cell"; board.appendChild(c); }
  let grid, score, best = +(store.get("kachelfusion-best") || 0), won, nextId = 1, tiles = new Map();
  bestEl.textContent = best;

  function empty() { return Array.from({length:N}, () => Array(N).fill(null)); }
  function addRandom() {
    const free = [];
    for (let r = 0; r < N; r++) for (let c = 0; c < N; c++) if (!grid[r][c]) free.push([r,c]);
    if (!free.length) return;
    const [r,c] = free[Math.floor(Math.random()*free.length)];
    grid[r][c] = { id: nextId++, v: Math.random() < 0.9 ? 2 : 4, isNew: true };
  }
  function start(saved) {
    board.querySelectorAll(".tile").forEach(t => t.remove()); tiles.clear();
    overlay.classList.remove("show"); won = false;
    if (saved) { grid = saved.grid.map(row => row.map(v => v ? { id: nextId++, v } : null)); score = saved.score; won = saved.won; }
    else { grid = empty(); score = 0; addRandom(); addRandom(); }
    render(); wrap.focus();
  }
  function save() {
    store.set("kachelfusion-state", JSON.stringify({ grid: grid.map(r => r.map(t => t ? t.v : 0)), score, won }));
  }
  function geometry() {
    const pad = 10, gap = 10, size = (board.clientWidth - 2*pad - (N-1)*gap) / N;
    return { pad, gap, size };
  }
  function render() {
    const { pad, gap, size } = geometry(), seen = new Set();
    for (let r = 0; r < N; r++) for (let c = 0; c < N; c++) {
      const t = grid[r][c]; if (!t) continue; seen.add(t.id);
      let el = tiles.get(t.id);
      if (!el) { el = document.createElement("div"); el.className = "tile"; el.innerHTML = '<div class="in"></div>';
                 board.appendChild(el); tiles.set(t.id, el); if (t.isNew) el.classList.add("new"); }
      const [bg, fg] = COLORS[t.v] || ["#212529","#fff"], inner = el.firstChild;
      inner.style.background = bg; inner.style.color = fg; inner.textContent = t.v;
      inner.style.fontSize = (t.v >= 1024 ? size*0.3 : t.v >= 128 ? size*0.36 : size*0.42) + "px";
      el.style.width = el.style.height = size + "px";
      el.style.transform = `translate(${pad + c*(size+gap)}px, ${pad + r*(size+gap)}px)`;
      el.classList.toggle("merged", !!t.merged);
      t.isNew = false; t.merged = false;
    }
    for (const [id, el] of tiles) if (!seen.has(id)) { el.remove(); tiles.delete(id); }
    scoreEl.textContent = score;
    if (score > best) { best = score; bestEl.textContent = best; store.set("kachelfusion-best", best); }
  }
  function move(dir) {   // 0 links, 1 rechts, 2 hoch, 3 runter
    if (overlay.classList.contains("show")) return;
    let moved = false;
    const line = i => Array.from({length:N}, (_, j) => dir < 2 ? [i, dir === 0 ? j : N-1-j] : [dir === 2 ? j : N-1-j, i]);
    const absorbed = [];
    for (let i = 0; i < N; i++) {
      const cells = line(i), items = cells.map(([r,c]) => grid[r][c]).filter(Boolean), out = [];
      for (let k = 0; k < items.length; k++) {
        if (k + 1 < items.length && items[k].v === items[k+1].v) {
          const keep = items[k]; keep.v *= 2; keep.merged = true; score += keep.v;
          absorbed.push([items[k+1], out.length]); out.push(keep); k++;
          if (keep.v === 2048 && !won) { won = true; setTimeout(() => show("Geschafft! 🎉", "Weiterspielen"), 250); }
        } else out.push(items[k]);
      }
      cells.forEach(([r,c], j) => { const t = out[j] || null; if (grid[r][c] !== t) moved = true; grid[r][c] = t; });
      // verschmolzene Kachel noch kurz an die Zielposition gleiten lassen
      absorbed.splice(0).forEach(([t, j]) => { const el = tiles.get(t.id); if (!el) return;
        const [r,c] = cells[j], { pad, gap, size } = geometry();
        el.style.transform = `translate(${pad + c*(size+gap)}px, ${pad + r*(size+gap)}px)`; el.style.zIndex = 0;
        tiles.delete(t.id); setTimeout(() => el.remove(), 120); });
    }
    if (!moved) return;
    addRandom(); render(); save();
    if (!canMove()) setTimeout(() => show("Keine Züge mehr", "Nochmal"), 250);
  }
  function canMove() {
    for (let r = 0; r < N; r++) for (let c = 0; c < N; c++) {
      const v = grid[r][c] && grid[r][c].v; if (!v) return true;
      if (c+1 < N && grid[r][c+1] && grid[r][c+1].v === v) return true;
      if (r+1 < N && grid[r+1][c] && grid[r+1][c].v === v) return true;
    }
    return false;
  }
  function show(text, label) { msg.textContent = text; document.getElementById("again").textContent = label;
                               overlay.classList.add("show"); }
  document.getElementById("again").onclick = () => {
    if (won && canMove() && msg.textContent.startsWith("Geschafft")) { overlay.classList.remove("show"); wrap.focus(); }
    else { store.set("kachelfusion-state", ""); start(); }
  };
  document.getElementById("new").onclick = () => { store.set("kachelfusion-state", ""); start(); };
  const KEYS = { ArrowLeft:0, ArrowRight:1, ArrowUp:2, ArrowDown:3, a:0, d:1, w:2, s:3, A:0, D:1, W:2, S:3 };
  wrap.addEventListener("keydown", e => { if (e.key in KEYS) { e.preventDefault(); move(KEYS[e.key]); } });
  board.addEventListener("pointerdown", () => wrap.focus());
  let sx = 0, sy = 0;
  board.addEventListener("touchstart", e => { sx = e.touches[0].clientX; sy = e.touches[0].clientY; }, {passive:true});
  board.addEventListener("touchend", e => {
    const dx = e.changedTouches[0].clientX - sx, dy = e.changedTouches[0].clientY - sy;
    if (Math.max(Math.abs(dx), Math.abs(dy)) < 25) return;
    move(Math.abs(dx) > Math.abs(dy) ? (dx < 0 ? 0 : 1) : (dy < 0 ? 2 : 3));
  });
  window.addEventListener("resize", render);
  window.__game = { move, state: () => ({ grid: grid.map(r => r.map(t => t ? t.v : 0)), score }) };
  let saved = null; try { saved = JSON.parse(store.get("kachelfusion-state") || "null"); } catch(e) {}
  start(saved && saved.grid ? saved : null);
})();
</script></body></html>
"""


def render_game() -> None:
    st.caption("Kurze Pause gefällig? Die Aufnahme läuft im Hintergrund weiter, auch wenn du hier spielst. "
               "Einmal ins Spielfeld klicken, dann mit den Pfeiltasten spielen.")
    components.html(GAME_HTML, height=560)


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
    _busy = get_recorder().running or get_live_processor().busy
    _models = sorted(MODELS_DIR.glob("*.gguf"), key=lambda p: (p.name != BUNDLED_MODEL, p.name.lower())) \
        if MODELS_DIR.exists() else []
    _bundled = MODELS_DIR / BUNDLED_MODEL
    if not _models:
        _models = [_bundled]
    # Standard: Qwen 3 8B, falls vorhanden – sonst das größte zusätzliche Modell, sonst das mitgelieferte
    _preferred = next((m for m in _models if PREFERRED_MODEL in m.name.lower()), None) \
        or max((m for m in _models if m.name != BUNDLED_MODEL), key=lambda m: m.stat().st_size, default=None)
    chosen_model = st.selectbox(
        "KI-Modell für die Zusammenfassung", _models, format_func=model_label, key="ai_model",
        index=_models.index(_preferred) if _preferred in _models else 0,
        disabled=_busy, help="Weitere Modelle als .gguf-Datei in den Modell-Ordner legen (siehe unten).",
    )
    context = 2048
    if chosen_model.name != BUNDLED_MODEL:
        context = st.select_slider(
            "Kontextgröße (wie viel Text die KI auf einmal sieht)", [4096, 8192, 16384], value=8192,
            key="ai_context", disabled=_busy,
            help="Größer = besserer roter Faden, braucht aber mehr Arbeitsspeicher.",
        )
        ram = total_ram_gb()
        size = chosen_model.stat().st_size / 1024 ** 3
        if ram and size > ram * 0.6:
            st.warning(f"Das Modell ({size:.1f} GB) ist sehr groß für {ram:.0f} GB Arbeitsspeicher.")
    _server = get_model_server()
    if chosen_model.name == BUNDLED_MODEL:
        _server.stop()
        AI.configure(BUNDLED_AI_URL, 2048, model_label(chosen_model))
    elif not _busy or not _server.running_with(chosen_model, context):
        try:
            if not _server.running_with(chosen_model, context):
                with st.spinner(f"{chosen_model.stem} wird geladen … (beim ersten Mal bis zu einigen Minuten)"):
                    _server.ensure(chosen_model, context)
            AI.configure(_server.url, context, chosen_model.stem)
        except Exception as exc:
            st.error(f"Modell konnte nicht gestartet werden – es wird das mitgelieferte genutzt.\n\n{exc}")
            AI.configure(BUNDLED_AI_URL, 2048, model_label(_bundled))
    st.caption(f"Aktiv: **{AI.label}** · Kontext {AI.context} · Abschnitte à ca. "
               f"{AI.chunk_chars:,} Zeichen".replace(f"{AI.chunk_chars:,}", f"{AI.chunk_chars:,}".replace(",", ".")))
    whisper_model = st.selectbox(
        "Spracherkennung (Whisper)", ["base", "small", "turbo"], index=0,
        format_func={"base": "base – mitgeliefert, schnell, macht mehr Hörfehler",
                     "small": "small – deutlich genauer (einmalig ca. 460 MB Download)",
                     "turbo": "turbo – sehr genau, langsam ohne Grafikkarte (einmalig ca. 1,6 GB)"}.get,
        help="Je genauer die Spracherkennung, desto besser werden auch die Notizen. "
             "„small“ und „turbo“ werden beim ersten Mal einmalig aus dem Internet geladen.",
        disabled=get_recorder().running or get_live_processor().busy,
    )
    st.caption(f"Modell-Ordner: `{MODELS_DIR}`")
    st.caption(f"Aufnahmen: `{RECORDINGS_DIR}`")
    st.caption(f"Notizen: `{NOTES_DIR}`")

lecture_tab, file_tab, game_tab = st.tabs(["🎓 Vorlesung aufnehmen", "📁 Datei auswerten", "🎮 Pause"])

recorder = get_recorder()
live = get_live_processor()
live.recorder = recorder

with lecture_tab:
    if not sys.platform.startswith("win"):
        st.error("Die Systemton-Aufnahme funktioniert nur unter Windows.")
    else:
        title = st.text_input("Thema / Fach der Vorlesung", key="lecture_title",
                              placeholder="z. B. Netzwerktechnik – Subnetting",
                              disabled=recorder.running)
        st.markdown(
            "**So geht's:** 1. LearnSpace3D öffnen und der Vorlesung beitreten · "
            "2. hier auf **Aufnahme starten** klicken · 3. am Ende **Stoppen & Notizen erstellen**.\n\n"
            "Jedes Mal, wenn du **Pause** drückst, wird der bisherige Teil schon **im Hintergrund "
            "ausgewertet**, während du danach mit **Weiter** den nächsten Teil aufnimmst. "
            "Aufgenommen wird alles, was über deine Lautsprecher bzw. dein Headset läuft – auf Wunsch "
            "zusätzlich dein Mikrofon. Andere Töne (Musik, Videos) bitte ausschalten."
        )
        consent = st.checkbox(
            "Ich habe die Erlaubnis der Lehrkraft bzw. der WBS, die Vorlesung für meine "
            "persönlichen Lernnotizen aufzunehmen, und gebe die Aufnahme nicht weiter.",
            key="consent", disabled=recorder.running,
        )

        @st.fragment(run_every=1.0 if (recorder.running or live.busy) else None)
        def recorder_panel() -> None:
            if recorder.running:
                state = "⏸️ Pausiert" if recorder.paused else "🔴 Aufnahme läuft"
                st.subheader(f"{state} – {fmt_time(recorder.wall_seconds)}")
                st.progress(min(recorder.levels.get("speaker", 0.0) * 4, 1.0),
                            text="🔊 Lautsprecher/Headset (bewegt sich, wenn Ton ankommt)")
                if recorder.with_mic:
                    st.progress(min(recorder.levels.get("mic", 0.0) * 4, 1.0), text="🎙️ Dein Mikrofon")
                auto = f" · neuer Teil automatisch alle {recorder.auto_split_minutes} Min" \
                    if recorder.auto_split_minutes else ""
                st.caption(f"Teil {recorder.part_number}: {fmt_time(recorder.part_seconds)} · "
                           f"gesamt {fmt_time(recorder.recorded_seconds)}{auto}")
                if recorder.error:
                    st.warning(recorder.error)
                col_pause, col_stop = st.columns(2)
                if recorder.paused:
                    if col_pause.button("▶️ Weiter (nächster Teil)", width="stretch"):
                        recorder.resume()
                        st.rerun()
                elif col_pause.button("⏸️ Pause & bisherigen Teil auswerten", width="stretch"):
                    recorder.pause()
                    st.rerun()
                if col_stop.button("⏹️ Stoppen & Notizen erstellen", type="primary", width="stretch"):
                    recorder.stop()
                    if not recorder.got_audio:
                        st.session_state["no_audio"] = recorder.error.replace(" – neuer Versuch …", "")
                    live.finish()
                    st.rerun()
            elif not live.busy:
                with_mic = st.checkbox(
                    "🎙️ Mein Mikrofon auch aufnehmen (damit deine Fragen und Beiträge in den Notizen stehen)",
                    key="with_mic",
                    help="Am besten mit Headset. Bei Laptop-Lautsprechern nimmt das Mikrofon die Lehrkraft "
                         "ein zweites Mal auf, das kann die Erkennung verschlechtern.",
                )
                auto_minutes = st.number_input(
                    "Automatisch alle … Minuten einen neuen Teil beginnen und auswerten (0 = nur bei Pause)",
                    min_value=0, max_value=120, value=30, step=5, key="auto_split",
                )
                if st.button("🔴 Aufnahme starten", type="primary", width="stretch", disabled=not consent):
                    title_now = st.session_state.get("lecture_title", "")
                    name = f"{dt.datetime.now():%Y-%m-%d_%H%M}_{safe_filename(title_now)}.wav"
                    live.reset(title_now, whisper_model)
                    st.session_state.pop("no_audio", None)
                    recorder.start(RECORDINGS_DIR / name, with_mic=with_mic,
                                   auto_split_minutes=int(auto_minutes), on_part=live.submit)
                    st.rerun()
                if not consent:
                    st.caption("Bitte zuerst die Erlaubnis bestätigen.")

            if recorder.running and not live.parts:
                st.caption("Der erste Teil wird ausgewertet, sobald du Pause drückst"
                           + (f" oder nach {recorder.auto_split_minutes} Minuten." if recorder.auto_split_minutes else "."))

            # Fortschritt der einzelnen Teile
            if live.active:
                st.markdown("**Auswertung der Teile**")
                for part in live.parts:
                    start = fmt_time(part["offset"])
                    st.write(f"Teil {part['number']} (ab {start}): {part['status']}")
                    if part["state"] == "fertig" and part["chunks"]:
                        with st.expander(f"Notizen zu Teil {part['number']} ansehen"):
                            for chunk in part["chunks"]:
                                st.markdown(f"**{chunk['label']}**\n\n{chunk.get('notes', '')}")
                if live.closed and live.busy:
                    st.info("⏳ " + (live.final_status or "Die restlichen Teile werden noch ausgewertet …"))
                if live.error and st.session_state.get("no_audio") is None:
                    st.error(live.error)
                if live.result and st.session_state.get("shown_result") != live.result["id"]:
                    st.session_state["shown_result"] = live.result["id"]
                    st.session_state["result"] = live.result
                    st.rerun()

        recorder_panel()

        if st.session_state.get("no_audio") is not None and live.closed and not live.busy:
            reason = st.session_state["no_audio"]
            st.error("Es kam kein Ton an. Lief die Vorlesung über das Standard-Ausgabegerät von Windows?"
                     + (f" ({reason})" if reason else ""))

        with st.expander("Ältere Aufnahme auswerten (z. B. nach einem Absturz)"):
            recordings = sorted(RECORDINGS_DIR.glob("*.wav"), reverse=True) if RECORDINGS_DIR.exists() else []
            if not recordings:
                st.caption("Noch keine gespeicherten Aufnahmen.")
            else:
                chosen = st.selectbox("Aufnahme", recordings, format_func=lambda p: p.name)
                if st.button("Diese Aufnahme auswerten", disabled=recorder.running or live.busy):
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

with game_tab:
    render_game()

with st.expander("Datenschutz und Hinweise"):
    st.markdown(
        "- Aufnahme, Spracherkennung und Zusammenfassung laufen **komplett lokal** auf diesem PC. Es wird nichts ins Internet geschickt.\n"
        "- Nimm Vorlesungen nur mit Erlaubnis auf und gib die Aufnahmen nicht weiter. Das heimliche Aufnehmen nicht öffentlicher Gespräche ist in Deutschland strafbar (§ 201 StGB).\n"
        "- Die Aufnahmen bleiben im Ordner „Aufnahmen“ gespeichert, bis du sie löschst (ca. 115 MB pro Stunde).\n"
        "- Das mitgelieferte KI-Modell ist sehr klein. Bessere Modelle (.gguf) kannst du in den Modell-Ordner legen und in der Seitenleiste auswählen.\n"
        "- Die Notizen sind eine Hilfe, ersetzen aber nicht das Nacharbeiten."
    )
