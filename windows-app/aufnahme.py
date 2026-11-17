# -*- coding: utf-8 -*-
"""Aufnahme von Systemton (WASAPI-Loopback) und optional Mikrofon – ohne Zusatz-Bibliotheken.

Mit Mikrofon wird in Stereo gespeichert: links Lautsprecher (Lehrkraft/Kurs), rechts Mikrofon (du).
So lässt sich später erkennen, wer gerade spricht. Die Spracherkennung hört beide Kanäle zusammen.

Übernommen aus app.py (Website-Version), damit das Windows-Programm dieselbe,
bereits getestete Aufnahme-Technik nutzt. Später sollen beide Versionen dieses Modul teilen.
"""
from __future__ import annotations

import ctypes
import threading
import time
import uuid
import wave
from pathlib import Path
from typing import Callable

import numpy as np

RECORD_RATE = 16_000  # Whisper arbeitet intern mit 16 kHz mono


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
        self._timeline = _Timeline()       # Lautsprecher
        self._mic_timeline = _Timeline()   # Mikrofon (eigener Kanal)
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
        self._timeline = _Timeline()       # Lautsprecher
        self._mic_timeline = _Timeline()   # Mikrofon (eigener Kanal)
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
        wav.setnchannels(2 if self.with_mic else 1)
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
                    target = max(self._timeline.end, self._mic_timeline.end, self._now_index())
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
        mic = self._mic_timeline.take(upto)  # gleich lang: beide Zeitachsen laufen im Gleichschritt
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
        if self.with_mic:  # Stereo: links Lautsprecher, rechts Mikrofon
            frames = np.column_stack((audio, mic[keep])).ravel()
        else:
            frames = audio
        pcm = (np.clip(frames, -1.0, 1.0) * 32767).astype("<i2")
        wav.writeframes(pcm.tobytes())  # Dateikopf wird dabei laufend aktualisiert
        self.samples_written += len(audio)
        self.part_samples += len(audio)

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
                        timeline = self._mic_timeline if source == "mic" else self._timeline
                        timeline.add(int(index), out)
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
