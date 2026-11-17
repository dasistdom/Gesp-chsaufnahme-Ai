# -*- coding: utf-8 -*-
"""Gesprächsnotiz-AI – Windows-Programm.

Eigene Oberfläche mit Qt statt Browser:
  • Aufnehmen: Systemton (LearnSpace3D) + optional Mikrofon, Pause, automatische Teile.
    Jeder fertige Teil wird sofort im Hintergrund ausgewertet (Spracherkennung + KI);
    nach dem Stoppen entsteht automatisch das Word-Dokument.
  • Notizen: alle erstellten Word-Dokumente durchsuchen, ansehen und in Word öffnen
  • Einstellungen: KI-Modell, Spracherkennung, Programmordner, Ordner

Die KI, die Modelle und ffmpeg kommen aus der vorhandenen GespraechsnotizAI-Installation.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import sys
import threading
import wave
from pathlib import Path

from PySide6.QtCore import QSize, Qt, QTimer, QUrl
from PySide6.QtGui import QColor, QDesktopServices, QFont, QGuiApplication, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFrame, QHBoxLayout, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMainWindow, QMessageBox, QProgressBar, QPushButton,
    QScrollArea, QSizePolicy, QSplitter, QStackedWidget, QTextBrowser, QVBoxLayout, QWidget,
)

# Ohne Konsolenfenster (pythonw.exe) gibt es keine Textausgabe: sys.stdout/sys.stderr sind None und jedes
# print oder jeder Fortschrittsbalken würde abstürzen. Deshalb alles in eine Log-Datei umleiten – noch bevor
# Bausteine wie Whisper geladen werden.
if sys.stdout is None or sys.stderr is None:
    _log_dir = Path.home() / "Documents" / "GespraechsnotizAI"
    try:
        _log_dir.mkdir(parents=True, exist_ok=True)
        _log = open(_log_dir / "windows-app.log", "a", encoding="utf-8", buffering=1)
    except OSError:
        _log = open(os.devnull, "w", encoding="utf-8")
    sys.stdout = sys.stdout or _log
    sys.stderr = sys.stderr or _log

import auswertung as aw  # noqa: E402
from aufnahme import LoopbackRecorder  # noqa: E402
from notizen import Notiz, list_notes, read_docx  # noqa: E402
import claude_chat  # noqa: E402
from teilen import DiscordShare, is_webhook  # noqa: E402

APP_NAME = "Gesprächsnotiz-AI"
DATA_DIR = aw.DATA_DIR  # dieselben Ordner wie die Website-Version
RECORDINGS_DIR = DATA_DIR / "Aufnahmen"
NOTES_DIR = aw.NOTES_DIR
SETTINGS_FILE = DATA_DIR / "windows-app-einstellungen.json"
IS_WINDOWS = sys.platform.startswith("win")
WHISPER_CHOICES = {
    "base": "base – mitgeliefert, schnell, mehr Hörfehler",
    "small": "small – deutlich genauer (einmalig ca. 460 MB Download)",
    "turbo": "turbo – sehr genau, langsam ohne Grafikkarte (einmalig ca. 1,6 GB)",
}

# --- Gestaltung ---------------------------------------------------------------
# Ruhiges Grau mit leichtem Blaustich, Indigo als Akzent, Rot nur für die Aufnahme.
THEMES = {
    "light": {"bg": "#f4f5f8", "panel": "#ffffff", "nav": "#eceef4", "text": "#1d2130",
              "muted": "#646b80", "line": "#dde0ea", "accent": "#4f5bd5", "accent_text": "#ffffff",
              "rec": "#e5484d", "ok": "#2f9e6e", "warn": "#c2410c", "hover": "#e2e5f0"},
    "dark": {"bg": "#16181f", "panel": "#1e212b", "nav": "#121419", "text": "#e7e9f0",
             "muted": "#9aa0b4", "line": "#2c303d", "accent": "#7a84f0", "accent_text": "#0f1016",
             "rec": "#ff6369", "ok": "#4cc38a", "warn": "#fb923c", "hover": "#262a36"},
}


def build_stylesheet(c: dict) -> str:
    return f"""
    QWidget {{ background: {c['bg']}; color: {c['text']}; font-family: 'Segoe UI', 'Inter', sans-serif;
               font-size: 10.5pt; }}
    #Nav {{ background: {c['nav']}; border: none; outline: none; padding: 12px 8px; }}
    #Nav::item {{ padding: 10px 14px; border-radius: 8px; margin: 2px 0; color: {c['muted']}; }}
    #Nav::item:hover {{ background: {c['hover']}; }}
    #Nav::item:selected {{ background: {c['panel']}; color: {c['text']}; font-weight: 600; }}
    #Brand {{ background: {c['nav']}; font-size: 13pt; font-weight: 700; padding: 18px 20px 6px 20px; }}
    #BrandSub {{ background: {c['nav']}; color: {c['muted']}; font-size: 9pt; padding: 0 20px 10px 20px; }}
    #AiStatus {{ background: {c['nav']}; color: {c['muted']}; font-size: 9pt; padding: 12px 20px 16px 20px; }}
    #Card {{ background: {c['panel']}; border: 1px solid {c['line']}; border-radius: 12px; }}
    #Card QWidget {{ background: {c['panel']}; }}
    #PageTitle {{ font-size: 18pt; font-weight: 700; }}
    #Muted {{ color: {c['muted']}; }}
    #Timer {{ font-size: 34pt; font-weight: 600; font-family: 'Segoe UI', 'Inter', sans-serif; }}
    #Banner {{ background: {c['hover']}; border-radius: 8px; padding: 8px 12px; color: {c['muted']}; }}
    #Result {{ color: {c['ok']}; font-weight: 600; }}
    #Error {{ color: {c['warn']}; }}
    QLineEdit {{ background: {c['bg']}; border: 1px solid {c['line']}; border-radius: 8px; padding: 7px 10px; }}
    QLineEdit:focus {{ border: 1px solid {c['accent']}; }}
    QComboBox {{ background: {c['bg']}; border: 1px solid {c['line']}; border-radius: 8px; padding: 6px 10px; }}
    QComboBox:focus {{ border: 1px solid {c['accent']}; }}
    QComboBox QAbstractItemView {{ background: {c['panel']}; border: 1px solid {c['line']};
                                   selection-background-color: {c['hover']}; selection-color: {c['text']}; }}
    QPushButton {{ background: {c['panel']}; border: 1px solid {c['line']}; border-radius: 8px;
                   padding: 8px 16px; }}
    QPushButton:hover {{ background: {c['hover']}; }}
    QPushButton:focus {{ border: 1px solid {c['accent']}; }}
    QPushButton:disabled {{ color: {c['muted']}; }}
    QPushButton#Primary {{ background: {c['accent']}; color: {c['accent_text']}; border: none; font-weight: 600; }}
    QPushButton#Record {{ background: {c['rec']}; color: #ffffff; border: none; border-radius: 48px;
                          font-size: 11pt; font-weight: 700; }}
    QPushButton#Record:disabled {{ background: {c['line']}; color: {c['muted']}; }}
    QPushButton#Stop {{ background: {c['text']}; color: {c['panel']}; border: none; font-weight: 600; }}
    QProgressBar {{ background: {c['bg']}; border: none; border-radius: 4px; height: 8px; }}
    QProgressBar::chunk {{ background: {c['ok']}; border-radius: 4px; }}
    QListWidget {{ background: {c['panel']}; border: 1px solid {c['line']}; border-radius: 10px; padding: 4px; }}
    QListWidget::item {{ padding: 8px 10px; border-radius: 6px; }}
    QListWidget::item:selected {{ background: {c['hover']}; color: {c['text']}; }}
    QTextBrowser {{ background: {c['panel']}; border: 1px solid {c['line']}; border-radius: 10px; padding: 14px; }}
    QSplitter::handle {{ background: transparent; width: 10px; }}
    QCheckBox {{ spacing: 8px; }}
    """


def card() -> tuple[QFrame, QVBoxLayout]:
    frame = QFrame()
    frame.setObjectName("Card")
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(20, 18, 20, 18)
    layout.setSpacing(12)
    return frame, layout


def label(text: str, name: str = "", wrap: bool = False) -> QLabel:
    widget = QLabel(text)
    if name:
        widget.setObjectName(name)
    widget.setWordWrap(wrap)
    return widget


def fmt_time(seconds: float) -> str:
    seconds = int(seconds)
    h, rest = divmod(seconds, 3600)
    m, s = divmod(rest, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


def safe_filename(text: str) -> str:
    text = re.sub(r"[^\w\- ]+", "", text, flags=re.UNICODE).strip().replace(" ", "_")
    return text[:60] or "Vorlesung"


def open_path(path: Path) -> None:
    if IS_WINDOWS:
        os.startfile(path)  # öffnet mit dem Standardprogramm (z. B. Word)
    else:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))


def path_label(path: str) -> QLabel:
    """Zeigt einen Pfad an, der bei Platzmangel gekürzt wird statt das Fenster zu verbreitern."""
    widget = label(path, "Muted")
    widget.setToolTip(path)
    widget.setMinimumWidth(1)
    widget.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
    return widget


def copy_for_claude(docx_path: Path) -> str:
    """Kopiert Anweisung + Transkript für claude.ai in die Zwischenablage und öffnet claude.ai."""
    try:
        paragraphs = read_docx(docx_path)
    except Exception as exc:
        return f"Die Notiz konnte nicht gelesen werden: {exc}"
    title = next((t for s, t in paragraphs if s == "Title"), docx_path.stem)
    note = Notiz(docx_path, title, dt.datetime.fromtimestamp(docx_path.stat().st_mtime), paragraphs)
    if not note.transcript.strip():
        return ("Diese Notiz enthält kein Transkript (z. B. eine schon von Claude erstellte Notiz). "
                "Bitte die ursprüngliche Vorlesungsnotiz wählen.")
    prompt = claude_chat.build_prompt(note.title, note.modified, note.transcript, note.has_speakers)
    QGuiApplication.clipboard().setText(prompt)
    QDesktopServices.openUrl(QUrl(claude_chat.CLAUDE_URL))
    message = (f"Kopiert ({len(prompt):,} Zeichen). claude.ai wurde geöffnet: dort mit Strg+V einfügen und "
               "absenden. Danach Claudes Antwort mit dem Kopieren-Knopf unter der Antwort kopieren und unter "
               "„Notizen“ auf „Claude-Antwort speichern“ klicken.").replace(",", ".", 1)
    if len(prompt) > claude_chat.LARGE_PROMPT:
        message += (" Hinweis: Der Text ist sehr lang – mit einem Gratis-Konto kann das Tageslimit erreicht "
                    "werden. Dann eine kürzere Aufnahme bzw. einen einzelnen Teil verwenden.")
    return message


def wav_seconds(path: Path) -> float:
    try:
        with wave.open(str(path), "rb") as wav:
            return wav.getnframes() / wav.getframerate()
    except (OSError, wave.Error, EOFError):
        return 0.0


# --- Einstellungen und KI-Zustand -------------------------------------------------
class Settings:
    def __init__(self) -> None:
        self.values = {"program_dir": "", "whisper": "base", "model": "", "context": 8192,
                       "with_mic": False, "split_minutes": 30, "discord_url": "", "discord_on": False}
        try:
            self.values.update(json.loads(SETTINGS_FILE.read_text(encoding="utf-8")))
        except (OSError, ValueError):
            pass

    def __getitem__(self, key):
        return self.values[key]

    def __setitem__(self, key, value) -> None:
        self.values[key] = value
        try:
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            SETTINGS_FILE.write_text(json.dumps(self.values, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass


class AiState:
    """Startet den KI-Server im Hintergrund und merkt sich, wie es steht."""

    def __init__(self) -> None:
        self.text = "KI: nicht eingerichtet"
        self.ok = False
        self._thread: threading.Thread | None = None

    def start(self, model: Path | None, context: int) -> None:
        if model is None:
            self.text, self.ok = "KI: kein Modell gefunden", False
            return
        if self._thread is not None and self._thread.is_alive():
            return
        self.text, self.ok = f"KI wird geladen …\n{model.stem}", False

        def run() -> None:
            try:
                aw.start_ai(model, context)
                self.text, self.ok = f"KI bereit\n{model.stem} · Kontext {aw.AI.context}", True
            except Exception as exc:
                self.text, self.ok = f"KI-Fehler: {exc}"[:300], False

        self._thread = threading.Thread(target=run, name="ki-start", daemon=True)
        self._thread.start()

    @property
    def loading(self) -> bool:
        return self._thread is not None and self._thread.is_alive()


# --- Seite: Aufnehmen ---------------------------------------------------------
STATE_ICONS = {"wartet": "⏳", "läuft": "⚙", "fertig": "✓", "fehler": "⚠"}


class RecordPage(QWidget):
    def __init__(self, colors: dict, settings: Settings, auswerter: aw.Auswerter) -> None:
        super().__init__()
        self.colors = colors
        self.settings = settings
        self.auswerter = auswerter
        self.recorder = LoopbackRecorder()
        auswerter.limit_threads = lambda: self.recorder.running
        self._shown_result = None

        root = QVBoxLayout(self)
        root.setContentsMargins(32, 28, 32, 28)
        root.setSpacing(18)
        root.addWidget(label("Vorlesung aufnehmen", "PageTitle"))
        root.addWidget(label("Nimmt auf, was über Lautsprecher oder Headset läuft – zum Beispiel "
                             "den LearnSpace3D. Jeder fertige Teil wird sofort ausgewertet. "
                             "Alles bleibt auf diesem PC.", "Muted", wrap=True))
        self.banner = label("", "Banner", wrap=True)
        self.banner.setVisible(False)
        root.addWidget(self.banner)

        body = QHBoxLayout()
        body.setSpacing(18)
        root.addLayout(body, 1)

        # links: Einstellungen der Aufnahme
        setup, setup_layout = card()
        setup_layout.addWidget(label("Thema / Fach"))
        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("z. B. Netzwerktechnik – Subnetting")
        setup_layout.addWidget(self.title_edit)
        self.mic_check = QCheckBox("Mein Mikrofon auch aufnehmen (erkennt „Du“ und „Lehrkraft“)")
        self.mic_check.setToolTip("Damit deine Fragen und Beiträge in den Notizen stehen. Das Transkript zeigt dann, "
                                  "was du gesagt hast („Du:“) und was die Lehrkraft („Lehrkraft:“). Am besten mit Headset.")
        self.mic_check.setChecked(bool(settings["with_mic"]))
        self.mic_check.toggled.connect(lambda v: settings.__setitem__("with_mic", v))
        setup_layout.addWidget(self.mic_check)
        split_row = QHBoxLayout()
        split_row.addWidget(label("Neuer Teil alle"))
        self.split_choice = QComboBox()
        for minutes in (0, 15, 30, 45, 60, 90):
            self.split_choice.addItem("nur bei Pause" if minutes == 0 else f"{minutes} Minuten", minutes)
        index = self.split_choice.findData(settings["split_minutes"])
        self.split_choice.setCurrentIndex(index if index >= 0 else 2)
        self.split_choice.currentIndexChanged.connect(
            lambda _: settings.__setitem__("split_minutes", self.split_choice.currentData()))
        split_row.addWidget(self.split_choice)
        split_row.addStretch()
        setup_layout.addLayout(split_row)
        setup_layout.addSpacing(6)
        self.consent_check = QCheckBox("Ich habe die Erlaubnis zur Aufnahme")
        self.consent_check.toggled.connect(self._update_buttons)
        setup_layout.addWidget(self.consent_check)
        setup_layout.addWidget(label("Von der Lehrkraft bzw. der WBS, nur für meine persönlichen "
                                     "Notizen. Heimliche Aufnahmen sind strafbar (§ 201 StGB).", "Muted", wrap=True))
        setup_layout.addStretch()
        self.existing_button = QPushButton("Vorhandene Aufnahme auswerten …")
        self.existing_button.setToolTip("Eine oder mehrere .wav-Dateien (z. B. Teile einer Vorlesung) auswerten.")
        self.existing_button.clicked.connect(self.evaluate_existing)
        setup_layout.addWidget(self.existing_button)
        setup.setFixedWidth(360)
        body.addWidget(setup)

        # rechts: Aufnahme-Steuerung und Auswertung
        control, control_layout = card()
        control_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.state_label = label("Bereit", "Muted")
        self.state_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        control_layout.addWidget(self.state_label)
        self.timer_label = label("00:00", "Timer")
        self.timer_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        control_layout.addWidget(self.timer_label)

        self.record_button = QPushButton("●\nStart")
        self.record_button.setObjectName("Record")
        self.record_button.setFixedSize(96, 96)
        self.record_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.record_button.clicked.connect(self.start)
        control_layout.addWidget(self.record_button, 0, Qt.AlignmentFlag.AlignHCenter)

        buttons = QHBoxLayout()
        self.pause_button = QPushButton("Pause")
        self.pause_button.setToolTip("Pausiert und wertet den bisherigen Teil schon aus.")
        self.pause_button.clicked.connect(self.toggle_pause)
        self.stop_button = QPushButton("Stoppen")
        self.stop_button.setObjectName("Stop")
        self.stop_button.clicked.connect(self.stop)
        buttons.addWidget(self.pause_button)
        buttons.addWidget(self.stop_button)
        control_layout.addLayout(buttons)

        self.speaker_bar = self._level_bar(control_layout, "Lautsprecher / Headset")
        self.mic_bar = self._level_bar(control_layout, "Mikrofon")

        control_layout.addWidget(label("Teile und Auswertung", "Muted"))
        self.parts_list = QListWidget()
        self.parts_list.setMinimumHeight(110)
        self.parts_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.parts_list.setWordWrap(True)
        control_layout.addWidget(self.parts_list, 1)

        result_row = QHBoxLayout()
        self.hint = label("Der erste Teil wird ausgewertet, sobald du Pause drückst oder die "
                          "eingestellte Zeit erreicht ist.", "Muted", wrap=True)
        result_row.addWidget(self.hint, 1)
        self.open_result_button = QPushButton("Notizen öffnen")
        self.open_result_button.setObjectName("Primary")
        self.open_result_button.setVisible(False)
        self.open_result_button.clicked.connect(self.open_result)
        result_row.addWidget(self.open_result_button)
        self.claude_result_button = QPushButton("Für Claude kopieren")
        self.claude_result_button.setToolTip("Kopiert das Transkript mit einer Anweisung für claude.ai (kostenlos, von Hand).")
        self.claude_result_button.setVisible(False)
        self.claude_result_button.clicked.connect(
            lambda: self.auswerter.result and self._show_info(copy_for_claude(Path(self.auswerter.result["path"]))))
        result_row.addWidget(self.claude_result_button)
        control_layout.addLayout(result_row)
        body.addWidget(control, 1)

        self.ticker = QTimer(self)
        self.ticker.timeout.connect(self.refresh)
        self.ticker.start(300)
        self._update_buttons()

    def _level_bar(self, layout: QVBoxLayout, text: str) -> QProgressBar:
        row = QHBoxLayout()
        name = label(text, "Muted")
        name.setFixedWidth(170)
        bar = QProgressBar()
        bar.setRange(0, 100)
        bar.setTextVisible(False)
        bar.setFixedHeight(8)
        row.addWidget(name)
        row.addWidget(bar, 1)
        layout.addLayout(row)
        return bar

    @property
    def can_evaluate(self) -> bool:
        return aw.PROGRAM_DIR is not None

    # -- Steuerung --
    def start(self) -> None:
        RECORDINGS_DIR.mkdir(parents=True, exist_ok=True)
        title = self.title_edit.text().strip()
        name = f"{dt.datetime.now():%Y-%m-%d_%H%M}_{safe_filename(title)}.wav"
        self.parts_list.clear()
        self.open_result_button.setVisible(False)
        self.claude_result_button.setVisible(False)
        self._info = ""
        self.auswerter.reset(title)
        self.recorder.start(RECORDINGS_DIR / name, with_mic=self.mic_check.isChecked(),
                            auto_split_minutes=self.split_choice.currentData(),
                            on_part=self.auswerter.submit if self.can_evaluate else None)
        self._update_buttons()

    def toggle_pause(self) -> None:
        if self.recorder.paused:
            self.recorder.resume()
        else:
            self.recorder.pause()
        self._update_buttons()

    def stop(self) -> None:
        self.recorder.stop()
        if not self.recorder.got_audio:
            self.auswerter.reset(self.auswerter.title)
            QMessageBox.warning(self, APP_NAME, "Es kam kein Ton an. Lief die Vorlesung über das "
                                "Standard-Ausgabegerät von Windows?\n\n" + self.recorder.error)
        elif self.can_evaluate:
            self.auswerter.finish()
        self.refresh()
        self._update_buttons()

    def evaluate_existing(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Aufnahme auswählen", str(RECORDINGS_DIR),
                                                "Aufnahmen (*.wav *.mp3 *.m4a)")
        if not files:
            return
        paths = sorted(Path(f) for f in files)
        title = self.title_edit.text().strip() or re.sub(r"^\d{4}-\d{2}-\d{2}_\d{4}_|_Teil\d+$", "",
                                                           paths[0].stem).replace("_", " ")
        self.parts_list.clear()
        self.open_result_button.setVisible(False)
        self.claude_result_button.setVisible(False)
        self._info = ""
        self.auswerter.reset(title)
        offset = 0.0
        for number, path in enumerate(paths, start=1):
            self.auswerter.submit(path, offset, number)
            offset += wav_seconds(path)
        self.auswerter.finish()
        self._update_buttons()

    def _show_info(self, text: str) -> None:
        self._info = text
        self.hint.setObjectName("Muted")
        self.hint.setText(text)

    def open_result(self) -> None:
        if self.auswerter.result:
            open_path(self.auswerter.result["path"])

    def _update_buttons(self) -> None:
        running = self.recorder.running
        busy = self.auswerter.busy
        self.record_button.setVisible(not running)
        self.record_button.setEnabled(IS_WINDOWS and self.consent_check.isChecked() and not busy)
        if busy:
            self.record_button.setToolTip("Die vorherige Aufnahme wird noch ausgewertet.")
        elif not self.consent_check.isChecked():
            self.record_button.setToolTip("Bitte zuerst die Erlaubnis bestätigen.")
        else:
            self.record_button.setToolTip("")
        self.pause_button.setVisible(running)
        self.stop_button.setVisible(running)
        self.pause_button.setText("Weiter" if self.recorder.paused else "Pause")
        for widget in (self.title_edit, self.mic_check, self.split_choice, self.consent_check):
            widget.setEnabled(not running)
        self.existing_button.setEnabled(not running and not busy and self.can_evaluate)
        self.mic_bar.setEnabled(self.mic_check.isChecked())
        problems = []
        if not IS_WINDOWS:
            problems.append("Die Aufnahme funktioniert nur unter Windows.")
        if not self.can_evaluate:
            problems.append("Auswertung nicht möglich: Der Ordner der bisherigen GespraechsnotizAI-Installation "
                            "wurde nicht gefunden. Bitte unter Einstellungen auswählen – aufnehmen geht trotzdem.")
        self.banner.setText("\n".join(problems))
        self.banner.setVisible(bool(problems))

    # -- Anzeige --
    def refresh(self) -> None:
        rec, ev = self.recorder, self.auswerter
        if rec.running:
            self.state_label.setText("⏸ Pausiert" if rec.paused else
                                     f"<span style='color:{self.colors['rec']}'>●</span> Aufnahme läuft")
            self.timer_label.setText(fmt_time(rec.wall_seconds))
        elif ev.busy:
            self.state_label.setText("Auswertung läuft …")
        elif self.state_label.text() != "Bereit":
            self.state_label.setText("Bereit")
        live = rec.running and not rec.paused
        self.speaker_bar.setValue(int(min(rec.levels.get("speaker", 0.0) * 400, 100)) if live else 0)
        self.mic_bar.setValue(int(min(rec.levels.get("mic", 0.0) * 400, 100)) if live else 0)

        # Liste der Teile mit ihrem Auswertungsstand
        while self.parts_list.count() < len(ev.parts):
            self.parts_list.addItem(QListWidgetItem())
        for row, part in enumerate(list(ev.parts)):
            length = wav_seconds(part["path"]) if part["state"] != "läuft" else 0
            text = (f"{STATE_ICONS.get(part['state'], '•')}  Teil {part['number']}  ·  ab {fmt_time(part['offset'])}"
                    + (f"  ·  {fmt_time(length)}" if length else "") + f"\n      {part['status']}")
            item = self.parts_list.item(row)
            if item.text() != text:
                item.setText(text)
                item.setToolTip(str(part["path"]))

        # Hinweis bzw. Ergebnis unten
        if ev.error:
            self.hint.setObjectName("Error")
            self.hint.setText(ev.error)
        elif ev.result:
            self.hint.setObjectName("Result")
            self.hint.setText(getattr(self, "_info", "") or f"Fertig! Notizen gespeichert: {Path(ev.result['path']).name}")
            self.open_result_button.setVisible(True)
            self.claude_result_button.setVisible(True)
            if self._shown_result is not ev.result:
                self._shown_result = ev.result
                self._update_buttons()
        elif ev.closed:
            self.hint.setObjectName("Muted")
            self.hint.setText("Endfassung wird erstellt … " + (ev.status or ""))
        elif rec.error and rec.running:
            self.hint.setObjectName("Error")
            self.hint.setText(rec.error)
        elif rec.running or not ev.parts:
            self.hint.setObjectName("Muted")
            self.hint.setText("Der erste Teil wird ausgewertet, sobald du Pause drückst oder die "
                              "eingestellte Zeit erreicht ist." if not ev.parts else
                              "Die fertigen Teile werden im Hintergrund ausgewertet.")
        self.hint.style().unpolish(self.hint)
        self.hint.style().polish(self.hint)
        if not rec.running and self.record_button.isEnabled() == ev.busy:
            self._update_buttons()


# --- Seite: Notizen -------------------------------------------------------------
class NotesPage(QWidget):
    def __init__(self, share: DiscordShare | None = None) -> None:
        super().__init__()
        self.share = share
        self.notes = []
        self.visible = []
        root = QVBoxLayout(self)
        root.setContentsMargins(32, 28, 32, 28)
        root.setSpacing(14)
        header = QHBoxLayout()
        header.addWidget(label("Meine Notizen", "PageTitle"))
        header.addStretch()
        folder_button = QPushButton("Ordner öffnen")
        folder_button.clicked.connect(lambda: (NOTES_DIR.mkdir(parents=True, exist_ok=True), open_path(NOTES_DIR)))
        header.addWidget(folder_button)
        root.addLayout(header)

        self.search = QLineEdit()
        self.search.setPlaceholderText("In allen Notizen suchen, z. B. „Subnetting“ …")
        self.search.textChanged.connect(self.fill_list)
        root.addWidget(self.search)

        splitter = QSplitter()
        self.list = QListWidget()
        self.list.currentRowChanged.connect(self.show_note)
        self.list.itemDoubleClicked.connect(lambda _: self.open_current())
        splitter.addWidget(self.list)
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        self.preview = QTextBrowser()
        right_layout.addWidget(self.preview, 1)
        self.info = label("", "Muted", wrap=True)
        self.info.setVisible(False)
        right_layout.addWidget(self.info)
        buttons = QHBoxLayout()
        self.claude_button = QPushButton("Für Claude kopieren")
        self.claude_button.setToolTip("Kopiert das Transkript mit einer fertigen Anweisung und öffnet claude.ai "
                                      "(kostenlos mit einem Gratis-Konto).")
        self.claude_button.clicked.connect(self.copy_current_for_claude)
        buttons.addWidget(self.claude_button)
        self.save_claude_button = QPushButton("Claude-Antwort speichern")
        self.save_claude_button.setToolTip("Speichert die kopierte Antwort von claude.ai als Word-Dokument bei den Notizen.")
        self.save_claude_button.clicked.connect(self.save_claude_answer)
        buttons.addWidget(self.save_claude_button)
        self.share_claude = QCheckBox("auch in Discord teilen")
        buttons.addWidget(self.share_claude)
        buttons.addStretch()
        self.open_button = QPushButton("In Word öffnen")
        self.open_button.setObjectName("Primary")
        self.open_button.clicked.connect(self.open_current)
        buttons.addWidget(self.open_button)
        right_layout.addLayout(buttons)
        splitter.addWidget(right)
        splitter.setSizes([320, 680])
        root.addWidget(splitter, 1)

    def reload(self) -> None:
        self.notes = list_notes(NOTES_DIR)
        self.fill_list()
        self.share_claude.setVisible(bool(self.share and self.share.active))

    def show_info(self, text: str) -> None:
        self.info.setText(text)
        self.info.setVisible(bool(text))

    def copy_current_for_claude(self) -> None:
        row = self.list.currentRow()
        if 0 <= row < len(self.visible):
            self.show_info(copy_for_claude(self.visible[row].path))

    def save_claude_answer(self) -> None:
        text = QGuiApplication.clipboard().text()
        if not claude_chat.looks_like_answer(text):
            self.show_info("In der Zwischenablage ist noch keine Antwort von Claude. Auf claude.ai unter der "
                           "Antwort auf „Kopieren“ klicken und dann erneut hier speichern.")
            return
        row = self.list.currentRow()
        original = self.visible[row].path if 0 <= row < len(self.visible) else NOTES_DIR / "Vorlesung.docx"
        try:
            path, title = claude_chat.save_answer(text, original, NOTES_DIR)
        except Exception as exc:
            self.show_info(f"Konnte nicht gespeichert werden: {exc}")
            return
        shared = ""
        if self.share_claude.isChecked() and self.share and self.share.active:
            self.share.finished(title, text, path)
            shared = " und in Discord geteilt"
        self.reload()
        for index, note in enumerate(self.visible):
            if note.path == path:
                self.list.setCurrentRow(index)
        self.show_info(f"Gespeichert{shared}: {path.name}")

    def fill_list(self) -> None:
        query = self.search.text()
        self.list.clear()
        self.visible = [n for n in self.notes if n.matches(query)]
        for note in self.visible:
            short = re.sub(r"^(Vorlesungsnotizen|Notizen nach Folien):\s*", "", note.title)
            item = QListWidgetItem(f"{short}\n{note.modified:%d.%m.%Y, %H:%M} Uhr")
            item.setToolTip(note.title)
            item.setSizeHint(QSize(0, 54))
            self.list.addItem(item)
        if self.visible:
            self.list.setCurrentRow(0)
        else:
            text = "Keine Notizen zu dieser Suche." if query else (
                "Noch keine Notizen. Sobald eine Vorlesung ausgewertet ist, erscheint sie hier.")
            self.preview.setHtml(f"<p style='color:gray'>{text}</p>")
        self.open_button.setEnabled(bool(self.visible))
        self.claude_button.setEnabled(bool(self.visible))

    def show_note(self, row: int) -> None:
        if 0 <= row < len(self.visible):
            self.preview.setHtml(self.visible[row].to_html())

    def open_current(self) -> None:
        row = self.list.currentRow()
        if 0 <= row < len(self.visible):
            open_path(self.visible[row].path)


# --- Seite: Einstellungen -------------------------------------------------------
class SettingsPage(QWidget):
    def __init__(self, settings: Settings, ai_state: AiState, auswerter: aw.Auswerter,
                 is_busy, share: DiscordShare) -> None:
        super().__init__()
        self.share = share
        self.settings = settings
        self.ai_state = ai_state
        self.auswerter = auswerter
        self.is_busy = is_busy
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        content = QWidget()
        scroll.setWidget(content)
        outer.addWidget(scroll)
        root = QVBoxLayout(content)
        root.setContentsMargins(32, 28, 32, 28)
        root.setSpacing(14)
        root.addWidget(label("Einstellungen", "PageTitle"))

        ai_box, ai = card()
        ai.addWidget(label("Auswertung", ""))
        program_row = QHBoxLayout()
        program_text = QVBoxLayout()
        program_text.addWidget(label("Programmordner von GespraechsnotizAI (KI, Modelle, ffmpeg)"))
        self.program_label = path_label("")
        program_text.addWidget(self.program_label)
        program_row.addLayout(program_text, 1)
        change = QPushButton("Ändern …")
        change.clicked.connect(self.choose_program_dir)
        program_row.addWidget(change)
        ai.addLayout(program_row)

        grid = QHBoxLayout()
        model_col = QVBoxLayout()
        model_col.addWidget(label("KI-Modell"))
        self.model_choice = QComboBox()
        model_col.addWidget(self.model_choice)
        grid.addLayout(model_col, 2)
        context_col = QVBoxLayout()
        context_col.addWidget(label("Kontextgröße"))
        self.context_choice = QComboBox()
        for value in (4096, 8192, 16384):
            self.context_choice.addItem(f"{value:,}".replace(",", "."), value)
        self.context_choice.setCurrentIndex(max(0, self.context_choice.findData(settings["context"])))
        context_col.addWidget(self.context_choice)
        grid.addLayout(context_col, 1)
        ai.addLayout(grid)

        ai.addWidget(label("Spracherkennung (Whisper)"))
        self.whisper_choice = QComboBox()
        for key, text in WHISPER_CHOICES.items():
            self.whisper_choice.addItem(text, key)
        self.whisper_choice.setCurrentIndex(max(0, self.whisper_choice.findData(settings["whisper"])))
        ai.addWidget(self.whisper_choice)

        apply_row = QHBoxLayout()
        self.ai_label = label("", "Muted", wrap=True)
        apply_row.addWidget(self.ai_label, 1)
        self.apply_button = QPushButton("Übernehmen")
        self.apply_button.setObjectName("Primary")
        self.apply_button.clicked.connect(self.apply)
        apply_row.addWidget(self.apply_button)
        ai.addLayout(apply_row)
        root.addWidget(ai_box)

        share_box, share_layout = card()
        share_layout.addWidget(label("Mit der Gruppe teilen (Discord)"))
        self.share_check = QCheckBox("Notizen automatisch in unseren Discord-Kanal schicken")
        self.share_check.setChecked(bool(settings["discord_on"]))
        self.share_check.toggled.connect(lambda v: settings.__setitem__("discord_on", v))
        share_layout.addWidget(self.share_check)
        url_row = QHBoxLayout()
        self.url_edit = QLineEdit(settings["discord_url"])
        self.url_edit.setPlaceholderText("Webhook-Adresse, beginnt mit https://discord.com/api/webhooks/ …")
        self.url_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.url_edit.editingFinished.connect(self.save_url)
        url_row.addWidget(self.url_edit, 1)
        test_button = QPushButton("Test senden")
        test_button.clicked.connect(self.send_test)
        url_row.addWidget(test_button)
        share_layout.addLayout(url_row)
        share_layout.addWidget(label("Geteilt werden nur Notizen und Zusammenfassungen – jeder fertige Teil schon "
                                     "während der Vorlesung, am Ende das Word-Dokument. Keine Aufnahmen, keine "
                                     "Transkripte. Die Webhook-Adresse nicht öffentlich weitergeben.", "Muted", wrap=True))
        self.share_label = label("", "Muted", wrap=True)
        share_layout.addWidget(self.share_label)
        root.addWidget(share_box)

        folders, layout = card()
        for title, path in (("Aufnahmen", RECORDINGS_DIR), ("Notizen", NOTES_DIR)):
            row = QHBoxLayout()
            text = QVBoxLayout()
            text.addWidget(label(title))
            text.addWidget(path_label(str(path)))
            row.addLayout(text, 1)
            button = QPushButton("Öffnen")
            button.clicked.connect(lambda _=False, p=path: (p.mkdir(parents=True, exist_ok=True), open_path(p)))
            row.addWidget(button)
            layout.addLayout(row)
        root.addWidget(folders)
        root.addStretch()
        self.fill_models()

        timer = QTimer(self)
        timer.timeout.connect(self.refresh)
        timer.start(500)

    def fill_models(self) -> None:
        self.program_label.setText(str(aw.PROGRAM_DIR) if aw.PROGRAM_DIR else "Nicht gefunden – bitte auswählen.")
        self.program_label.setToolTip(self.program_label.text())
        self.model_choice.clear()
        models = aw.available_models()
        for model in models:
            self.model_choice.addItem(aw.model_label(model), str(model))
        chosen = next((m for m in models if m.name == self.settings["model"]), None) or aw.preferred_model(models)
        if chosen is not None:
            self.model_choice.setCurrentIndex(self.model_choice.findData(str(chosen)))

    def choose_program_dir(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Ordner mit GespraechsnotizAI.exe auswählen",
                                                  str(aw.PROGRAM_DIR or Path.home()))
        if not folder:
            return
        if not aw.configure_program_dir(Path(folder)):
            QMessageBox.warning(self, APP_NAME, "In diesem Ordner wurde keine GespraechsnotizAI-Installation "
                                "gefunden. Bitte den Ordner wählen, in dem GespraechsnotizAI.exe liegt.")
            return
        self.settings["program_dir"] = str(aw.PROGRAM_DIR)
        self.fill_models()
        self.apply()

    def apply(self) -> None:
        if self.is_busy():
            QMessageBox.information(self, APP_NAME, "Bitte warten, bis Aufnahme und Auswertung fertig sind.")
            return
        model = Path(self.model_choice.currentData()) if self.model_choice.currentData() else None
        self.settings["model"] = model.name if model else ""
        self.settings["context"] = self.context_choice.currentData()
        self.settings["whisper"] = self.whisper_choice.currentData()
        self.auswerter.whisper_model = self.settings["whisper"]
        self.auswerter.ai_model = model
        self.auswerter.ai_context = self.settings["context"]
        if model is not None and aw.PROGRAM_DIR is not None:
            if not aw._SERVER.running_with(model, 2048 if model.name == aw.BUNDLED_MODEL else self.settings["context"]):
                aw.stop_ai()
                self.ai_state.start(model, self.settings["context"])

    def save_url(self) -> None:
        url = self.url_edit.text().strip()
        if url and not is_webhook(url):
            QMessageBox.warning(self, APP_NAME, "Das ist keine Discord-Webhook-Adresse. Sie beginnt mit "
                                "https://discord.com/api/webhooks/ – siehe Anleitung.")
            return
        self.settings["discord_url"] = url

    def send_test(self) -> None:
        self.save_url()
        if not is_webhook(self.settings["discord_url"]):
            QMessageBox.information(self, APP_NAME, "Bitte zuerst die Webhook-Adresse einfügen.")
            return
        self.share.status = "Discord: Test wird gesendet …"
        self.share.test(self.settings["discord_url"])

    def refresh(self) -> None:
        self.ai_label.setText(self.ai_state.text.replace("\n", " · "))
        self.apply_button.setEnabled(not self.ai_state.loading)
        self.share_label.setText(self.share.status)


# --- Hauptfenster ---------------------------------------------------------------
def app_icon(color: str) -> QIcon:
    pixmap = QPixmap(64, 64)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setBrush(QColor(color))
    painter.setPen(Qt.PenStyle.NoPen)
    painter.drawRoundedRect(4, 4, 56, 56, 14, 14)
    painter.setPen(QColor("#ffffff"))
    painter.setFont(QFont("Segoe UI", 26, QFont.Weight.Bold))
    painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, "G")
    painter.end()
    return QIcon(pixmap)


class MainWindow(QMainWindow):
    def __init__(self, colors: dict, start_ai: bool = True) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(app_icon(colors["accent"]))
        self.resize(1180, 780)
        self.setMinimumSize(980, 660)

        self.settings = Settings()
        program_dir = Path(self.settings["program_dir"]) if self.settings["program_dir"] else None
        if not (program_dir and aw.configure_program_dir(program_dir)):
            found = aw.find_program_dir(Path(__file__).resolve().parent)
            if found and aw.configure_program_dir(found):
                self.settings["program_dir"] = str(found)
        self.auswerter = aw.Auswerter()
        self.ai_state = AiState()
        self.share = DiscordShare(lambda: self.settings["discord_url"], lambda: bool(self.settings["discord_on"]))
        self.auswerter.on_part_done = self.share.part_done
        self.auswerter.on_finished = self.share.finished

        central = QWidget()
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        side = QWidget()
        side.setObjectName("Nav")
        side.setFixedWidth(220)
        side_layout = QVBoxLayout(side)
        side_layout.setContentsMargins(0, 0, 0, 0)
        side_layout.setSpacing(0)
        side_layout.addWidget(label(APP_NAME, "Brand"))
        side_layout.addWidget(label("Lernnotizen aus Vorlesungen", "BrandSub"))
        self.nav = QListWidget()
        self.nav.setObjectName("Nav")
        for text in ("🎙  Aufnehmen", "📝  Notizen", "⚙  Einstellungen"):
            self.nav.addItem(text)
        side_layout.addWidget(self.nav, 1)
        self.ai_label = label("", "AiStatus", wrap=True)
        side_layout.addWidget(self.ai_label)
        layout.addWidget(side)

        self.pages = QStackedWidget()
        self.record_page = RecordPage(colors, self.settings, self.auswerter)
        self.notes_page = NotesPage(self.share)
        self.settings_page = SettingsPage(self.settings, self.ai_state, self.auswerter,
                                          lambda: self.record_page.recorder.running or self.auswerter.busy,
                                          self.share)
        self.pages.addWidget(self.record_page)
        self.pages.addWidget(self.notes_page)
        self.pages.addWidget(self.settings_page)
        layout.addWidget(self.pages, 1)
        self.setCentralWidget(central)

        self.nav.currentRowChanged.connect(self.switch_page)
        self.nav.setCurrentRow(0)
        status_timer = QTimer(self)
        status_timer.timeout.connect(lambda: self.ai_label.setText(
            self.ai_state.text + ("\nTeilen in Discord: an" if self.share.active else "")))
        status_timer.start(500)
        if start_ai:
            self.settings_page.apply()  # KI im Hintergrund vorladen, damit sie bereit ist, wenn Teil 1 fertig ist

    def switch_page(self, index: int) -> None:
        self.pages.setCurrentIndex(index)
        if index == 1:
            self.notes_page.reload()

    def closeEvent(self, event) -> None:
        recorder = self.record_page.recorder
        if recorder.running or self.auswerter.busy:
            text = ("Es läuft noch eine Aufnahme. Aufnahme speichern und beenden?" if recorder.running else
                    "Die Auswertung läuft noch und würde abgebrochen. Trotzdem beenden?\n"
                    "(Die Aufnahme bleibt gespeichert und kann später ausgewertet werden.)")
            if QMessageBox.question(self, APP_NAME, text) != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
            if recorder.running:
                recorder.stop()
        aw.stop_ai()
        event.accept()


def main() -> int:
    if IS_WINDOWS:  # eigenes Taskleisten-Symbol statt dem von Python
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Gespraechsnotiz.AI.Desktop")
        except Exception:
            pass
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    scheme = QGuiApplication.styleHints().colorScheme()
    colors = THEMES["dark" if scheme == Qt.ColorScheme.Dark else "light"]
    app.setStyle("Fusion")
    app.setStyleSheet(build_stylesheet(colors))
    window = MainWindow(colors)
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
