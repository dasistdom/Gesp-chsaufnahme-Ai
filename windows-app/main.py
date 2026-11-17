# -*- coding: utf-8 -*-
"""Gesprächsnotiz-AI – Windows-Programm (Demo).

Eigene Oberfläche mit Qt statt Browser. In dieser ersten Demo:
  • Aufnehmen: Systemton (LearnSpace3D) + optional Mikrofon, Pause, automatische Teile
  • Notizen: alle erstellten Word-Dokumente durchsuchen, ansehen und in Word öffnen
  • Einstellungen: Ordner

Die Auswertung (Spracherkennung + KI) läuft in dieser Demo noch über die bisherige
Website-Version: dort unter „Ältere Aufnahme auswerten“. Beide nutzen dieselben Ordner.
"""
from __future__ import annotations

import datetime as dt
import os
import re
import subprocess
import sys
import threading
from pathlib import Path

from PySide6.QtCore import QSize, Qt, QTimer, QUrl
from PySide6.QtGui import QColor, QDesktopServices, QFont, QGuiApplication, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFrame, QHBoxLayout, QLabel, QLineEdit, QListWidget,
    QListWidgetItem, QMainWindow, QMessageBox, QProgressBar, QPushButton,
    QSplitter, QStackedWidget, QTextBrowser, QVBoxLayout, QWidget,
)

from aufnahme import RECORD_RATE, LoopbackRecorder
from notizen import list_notes

APP_NAME = "Gesprächsnotiz-AI"
DATA_DIR = Path.home() / "Documents" / "GespraechsnotizAI"  # dieselben Ordner wie die Website-Version
RECORDINGS_DIR = DATA_DIR / "Aufnahmen"
NOTES_DIR = DATA_DIR / "Notizen"
IS_WINDOWS = sys.platform.startswith("win")

# --- Gestaltung ---------------------------------------------------------------
# Ruhiges Grau mit leichtem Blaustich, Indigo als Akzent, Rot nur für die Aufnahme.
THEMES = {
    "light": {"bg": "#f4f5f8", "panel": "#ffffff", "nav": "#eceef4", "text": "#1d2130",
              "muted": "#646b80", "line": "#dde0ea", "accent": "#4f5bd5", "accent_text": "#ffffff",
              "rec": "#e5484d", "ok": "#2f9e6e", "hover": "#e2e5f0"},
    "dark": {"bg": "#16181f", "panel": "#1e212b", "nav": "#121419", "text": "#e7e9f0",
             "muted": "#9aa0b4", "line": "#2c303d", "accent": "#7a84f0", "accent_text": "#0f1016",
             "rec": "#ff6369", "ok": "#4cc38a", "hover": "#262a36"},
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
    #Card {{ background: {c['panel']}; border: 1px solid {c['line']}; border-radius: 12px; }}
    #Card QWidget {{ background: {c['panel']}; }}
    #PageTitle {{ font-size: 18pt; font-weight: 700; }}
    #Muted {{ color: {c['muted']}; }}
    #Timer {{ font-size: 34pt; font-weight: 600; font-family: 'Segoe UI', 'Inter', sans-serif; }}
    #Banner {{ background: {c['hover']}; border-radius: 8px; padding: 8px 12px; color: {c['muted']}; }}
    QLineEdit {{ background: {c['bg']}; border: 1px solid {c['line']}; border-radius: 8px;
                           padding: 7px 10px; }}
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


# --- Seite: Aufnehmen ---------------------------------------------------------
class RecordPage(QWidget):
    def __init__(self, colors: dict) -> None:
        super().__init__()
        self.colors = colors
        self.recorder = LoopbackRecorder()
        self.finished_parts: list[tuple[Path, float, int]] = []
        self._lock = threading.Lock()

        root = QVBoxLayout(self)
        root.setContentsMargins(32, 28, 32, 28)
        root.setSpacing(18)
        root.addWidget(label("Vorlesung aufnehmen", "PageTitle"))
        root.addWidget(label("Nimmt auf, was über Lautsprecher oder Headset läuft – zum Beispiel "
                             "den LearnSpace3D. Alles bleibt auf diesem PC.", "Muted", wrap=True))

        if not IS_WINDOWS:
            root.addWidget(label("Die Aufnahme funktioniert nur unter Windows.", "Banner"))

        body = QHBoxLayout()
        body.setSpacing(18)
        root.addLayout(body, 1)

        # links: Einstellungen der Aufnahme
        setup, setup_layout = card()
        setup_layout.addWidget(label("Thema / Fach"))
        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("z. B. Netzwerktechnik – Subnetting")
        setup_layout.addWidget(self.title_edit)
        self.mic_check = QCheckBox("Mein Mikrofon auch aufnehmen")
        self.mic_check.setToolTip("Damit deine Fragen und Beiträge in den Notizen stehen. Am besten mit Headset.")
        setup_layout.addWidget(self.mic_check)
        split_row = QHBoxLayout()
        split_row.addWidget(label("Neuer Teil alle"))
        self.split_choice = QComboBox()
        for minutes in (0, 15, 30, 45, 60, 90):
            self.split_choice.addItem("nur bei Pause" if minutes == 0 else f"{minutes} Minuten", minutes)
        self.split_choice.setCurrentIndex(2)
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
        setup.setFixedWidth(360)
        body.addWidget(setup)

        # rechts: Aufnahme-Steuerung
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
        self.pause_button.clicked.connect(self.toggle_pause)
        self.stop_button = QPushButton("Stoppen")
        self.stop_button.setObjectName("Stop")
        self.stop_button.clicked.connect(self.stop)
        buttons.addWidget(self.pause_button)
        buttons.addWidget(self.stop_button)
        control_layout.addLayout(buttons)

        self.speaker_bar = self._level_bar(control_layout, "Lautsprecher / Headset")
        self.mic_bar = self._level_bar(control_layout, "Mikrofon")

        control_layout.addWidget(label("Gespeicherte Teile", "Muted"))
        self.parts_list = QListWidget()
        self.parts_list.setMinimumHeight(110)
        self.parts_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        control_layout.addWidget(self.parts_list, 1)
        self.hint = label("Zum Auswerten in der Website-Version unter „Ältere Aufnahme auswerten“ "
                          "den Teil wählen. In der nächsten Version passiert das hier automatisch.",
                          "Muted", wrap=True)
        control_layout.addWidget(self.hint)
        body.addWidget(control, 1)

        self.ticker = QTimer(self)
        self.ticker.timeout.connect(self.refresh)
        self.ticker.start(250)
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

    # -- Steuerung --
    def _on_part(self, path: Path, start: float, number: int) -> None:  # läuft im Aufnahme-Thread
        with self._lock:
            self.finished_parts.append((path, start, number))

    def start(self) -> None:
        RECORDINGS_DIR.mkdir(parents=True, exist_ok=True)
        name = f"{dt.datetime.now():%Y-%m-%d_%H%M}_{safe_filename(self.title_edit.text())}.wav"
        self.parts_list.clear()
        self.recorder.start(RECORDINGS_DIR / name, with_mic=self.mic_check.isChecked(),
                            auto_split_minutes=self.split_choice.currentData(), on_part=self._on_part)
        self._update_buttons()

    def toggle_pause(self) -> None:
        if self.recorder.paused:
            self.recorder.resume()
        else:
            self.recorder.pause()
        self._update_buttons()

    def stop(self) -> None:
        self.recorder.stop()
        self.refresh()
        if not self.recorder.got_audio:
            QMessageBox.warning(self, APP_NAME, "Es kam kein Ton an. Lief die Vorlesung über das "
                                "Standard-Ausgabegerät von Windows?\n\n" + self.recorder.error)
        self._update_buttons()

    def _update_buttons(self) -> None:
        running = self.recorder.running
        self.record_button.setVisible(not running)
        self.record_button.setEnabled(IS_WINDOWS and self.consent_check.isChecked())
        self.record_button.setToolTip("" if self.consent_check.isChecked() else "Bitte zuerst die Erlaubnis bestätigen.")
        self.pause_button.setVisible(running)
        self.stop_button.setVisible(running)
        self.pause_button.setText("Weiter" if self.recorder.paused else "Pause")
        for widget in (self.title_edit, self.mic_check, self.split_choice, self.consent_check):
            widget.setEnabled(not running)
        self.mic_bar.setEnabled(self.mic_check.isChecked())

    # -- Anzeige alle 250 ms --
    def refresh(self) -> None:
        rec = self.recorder
        if rec.running:
            self.state_label.setText("⏸ Pausiert" if rec.paused else
                                     f"<span style='color:{self.colors['rec']}'>●</span> Aufnahme läuft")
            self.timer_label.setText(fmt_time(rec.wall_seconds))
        elif self.state_label.text() != "Bereit":
            self.state_label.setText("Bereit")
        self.speaker_bar.setValue(int(min(rec.levels.get("speaker", 0.0) * 400, 100)))
        self.mic_bar.setValue(int(min(rec.levels.get("mic", 0.0) * 400, 100)))
        with self._lock:
            new_parts, self.finished_parts = self.finished_parts, []
        for path, start, number in new_parts:
            seconds = 0.0
            try:
                seconds = path.stat().st_size / (RECORD_RATE * 2)
            except OSError:
                pass
            item = QListWidgetItem(f"✓  Teil {number}   ·   ab {fmt_time(start)}   ·   Länge {fmt_time(seconds)}")
            item.setToolTip(str(path))
            self.parts_list.addItem(item)
        if rec.error and rec.running:
            self.hint.setText(rec.error)


# --- Seite: Notizen -------------------------------------------------------------
class NotesPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.notes = []
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
        self.open_button = QPushButton("In Word öffnen")
        self.open_button.setObjectName("Primary")
        self.open_button.clicked.connect(self.open_current)
        right_layout.addWidget(self.open_button, 0, Qt.AlignmentFlag.AlignRight)
        splitter.addWidget(right)
        splitter.setSizes([320, 680])
        root.addWidget(splitter, 1)

    def reload(self) -> None:
        self.notes = list_notes(NOTES_DIR)
        self.fill_list()

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

    def show_note(self, row: int) -> None:
        if 0 <= row < len(self.visible):
            self.preview.setHtml(self.visible[row].to_html())

    def open_current(self) -> None:
        row = self.list.currentRow()
        if 0 <= row < len(self.visible):
            open_path(self.visible[row].path)


# --- Seite: Einstellungen -------------------------------------------------------
class SettingsPage(QWidget):
    def __init__(self) -> None:
        super().__init__()
        root = QVBoxLayout(self)
        root.setContentsMargins(32, 28, 32, 28)
        root.setSpacing(14)
        root.addWidget(label("Einstellungen", "PageTitle"))
        box, layout = card()
        for title, path in (("Aufnahmen", RECORDINGS_DIR), ("Notizen", NOTES_DIR)):
            row = QHBoxLayout()
            text = QVBoxLayout()
            text.addWidget(label(title))
            text.addWidget(label(str(path), "Muted"))
            row.addLayout(text, 1)
            button = QPushButton("Öffnen")
            button.clicked.connect(lambda _=False, p=path: (p.mkdir(parents=True, exist_ok=True), open_path(p)))
            row.addWidget(button)
            layout.addLayout(row)
        root.addWidget(box)
        root.addWidget(label("Demo-Version. Die Website-Version und dieses Programm nutzen dieselben "
                             "Ordner und können parallel verwendet werden.", "Muted", wrap=True))
        root.addStretch()


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
    def __init__(self, colors: dict) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(app_icon(colors["accent"]))
        self.resize(1180, 760)
        self.setMinimumSize(980, 640)

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
        layout.addWidget(side)

        self.pages = QStackedWidget()
        self.record_page = RecordPage(colors)
        self.notes_page = NotesPage()
        self.pages.addWidget(self.record_page)
        self.pages.addWidget(self.notes_page)
        self.pages.addWidget(SettingsPage())
        layout.addWidget(self.pages, 1)
        self.setCentralWidget(central)

        self.nav.currentRowChanged.connect(self.switch_page)
        self.nav.setCurrentRow(0)

    def switch_page(self, index: int) -> None:
        self.pages.setCurrentIndex(index)
        if index == 1:
            self.notes_page.reload()

    def closeEvent(self, event) -> None:
        if self.record_page.recorder.running:
            answer = QMessageBox.question(
                self, APP_NAME, "Es läuft noch eine Aufnahme. Aufnahme speichern und beenden?")
            if answer != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
            self.record_page.recorder.stop()
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
