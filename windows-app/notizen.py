# -*- coding: utf-8 -*-
"""Notiz-Archiv: findet die erstellten Word-Dokumente und liest ihren Text für Vorschau und Suche.

Kommt ohne python-docx aus – eine .docx-Datei ist ein Zip-Archiv mit XML.
"""
from __future__ import annotations

import datetime as dt
import html
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@dataclass
class Notiz:
    path: Path
    title: str
    modified: dt.datetime
    paragraphs: list[tuple[str, str]] = field(default_factory=list)  # (Stil, Text)

    @property
    def text(self) -> str:
        return "\n".join(t for _, t in self.paragraphs)

    def matches(self, query: str) -> bool:
        query = query.strip().lower()
        return not query or query in self.title.lower() or query in self.text.lower()

    def to_html(self) -> str:
        parts, in_list = [], False
        for style, text in self.paragraphs:
            safe = html.escape(text)
            is_item = style.startswith("List")
            if is_item and not in_list:
                parts.append("<ul>")
            if not is_item and in_list:
                parts.append("</ul>")
            in_list = is_item
            if style == "Title":
                parts.append(f"<h1>{safe}</h1>")
            elif style == "Heading1":
                parts.append(f"<h2>{safe}</h2>")
            elif style in ("Heading2", "Heading3"):
                parts.append(f"<h3>{safe}</h3>")
            elif is_item:
                parts.append(f"<li>{safe}</li>")
            else:
                parts.append(f"<p>{safe}</p>")
        if in_list:
            parts.append("</ul>")
        return "\n".join(parts)


def read_docx(path: Path) -> list[tuple[str, str]]:
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("word/document.xml"))
    paragraphs = []
    for para in root.iter(f"{_W}p"):
        text = "".join(t.text or "" for t in para.iter(f"{_W}t")).strip()
        if not text:
            continue
        style_el = para.find(f"{_W}pPr/{_W}pStyle")
        style = style_el.get(f"{_W}val", "") if style_el is not None else ""
        paragraphs.append((style, text))
    return paragraphs


def list_notes(folder: Path) -> list[Notiz]:
    notes = []
    for path in folder.glob("*.docx") if folder.exists() else []:
        if path.name.startswith("~$"):  # Sperrdatei von Word
            continue
        try:
            paragraphs = read_docx(path)
        except (OSError, KeyError, zipfile.BadZipFile, ElementTree.ParseError):
            continue
        title = next((t for s, t in paragraphs if s == "Title"), path.stem)
        modified = dt.datetime.fromtimestamp(path.stat().st_mtime)
        notes.append(Notiz(path, title, modified, paragraphs))
    return sorted(notes, key=lambda n: n.modified, reverse=True)
