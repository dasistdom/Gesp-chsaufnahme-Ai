# ENTWURF – noch NICHT in app.py eingebaut und noch nicht getestet.
# Idee: Unterrichtsmaterial (PowerPoint/PDF/Word aus dem WBS eCampus) einem Vorlesungstag
# zuordnen, das Gesagte den Folien zuordnen und Notizen pro Folie schreiben.
# Siehe Fahrplan-Tabelle, Zeile "Notizen automatisch auf die PowerPoint abstimmen".
#
# Zum Einbauen in app.py nötig: Imports (posixpath, xml.etree.ElementTree as ElementTree),
# save_session(...) am Ende von run_pipeline und LiveProcessor._do_final aufrufen,
# und einen Tab "Unterrichtsmaterial" anlegen, der run_material_pipeline(...) startet.

# ---------------------------------------------------------------------------
# Gespeicherte Auswertungen („Sitzungen“) – damit Material später dazukommen kann
# ---------------------------------------------------------------------------
SESSIONS_DIR = DATA_DIR / "Sitzungen"
MATERIAL_DIR = DATA_DIR / "Material"


def save_session(title: str, created: dt.datetime, segments: list[dict], parts: list[dict],
                 summary: str) -> Path | None:
    """Speichert Transkript und Notizen einer Auswertung, damit Unterrichtsmaterial später dazukommen kann."""
    if not segments:
        return None
    try:
        SESSIONS_DIR.mkdir(parents=True, exist_ok=True)
        path = SESSIONS_DIR / f"{created:%Y-%m-%d_%H%M%S}_{safe_filename(title)}.json"
        data = {
            "title": title, "created": created.isoformat(timespec="seconds"),
            "date": created.date().isoformat(), "duration": segments[-1]["end"],
            "segments": segments, "summary": summary,
            "parts": [{k: p.get(k, "") for k in ("label", "text", "notes")} for p in parts],
        }
        path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        return path
    except OSError:
        return None


def load_sessions() -> list[dict]:
    sessions = []
    for path in sorted(SESSIONS_DIR.glob("*.json")) if SESSIONS_DIR.exists() else []:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            data["path"] = str(path)
            sessions.append(data)
        except (OSError, ValueError):
            continue
    return sessions


def sessions_by_day() -> dict[str, list[dict]]:
    days: dict[str, list[dict]] = {}
    for session in load_sessions():
        days.setdefault(session["date"], []).append(session)
    return dict(sorted(days.items(), reverse=True))


_WEEKDAYS = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]


def day_label(day: str, sessions: list[dict]) -> str:
    date = dt.date.fromisoformat(day)
    titles = ", ".join(dict.fromkeys(s["title"] for s in sessions if s.get("title"))) or "ohne Thema"
    hours = sum(s.get("duration", 0) for s in sessions)
    count = f"{len(sessions)} Aufnahmen" if len(sessions) > 1 else "1 Aufnahme"
    return f"{_WEEKDAYS[date.weekday()]}, {date:%d.%m.%Y} – {titles} ({count}, {fmt_time(hours)})"


def day_segments(sessions: list[dict]) -> list[dict]:
    """Alle Transkript-Abschnitte eines Tages; bei mehreren Aufnahmen mit Nummer der Aufnahme."""
    segments = []
    for number, session in enumerate(sorted(sessions, key=lambda s: s["created"]), start=1):
        for seg in session["segments"]:
            segments.append({**seg, "rec": number if len(sessions) > 1 else 0})
    return segments


def seg_time(seg: dict) -> str:
    return (f"A{seg['rec']} " if seg.get("rec") else "") + fmt_time(seg["start"])


# ---------------------------------------------------------------------------
# Unterrichtsmaterial einlesen (PowerPoint, PDF, Word)
# ---------------------------------------------------------------------------
_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _pptx_rels(archive, part: str) -> dict[str, str]:
    folder, name = part.rsplit("/", 1)
    rels_path = f"{folder}/_rels/{name}.rels"
    if rels_path not in archive.namelist():
        return {}
    root = ElementTree.fromstring(archive.read(rels_path))
    result = {}
    for rel in root.findall("rel:Relationship", _NS):
        target = rel.get("Target", "")
        if rel.get("TargetMode") == "External":
            continue
        result[rel.get("Id")] = posixpath.normpath(posixpath.join(folder, target)) \
            if not target.startswith("/") else target.lstrip("/")
    return result


def _shape_texts(root, skip_placeholders: tuple[str, ...] = ()) -> tuple[str, list[str]]:
    """Gibt (Titel, Textzeilen) einer Folie oder Notizseite zurück."""
    title, lines = "", []
    for shape in root.iter(f"{{{_NS['p']}}}sp"):
        placeholder = shape.find(".//p:nvPr/p:ph", _NS)
        kind = placeholder.get("type", "body") if placeholder is not None else ""
        if kind in skip_placeholders:
            continue
        paragraphs = []
        for para in shape.iter(f"{{{_NS['a']}}}p"):
            text = "".join(t.text or "" for t in para.iter(f"{{{_NS['a']}}}t")).strip()
            if text:
                paragraphs.append(text)
        if not paragraphs:
            continue
        if kind in ("title", "ctrTitle") and not title:
            title = " ".join(paragraphs)
        else:
            lines.extend(paragraphs)
    for frame in root.iter(f"{{{_NS['a']}}}tbl"):  # Tabellen auf Folien
        for row in frame.iter(f"{{{_NS['a']}}}tr"):
            cells = ["".join(t.text or "" for t in cell.iter(f"{{{_NS['a']}}}t")).strip()
                     for cell in row.iter(f"{{{_NS['a']}}}tc")]
            if any(cells):
                lines.append(" | ".join(cells))
    return title, lines


def extract_pptx(data: bytes, source: str) -> list[dict]:
    import zipfile
    slides = []
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        presentation = ElementTree.fromstring(archive.read("ppt/presentation.xml"))
        rels = _pptx_rels(archive, "ppt/presentation.xml")
        order = [rels.get(s.get(f"{{{_NS['r']}}}id"))
                 for s in presentation.findall("p:sldIdLst/p:sldId", _NS)]
        for number, part in enumerate([p for p in order if p], start=1):
            root = ElementTree.fromstring(archive.read(part))
            if root.get("show") == "0":  # ausgeblendete Folie
                continue
            title, lines = _shape_texts(root)
            notes = ""
            for target in _pptx_rels(archive, part).values():
                if "notesSlide" in target and target in archive.namelist():
                    _, note_lines = _shape_texts(ElementTree.fromstring(archive.read(target)),
                                                 skip_placeholders=("sldImg", "sldNum", "hdr", "ftr", "dt"))
                    notes = "\n".join(note_lines)
            slides.append({"number": number, "title": title or (lines[0][:80] if lines else ""),
                           "text": "\n".join(lines), "notes": notes, "source": source})
    return slides


def extract_material(name: str, data: bytes) -> list[dict]:
    suffix = Path(name).suffix.lower()
    if suffix == ".pptx":
        return extract_pptx(data, name)
    if suffix == ".pdf":
        reader = PdfReader(io.BytesIO(data))
        pages = []
        for number, page in enumerate(reader.pages, start=1):
            text = (page.extract_text() or "").strip()
            if text:
                first = text.splitlines()[0].strip()
                pages.append({"number": number, "title": first[:80], "text": text, "notes": "", "source": name})
        return pages
    if suffix == ".docx":
        doc = Document(io.BytesIO(data))
        sections, current = [], {"title": "", "lines": []}
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            if para.style is not None and para.style.name.lower().startswith(("heading", "überschrift", "title")):
                if current["lines"] or current["title"]:
                    sections.append(current)
                current = {"title": text, "lines": []}
            else:
                current["lines"].append(text)
        if current["lines"] or current["title"]:
            sections.append(current)
        return [{"number": i, "title": s["title"] or (s["lines"][0][:80] if s["lines"] else ""),
                 "text": "\n".join(s["lines"]), "notes": "", "source": name}
                for i, s in enumerate(sections, start=1)]
    if suffix == ".ppt":
        raise ValueError("Alte PowerPoint-Dateien (.ppt) werden nicht unterstützt. Bitte in PowerPoint "
                         "als .pptx oder PDF speichern.")
    raise ValueError(f"Dateityp {suffix} wird nicht unterstützt.")


# ---------------------------------------------------------------------------
# Vorlesung den Folien zuordnen
# ---------------------------------------------------------------------------
_STOPWORDS = set("""aber alle allem allen aller alles also auch auf aus bei beim bin bis bist bitte dabei dafür
damit dann darauf darum dass dein deine dem den denn der des dessen deshalb die dies diese diesem diesen
dieser dieses doch dort durch eben eigentlich ein eine einem einen einer eines einfach einmal etwas euch
euer eure für gibt ganz gerade gleich habe haben hast hat hatte hier hinter ich ihr ihre immer ist jetzt
kann kannst kein keine können könnt machen macht mal man mehr mein meine mich mir mit muss müssen nach
nicht nichts noch nun nur oder ohne quasi schon sehr sein seine sich sie sind soll sollen sondern sowie
über also okay unter und uns unser vielleicht viel vom von vor wann war waren warum was weil weiter
welche wenn wer werden wie wieder will wir wird wirklich wo wollen worden wurde würde zum zur zwischen
genau halt sozusagen natürlich richtig okay folie folien""".split())


def _tokens(text: str) -> list[str]:
    words = re.findall(r"[a-zäöüß0-9][a-zäöüß0-9\-]{2,}", text.lower())
    return [w[:7] for w in words if w not in _STOPWORDS and not w.isdigit()]  # grobe Wortstämme


def align_to_slides(segments: list[dict], slides: list[dict], window_seconds: float = 45.0) -> list[dict]:
    """Ordnet Zeitfenster der Vorlesung den Folien zu – in Folienreihenfolge, Folien dürfen übersprungen werden."""
    import math
    from collections import Counter

    windows: list[dict] = []
    for seg in segments:
        if not windows or seg.get("rec") != windows[-1]["rec"] \
                or seg["end"] - windows[-1]["start"] > window_seconds:
            windows.append({"rec": seg.get("rec", 0), "start": seg["start"], "end": seg["end"],
                            "segments": [], "text": ""})
        windows[-1]["segments"].append(seg)
        windows[-1]["end"] = seg["end"]
        windows[-1]["text"] += " " + seg["text"]
    if not windows or not slides:
        return windows

    slide_tokens = [Counter(_tokens(f"{s['title']} {s['title']} {s['text']} {s['notes']}")) for s in slides]
    doc_freq = Counter(t for counts in slide_tokens for t in counts)
    idf = {t: math.log((1 + len(slides)) / (1 + df)) + 1 for t, df in doc_freq.items()}

    def vector(counts: Counter) -> dict[str, float]:
        vec = {t: (1 + math.log(c)) * idf[t] for t, c in counts.items() if t in idf}
        norm = math.sqrt(sum(v * v for v in vec.values())) or 1.0
        return {t: v / norm for t, v in vec.items()}

    slide_vecs = [vector(c) for c in slide_tokens]
    win_vecs = [vector(Counter(_tokens(w["text"]))) for w in windows]
    score = [[sum(v * sv.get(t, 0.0) for t, v in wv.items()) for sv in slide_vecs] for wv in win_vecs]

    # Dynamische Programmierung: Folien laufen nur vorwärts; Überspringen kostet etwas
    skip_penalty, n_slides = 0.02, len(slides)
    dp = [[0.0] * n_slides for _ in windows]
    back = [[0] * n_slides for _ in windows]
    for j in range(n_slides):
        dp[0][j] = score[0][j] - skip_penalty * j
    for i in range(1, len(windows)):
        best_value, best_k = float("-inf"), 0
        for j in range(n_slides):
            candidate = dp[i - 1][j] + skip_penalty * j  # bester Vorgänger k <= j: dp[k] - p*(j-k)
            if candidate > best_value:
                best_value, best_k = candidate, j
            dp[i][j] = score[i][j] + best_value - skip_penalty * j
            back[i][j] = best_k
    j = max(range(n_slides), key=lambda k: dp[-1][k])
    for i in range(len(windows) - 1, -1, -1):
        windows[i]["slide"] = j
        windows[i]["score"] = score[i][j]
        j = back[i][j]
    return windows


# ---------------------------------------------------------------------------
# Notizen pro Folie
# ---------------------------------------------------------------------------
SLIDE_INSTRUCTIONS = """Oben stehen Folien aus dem Unterrichtsmaterial und darunter jeweils, was in der Vorlesung
dazu gesagt wurde (automatisches Transkript mit Hörfehlern).

Schreibe für JEDE Folie Lernnotizen in genau diesem Format:

### Folie <Nummer>
- 2 bis 6 Stichpunkte: was die Folie aussagt, ergänzt um Erklärungen und Beispiele aus der Vorlesung
- Beginne einen Punkt mit „Zusätzlich:“, wenn er nur mündlich erklärt wurde und nicht auf der Folie steht
- Beginne einen Punkt mit „Wichtig:“ bei Prüfungshinweisen, Aufgaben oder Terminen

Verwende Fachbegriffe in der Schreibweise der Folien – das Transkript schreibt sie oft falsch.
Wurde zu einer Folie nichts gesagt, fasse nur die Folie kurz zusammen. Erfinde nichts."""


def _slide_block(slide: dict, spoken: str, times: str) -> str:
    block = f"### Folie {slide['number']}: {slide['title']}\nAuf der Folie:\n{slide['text'][:1200]}\n"
    if slide["notes"]:
        block += f"Sprechernotizen:\n{slide['notes'][:600]}\n"
    block += f"In der Vorlesung gesagt{f' ({times})' if times else ''}:\n{spoken or '(nichts erkannt)'}\n"
    return block


def _parse_slide_notes(text: str) -> dict[int, str]:
    result: dict[int, str] = {}
    current = None
    for line in text.splitlines():
        match = re.match(r"^\s*#{1,4}\s*Folie\s*(\d+)", line, re.IGNORECASE)
        if match:
            current = int(match.group(1))
            result.setdefault(current, "")
            continue
        if current is not None:
            result[current] += line + "\n"
    return {k: v.strip() for k, v in result.items()}


def slide_notes(slides: list[dict], windows: list[dict], title: str,
                progress: Callable[[float, str], None]) -> list[dict]:
    """Ergänzt jede Folie um das Gesagte, die Zeitangaben und KI-Notizen."""
    for index, slide in enumerate(slides):
        own = [w for w in windows if w.get("slide") == index]
        slide["spoken"] = " ".join(w["text"].strip() for w in own)
        slide["times"] = f"{seg_time(own[0]['segments'][0])} – {fmt_time(own[-1]['end'])}" if own else ""
        slide["match"] = max((w["score"] for w in own), default=0.0)
        slide["notes_ai"] = ""

    # Folien zu Gruppen bündeln, die in eine KI-Anfrage passen; zu langes Gesagtes wird aufgeteilt
    budget = AI.chunk_chars
    jobs: list[tuple[list[dict], str]] = []
    group, size = [], 0
    for slide in slides:
        spoken_parts = split_text(slide["spoken"], max(800, budget - 1500)) if slide["spoken"] else [""]
        for part_index, spoken in enumerate(spoken_parts):
            block = _slide_block(slide, spoken, slide["times"] if part_index == 0 else "Fortsetzung")
            if group and size + len(block) > budget:
                jobs.append((group, "".join(b for _, b in group)))
                group, size = [], 0
            group.append((slide, block))
            size += len(block)
    if group:
        jobs.append((group, "".join(b for _, b in group)))

    topic = f" zum Thema „{title}“" if title else ""
    for number, (members, text) in enumerate(jobs, start=1):
        progress((number - 1) / len(jobs), f"Folien werden ausgewertet ({number} von {len(jobs)}) …")
        answer = ask_ai_safe(lambda t: f"Unterrichtsmaterial und Vorlesung{topic}:\n\n<<<\n{t}\n>>>\n\n{SLIDE_INSTRUCTIONS}",
                             text, AI.notes_tokens + 200)
        parsed = _parse_slide_notes(answer)
        for slide, _ in members:
            note = parsed.get(slide["number"], "")
            if note and note not in slide["notes_ai"]:
                slide["notes_ai"] = (slide["notes_ai"] + "\n" + note).strip()
        if not parsed and len(members) == 1:  # Modell hat das Format ignoriert
            members[0][0]["notes_ai"] = (members[0][0]["notes_ai"] + "\n" + answer).strip()
    progress(1.0, "Folien fertig.")
    return slides


def build_slides_docx(title: str, day: str, summary: str, slides: list[dict], windows: list[dict],
                      segments: list[dict], material_names: list[str], coverage: float) -> bytes:
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    date = dt.date.fromisoformat(day)
    doc.add_heading(f"Notizen nach Folien: {title}" if title else "Notizen nach Folien", level=0)
    doc.add_paragraph(f"Vorlesung vom {_WEEKDAYS[date.weekday()]}, {date:%d.%m.%Y} · Material: "
                      + ", ".join(material_names))
    hint = doc.add_paragraph(
        "Automatisch erstellt. Die Zuordnung von Gesagtem zu Folien ist eine Schätzung "
        f"(etwa {coverage:.0%} der Vorlesung ließ sich eindeutig einer Folie zuordnen). "
        "„Zusätzlich:“ = nur mündlich erklärt, „Wichtig:“ = Prüfungshinweis, Aufgabe oder Termin.")
    hint.runs[0].italic = True

    doc.add_heading("Zusammenfassung", level=1)
    _add_markdown(doc, summary or "Keine Zusammenfassung erstellt.")

    doc.add_heading("Notizen zu jeder Folie", level=1)
    for slide in slides:
        heading = f"{slide['source'] if len(material_names) > 1 else 'Folie'} {slide['number']}"
        doc.add_heading(f"{heading}: {slide['title']}" if slide["title"] else heading, level=2)
        info = f"Besprochen: {slide['times']}" if slide["times"] else "In der Aufnahme nicht erkennbar besprochen"
        meta = doc.add_paragraph(info)
        meta.runs[0].italic = True
        _add_markdown(doc, slide.get("notes_ai") or "\n".join(f"- {l}" for l in slide["text"].splitlines()[:8]))

    if segments:
        doc.add_page_break()
        doc.add_heading("Transkript mit Folienbezug", level=1)
        for window in windows:
            slide = slides[window["slide"]] if "slide" in window and window["score"] > 0.05 else None
            prefix = f"[{seg_time(window['segments'][0])}]"
            if slide:
                prefix += f" (Folie {slide['number']})"
            doc.add_paragraph(f"{prefix} {window['text'].strip()}")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def run_material_pipeline(day: str, sessions: list[dict], files: list[tuple[str, bytes]]) -> None:
    title = ", ".join(dict.fromkeys(s["title"] for s in sessions if s.get("title")))
    with st.status("Unterrichtsmaterial wird mit der Vorlesung abgeglichen …", expanded=True) as status:
        slides: list[dict] = []
        for name, data in files:
            try:
                found = extract_material(name, data)
            except Exception as exc:
                st.warning(f"{name}: {exc}")
                continue
            if not found:
                st.warning(f"{name}: Kein Text gefunden. Bei eingescannten PDFs fehlt eine Texterkennung.")
            slides.extend(found)
        if not slides:
            status.update(label="Kein verwertbares Material", state="error")
            return
        st.write(f"📑 {len(slides)} Folien/Seiten eingelesen.")

        # Material für diesen Tag aufbewahren
        try:
            folder = MATERIAL_DIR / day
            folder.mkdir(parents=True, exist_ok=True)
            for name, data in files:
                (folder / Path(name).name).write_bytes(data)
        except OSError:
            pass

        segments = day_segments(sessions)
        windows = align_to_slides(segments, slides)
        matched = [w for w in windows if w.get("score", 0) > 0.05]
        coverage = (sum(w["end"] - w["start"] for w in matched)
                    / max(sum(w["end"] - w["start"] for w in windows), 1)) if windows else 0.0
        st.write(f"🔗 Etwa {coverage:.0%} der Vorlesung passt eindeutig zu einer Folie.")
        if coverage < 0.25:
            st.warning("Aufnahme und Material passen kaum zusammen. Ist es das richtige Material für diesen Tag?")

        bar = st.progress(0.0)
        try:
            slide_notes(slides, windows, title, lambda v, t: bar.progress(min(v * 0.85, 1.0), text=t))
            pseudo_parts = [{"notes": s["notes_ai"]} for s in slides if s.get("notes_ai")]
            summary = summarize_parts(pseudo_parts, title,
                                      lambda v, t: bar.progress(min(0.85 + v * 0.15, 1.0), text=t))
        except Exception as exc:
            status.update(label="KI-Auswertung fehlgeschlagen", state="error")
            st.exception(exc)
            return

        docx_bytes = build_slides_docx(title, day, summary, slides, windows, segments,
                                       [Path(n).name for n, _ in files], coverage)
        filename = f"{day}_{safe_filename(title)}_Folien.docx"
        saved_to = ""
        try:
            NOTES_DIR.mkdir(parents=True, exist_ok=True)
            (NOTES_DIR / filename).write_bytes(docx_bytes)
            saved_to = str(NOTES_DIR / filename)
        except OSError:
            pass
        status.update(label="Fertig! Die Notizen nach Folien sind unten.", state="complete", expanded=False)

    st.session_state["result"] = {
        "title": title, "summary": summary,
        "parts": [{"label": f"Folie {s['number']}: {s['title']}" + (f" · {s['times']}" if s["times"] else ""),
                   "notes": s.get("notes_ai", "")} for s in slides],
        "transcript": "\n".join(f"[{seg_time(s)}] {s['text']}" for s in segments),
        "docx": docx_bytes, "filename": filename, "saved_to": saved_to,
    }


