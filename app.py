import io
import re
import shutil
import zipfile
from pathlib import PurePosixPath

import pandas as pd
import streamlit as st

# Android export line, e.g. "12/31/23, 10:15 PM - Sender: message"
HEADER = re.compile(
    r"^(\d{1,2}[/.]\d{1,2}[/.]\d{2,4}),?\s+"
    r"(\d{1,2}:\d{2}(?::\d{2})?(?:\s?[AaPp]\.?[Mm]\.?)?)\s+-\s+(.*)$"
)
AUDIO_EXT = r"opus|aac|mp3|m4a|amr|ogg|wav|mmr|3gp|wma|flac"
AUDIO = re.compile(rf"^(.+?\.(?:{AUDIO_EXT}))\s+\(file attached\)$", re.I)
ATTACH = re.compile(r"\(file attached\)\s*$", re.I)
COLUMNS = ["Audio File", "Text", "Sender", "Timestamp", "Renamed File"]


def parse_chat(text):
    """Return a list of messages: {ts, sender, lines}."""
    text = text.replace("\u202f", " ").replace("\u200e", "").replace("\ufeff", "")
    messages, current = [], None
    for raw in text.splitlines():
        m = HEADER.match(raw)
        if m:
            date, time, rest = m.groups()
            if ": " in rest:
                sender, body = rest.split(": ", 1)
                current = {"ts": f"{date} {time}", "sender": sender, "lines": [body]}
                messages.append(current)
            else:
                current = None  # system message (encryption notice, joins, etc.)
        elif current is not None:
            current["lines"].append(raw)  # continuation of a multi-line message
    return messages


def build_rows(messages):
    rows, audio = [], None
    for msg in messages:
        lines = msg["lines"]
        first = lines[0].strip()

        if first.lower() == "<media omitted>":
            continue
        am = AUDIO.match(first)
        if am:
            if audio:  # previous audio never got a text
                rows.append([audio, "", "", ""])
            audio = am.group(1)
            lines = lines[1:]
        elif ATTACH.search(first):
            continue  # non-audio attachment, ignore

        lines = [l.strip() for l in lines if l.strip()]
        for i, line in enumerate(lines):
            rows.append([audio if i == 0 and audio else "", line, msg["sender"], msg["ts"]])
            if i == 0:
                audio = None

    if audio:
        rows.append([audio, "", "", ""])
    return rows


def clean_name(text):
    """Make text safe to use as a file name."""
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    return text[:100]


def add_renamed(rows):
    """Add a 5th value per row: new file name = first line of text + original extension.
    Audio with no text keeps its original name. Duplicate names get _2, _3, ..."""
    used, out = set(), []
    for audio, text, sender, ts in rows:
        new = ""
        if audio:
            ext = PurePosixPath(audio).suffix
            stem = clean_name(text) if text else ""
            if stem:
                new, n = stem + ext, 2
                while new.lower() in used:
                    new = f"{stem}_{n}{ext}"
                    n += 1
            else:
                new = audio
            used.add(new.lower())
        out.append([audio, text, sender, ts, new])
    return out


def safe_sheet_name(name, used):
    base = re.sub(r"[\[\]:*?/\\]", "", name.rsplit(".", 1)[0])[:28] or "Sheet"
    candidate, n = base, 2
    while candidate in used:
        candidate = f"{base}_{n}"
        n += 1
    used.add(candidate)
    return candidate


def make_excel(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            for row in ws.iter_rows():
                for cell in row:
                    cell.number_format = "@"  # Excel "Text" format
                    if cell.value is not None:
                        cell.value = str(cell.value)
                        cell.data_type = "s"  # stored as a string, never a number
    return buf.getvalue()


def build_zip(chats, excel_bytes):
    """Zip = Excel at the root + one folder per chat holding the renamed audio."""
    buf = io.BytesIO()
    missing = 0
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as out:  # audio is already compressed
        out.writestr("whatsapp_audio_text.xlsx", excel_bytes)
        for sheet, f, rows in chats:
            with zipfile.ZipFile(f) as zin:
                by_name = {
                    PurePosixPath(n).name: n for n in zin.namelist() if not n.endswith("/")
                }
                for audio, _, _, _, new in rows:
                    if not audio:
                        continue
                    src = by_name.get(audio)
                    if src is None:
                        missing += 1
                        continue
                    with zin.open(src) as s, out.open(f"{sheet}/{new}", "w") as d:
                        shutil.copyfileobj(s, d)
    return buf.getvalue(), missing


st.title("WhatsApp Audio ↔ Text to Excel")
st.caption("Upload one or more Android WhatsApp chat exports (.zip).")

files = st.file_uploader("Chat zip files", type="zip", accept_multiple_files=True)

if files:
    sheets, chats, used = {}, [], set()
    for f in files:
        try:
            zf = zipfile.ZipFile(f)
        except zipfile.BadZipFile:
            st.warning(f"{f.name}: not a valid zip file, skipped.")
            continue
        txts = [
            n for n in zf.namelist()
            if n.lower().endswith(".txt") and not n.startswith("__MACOSX")
        ]
        if not txts:
            st.warning(f"{f.name}: no .txt chat file found, skipped.")
            continue
        text = zf.read(txts[0]).decode("utf-8", errors="ignore")
        rows = add_renamed(build_rows(parse_chat(text)))
        sheet = safe_sheet_name(f.name, used)
        sheets[sheet] = pd.DataFrame(rows, columns=COLUMNS)
        chats.append((sheet, f, rows))
        st.write(f"**{f.name}**: {len(rows)} rows")

    if sheets:
        excel_bytes = make_excel(sheets)
        st.download_button(
            "Download Excel only",
            data=excel_bytes,
            file_name="whatsapp_audio_text.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        first_name = next(iter(sheets))
        st.subheader(f"Preview: {first_name}")
        st.dataframe(sheets[first_name].head(50), hide_index=True)

        st.divider()
        sig = tuple((f.name, f.size) for f in files)
        if st.button("Build zip with renamed audio + Excel"):
            with st.spinner("Building zip..."):
                data, missing = build_zip(chats, excel_bytes)
            st.session_state["zip"] = (sig, data, missing)

        saved = st.session_state.get("zip")
        if saved and saved[0] == sig:
            if saved[2]:
                st.warning(
                    f"{saved[2]} audio file(s) named in the chat were not found inside "
                    "the zip (exported without media?), so they are not in the output zip."
                )
            st.download_button(
                "Download renamed audio + Excel (.zip)",
                data=saved[1],
                file_name="renamed_audio.zip",
                mime="application/zip",
            )
