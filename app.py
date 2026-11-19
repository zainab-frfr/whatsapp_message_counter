import io
import re
import zipfile

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
COLUMNS = ["Audio File", "Text", "Sender", "Timestamp"]


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


def safe_sheet_name(name, used):
    base = re.sub(r"[\[\]:*?/\\]", "", name.rsplit(".", 1)[0])[:28] or "Sheet"
    candidate, n = base, 2
    while candidate in used:
        candidate = f"{base}_{n}"
        n += 1
    used.add(candidate)
    return candidate


st.title("WhatsApp Audio ↔ Text to Excel")
st.caption("Upload one or more Android WhatsApp chat exports (.zip).")

files = st.file_uploader("Chat zip files", type="zip", accept_multiple_files=True)

if files:
    sheets, used = {}, set()
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
        rows = build_rows(parse_chat(text))
        sheets[safe_sheet_name(f.name, used)] = pd.DataFrame(rows, columns=COLUMNS)
        st.write(f"**{f.name}**: {len(rows)} rows")

    if sheets:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            for name, df in sheets.items():
                df.to_excel(writer, sheet_name=name, index=False)
        st.download_button(
            "Download Excel",
            data=buf.getvalue(),
            file_name="whatsapp_audio_text.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        first_name = next(iter(sheets))
        st.subheader(f"Preview: {first_name}")
        st.dataframe(sheets[first_name].head(50))
