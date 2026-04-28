"""
app.py
Audio / Video → Executive Summary + Minutes of Meeting + Feedback Email
Transasia Biomedicals Ltd.
"""

import os
import re
import tempfile
import subprocess
from io import BytesIO
from pathlib import Path

import streamlit as st
import anthropic
from docx import Document
from docx.shared import Pt
from groq import Groq
from dotenv import load_dotenv

# ── API keys (Streamlit Cloud secrets → local .env fallback) ──
try:
    ANTHROPIC_API_KEY = st.secrets["ANTHROPIC_API_KEY"]
    GROQ_API_KEY      = st.secrets["GROQ_API_KEY"]
except Exception:
    load_dotenv()
    ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
    GROQ_API_KEY      = os.environ.get("GROQ_API_KEY", "")

# ── Page config ───────────────────────────────────────────────
st.set_page_config(
    page_title="Audio/Video to Text Data Conversion",
    page_icon="🎙️",
    layout="centered",
)

st.title("Audio/Video to Text Data Conversion")
st.caption("This app transforms the audio in the video and audio notes to its Summary, MOM and Email text for further records and usage.")

# ── Sidebar ───────────────────────────────────────────────────
with st.sidebar:
    st.header("Configuration")
    language = st.selectbox("Transcription language", ["en", "hi", "auto"], index=0)
    st.divider()
    st.subheader("Generate")
    do_summary = st.checkbox("📋 Summary",          value=True)
    do_mom     = st.checkbox("📝 Minutes of Meeting", value=True)
    do_email   = st.checkbox("📧 Feedback Email",   value=True)
    if not any([do_summary, do_mom, do_email]):
        st.warning("Select at least one output.")
    st.divider()
    st.subheader("Document metadata")
    sender_name      = st.text_input("Your name",    "Kshitij Buch")
    sender_title     = st.text_input("Your title",   "Digital Projects & Technical Support Manager")
    company          = st.text_input("Company",      "Transasia Biomedicals Ltd.")
    recipient_name   = st.text_input("Recipient",    "Kshitij Buch")
    email_subject    = st.text_input("Email subject","Feedback – XL200 Troubleshooting Agent CRU Error")
    feedback_context = st.text_input("Context",      "XL200 Troubleshooting Agent Response Feedback")

# ── Markdown → .docx helper ──────────────────────────────────
def _add_inline(para, text: str):
    """Write text into a paragraph, honouring **bold** and *italic* markers."""
    for part in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text):
        if part.startswith("**") and part.endswith("**"):
            para.add_run(part[2:-2]).bold = True
        elif part.startswith("*") and part.endswith("*"):
            para.add_run(part[1:-1]).italic = True
        elif part:
            para.add_run(part)


def markdown_to_docx(markdown_text: str, title: str = "") -> bytes:
    doc = Document()
    if title:
        doc.add_heading(title, level=0)

    lines = markdown_text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]

        if line.startswith("### "):
            doc.add_heading(line[4:].strip(), level=3)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(), level=2)
        elif line.startswith("# "):
            doc.add_heading(line[2:].strip(), level=1)
        elif line.startswith("|"):
            # Collect all consecutive table lines
            tbl_lines = []
            while i < len(lines) and lines[i].startswith("|"):
                tbl_lines.append(lines[i])
                i += 1
            # Split each row into cells (trim leading/trailing |)
            rows = [
                [c.strip() for c in row.strip().strip("|").split("|")]
                for row in tbl_lines
            ]
            # Remove separator rows (cells contain only dashes/spaces)
            rows = [r for r in rows if not all(re.fullmatch(r"[-: ]+", c) for c in r)]
            if rows:
                ncols = max(len(r) for r in rows)
                tbl = doc.add_table(rows=len(rows), cols=ncols)
                tbl.style = "Table Grid"
                for ri, row in enumerate(rows):
                    for ci, cell_text in enumerate(row):
                        cell = tbl.cell(ri, ci)
                        cell.text = ""
                        p = cell.paragraphs[0]
                        if ri == 0:
                            p.add_run(cell_text).bold = True
                        else:
                            _add_inline(p, cell_text)
            continue  # i already advanced past the table
        elif line.startswith("- ") or line.startswith("* "):
            p = doc.add_paragraph(style="List Bullet")
            _add_inline(p, line[2:].strip())
        elif re.match(r"^\d+\. ", line):
            p = doc.add_paragraph(style="List Number")
            _add_inline(p, re.sub(r"^\d+\. ", "", line).strip())
        elif line.strip():
            p = doc.add_paragraph()
            _add_inline(p, line.strip())

        i += 1

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

# ── File upload ───────────────────────────────────────────────
uploaded = st.file_uploader(
    "Upload a video or audio file",
    type=["mkv","mp4","avi","mov","webm","mp3","m4a","wav","ogg","flac","aac"],
)

if not uploaded:
    st.info("Upload a file above to get started.")
    st.stop()

if not any([do_summary, do_mom, do_email]):
    st.warning("Please select at least one output in the sidebar.")
    st.stop()

if not st.button("▶ Process", type="primary"):
    st.stop()

# ── Pipeline ──────────────────────────────────────────────────
with tempfile.TemporaryDirectory() as tmp:
    tmp        = Path(tmp)
    input_path = tmp / uploaded.name
    audio_path = tmp / "audio.mp3"
    stem       = Path(uploaded.name).stem

    input_path.write_bytes(uploaded.read())

    # Step 1 — Extract / convert audio (MP3 32kbps mono keeps file well under Groq's 25 MB limit)
    with st.status("Step 1 / 3 — Extracting audio …") as status:
        result = subprocess.run(
            [
                "ffmpeg", "-y", "-i", str(input_path),
                "-vn", "-acodec", "libmp3lame",
                "-ar", "16000", "-ac", "1", "-b:a", "32k",
                str(audio_path),
            ],
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            st.error(f"FFmpeg failed:\n{result.stderr}")
            st.stop()
        size_mb = audio_path.stat().st_size / 1_000_000
        if size_mb > 24:
            st.error(f"Audio file is {size_mb:.1f} MB after compression — exceeds Groq's 25 MB limit. Please upload a shorter clip.")
            st.stop()
        status.update(label="Step 1 / 3 — Audio ready", state="complete")

    # Step 2 — Transcribe via Groq Whisper
    with st.status("Step 2 / 3 — Transcribing …") as status:
        groq_client = Groq(api_key=GROQ_API_KEY)
        with open(audio_path, "rb") as f:
            response = groq_client.audio.transcriptions.create(
                file=(audio_path.name, f.read()),
                model="whisper-large-v3",
                language=None if language == "auto" else language,
                response_format="text",
            )
        transcript = response
        status.update(label="Step 2 / 3 — Transcription complete", state="complete")

    # Step 3 — Generate documents via Claude
    claude = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    def claude_call(system_prompt: str, user_prompt: str) -> str:
        msg = claude.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2000,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )
        return msg.content[0].text

    base_system = (
        f"You are an expert technical writer for a medical diagnostics company.\n"
        f"Company: {company}\n"
        f"Author: {sender_name}, {sender_title}\n"
        f"Context: {feedback_context}"
    )

    summary = mom = email = None

    if do_summary:
        with st.status("Step 3 / 3 — Generating summary …") as status:
            summary = claude_call(
                system_prompt=base_system,
                user_prompt=f"""Write a concise executive summary (5–7 bullet points) of the transcript below.
Focus on key observations, issues identified, and outcomes.

Transcript:
\"\"\"{transcript}\"\"\"
""",
            )
            status.update(label="Step 3 / 3 — Summary done", state="complete")

    if do_mom:
        with st.status("Step 3 / 3 — Generating minutes of meeting …") as status:
            mom = claude_call(
                system_prompt=base_system,
                user_prompt=f"""Generate structured Minutes of Meeting from the transcript below.

Use this format exactly:

## Minutes of Meeting
**Date:** [infer from transcript or leave blank]
**Prepared by:** {sender_name}, {sender_title}, {company}
**Attendees:** [infer from transcript or leave blank]

### Agenda
[infer from transcript]

### Discussion
[key points discussed]

### Observations & Findings
[technical observations and issues noted]

### Decisions
[decisions made, if any]

### Action Items
| # | Action | Owner | Due |
|---|--------|-------|-----|
[rows]

Transcript:
\"\"\"{transcript}\"\"\"
""",
            )
            status.update(label="Step 3 / 3 — Minutes done", state="complete")

    if do_email:
        with st.status("Step 3 / 3 — Generating feedback email …") as status:
            email = claude_call(
                system_prompt=base_system,
                user_prompt=f"""Convert the transcript below into a professional feedback email.

Subject: {email_subject}
From: {sender_name}, {sender_title}, {company}
To: {recipient_name}

Guidelines:
- Preserve ALL technical observations and findings.
- Organise: observations → findings → suggested actions.
- Use bullet points where appropriate.
- Do NOT add information absent from the transcript.

Transcript:
\"\"\"{transcript}\"\"\"
""",
            )
            status.update(label="Step 3 / 3 — Email done", state="complete")

# ── Outputs ───────────────────────────────────────────────────
st.success("Done!")

tab_labels = []
if do_summary: tab_labels.append("📋 Summary")
if do_mom:     tab_labels.append("📝 Minutes of Meeting")
if do_email:   tab_labels.append("📧 Feedback Email")
tab_labels.append("🗒️ Transcript")

tabs = st.tabs(tab_labels)
idx = 0

if do_summary:
    with tabs[idx]:
        st.markdown(summary)
        col1, col2 = st.columns(2)
        col1.download_button("⬇ Download Summary (.md)", summary,
                             f"{stem}_summary.md", mime="text/markdown")
        col2.download_button("⬇ Download Summary (.docx)",
                             markdown_to_docx(summary, "Summary"),
                             f"{stem}_summary.docx", mime=DOCX_MIME)
    idx += 1

if do_mom:
    with tabs[idx]:
        st.markdown(mom)
        col1, col2 = st.columns(2)
        col1.download_button("⬇ Download MoM (.md)", mom,
                             f"{stem}_mom.md", mime="text/markdown")
        col2.download_button("⬇ Download MoM (.docx)",
                             markdown_to_docx(mom, "Minutes of Meeting"),
                             f"{stem}_mom.docx", mime=DOCX_MIME)
    idx += 1

if do_email:
    with tabs[idx]:
        st.markdown(email)
        col1, col2 = st.columns(2)
        col1.download_button("⬇ Download Email (.md)", email,
                             f"{stem}_feedback_email.md", mime="text/markdown")
        col2.download_button("⬇ Download Email (.docx)",
                             markdown_to_docx(email, "Feedback Email"),
                             f"{stem}_feedback_email.docx", mime=DOCX_MIME)
    idx += 1

with tabs[idx]:
    st.text_area("Raw transcript", transcript, height=300)
    st.download_button("⬇ Download Transcript", transcript,
                       f"{stem}_transcript.txt", mime="text/plain")
