"""
app.py
Audio / Video → Executive Summary + Minutes of Meeting + Feedback Email
Transasia Biomedicals Ltd.
"""

import os
import re
import time
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

# ── API keys ──────────────────────────────────────────────────
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
    chunk_minutes = st.slider(
        "Chunk size (minutes)", min_value=10, max_value=30, value=20, step=5,
        help="Audio is split into chunks of this size. Smaller = safer for rate limits.",
    )
    st.divider()
    st.subheader("Generate")
    do_summary = st.checkbox("📋 Summary",           value=True)
    do_mom     = st.checkbox("📝 Minutes of Meeting", value=True)
    do_email   = st.checkbox("📧 Feedback Email",    value=True)
    if not any([do_summary, do_mom, do_email]):
        st.warning("Select at least one output.")
    st.divider()
    st.subheader("Document metadata")
    sender_name      = st.text_input("Your name",     "Kshitij Buch")
    sender_title     = st.text_input("Your title",    "Digital Projects & Technical Support Manager")
    company          = st.text_input("Company",       "Transasia Biomedicals Ltd.")
    recipient_name   = st.text_input("Recipient",     "Kshitij Buch")
    email_subject    = st.text_input("Email subject", "Feedback – XL200 Troubleshooting Agent CRU Error")
    feedback_context = st.text_input("Context",       "XL200 Troubleshooting Agent Response Feedback")

# ── Helpers ───────────────────────────────────────────────────
def fmt_time(seconds: float) -> str:
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


def _add_inline(para, text: str):
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
            tbl_lines = []
            while i < len(lines) and lines[i].startswith("|"):
                tbl_lines.append(lines[i])
                i += 1
            rows = [
                [c.strip() for c in row.strip().strip("|").split("|")]
                for row in tbl_lines
            ]
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
            continue
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

# ── Session state ─────────────────────────────────────────────
_defaults = {
    "stage":      "upload",
    "chunks":     None,
    "stem":       None,
    "file_id":    None,
    "transcript": None,
    "summary":    None,
    "mom":        None,
    "email":      None,
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── File upload ───────────────────────────────────────────────
uploaded = st.file_uploader(
    "Upload a video or audio file",
    type=["mkv","mp4","avi","mov","webm","mp3","m4a","wav","ogg","flac","aac"],
)

if not uploaded:
    st.session_state.stage  = "upload"
    st.session_state.chunks = None
    st.info("Upload a file above to get started.")
    st.stop()

# Reset if a different file is uploaded
file_id = f"{uploaded.name}_{uploaded.size}"
if st.session_state.file_id != file_id:
    for k, v in _defaults.items():
        st.session_state[k] = v
    st.session_state.file_id = file_id

if not any([do_summary, do_mom, do_email]):
    st.warning("Please select at least one output in the sidebar.")
    st.stop()

# ═══════════════════════════════════════════════════════════════
# STAGE 1 — Extract & split audio
# ═══════════════════════════════════════════════════════════════
if st.session_state.stage == "upload":
    if st.button("▶ Extract & Split Audio", type="primary"):
        with tempfile.TemporaryDirectory() as tmp:
            tmp        = Path(tmp)
            input_path = tmp / uploaded.name
            audio_path = tmp / "audio.mp3"
            input_path.write_bytes(uploaded.read())

            # Extract / convert to MP3 32kbps mono
            with st.status("Step 1 — Extracting audio …") as s:
                res = subprocess.run(
                    ["ffmpeg", "-y", "-i", str(input_path),
                     "-vn", "-acodec", "libmp3lame",
                     "-ar", "16000", "-ac", "1", "-b:a", "32k",
                     str(audio_path)],
                    capture_output=True, text=True,
                )
                if res.returncode != 0:
                    st.error(f"FFmpeg failed:\n{res.stderr}")
                    st.stop()
                s.update(label="Step 1 — Audio extracted", state="complete")

            # Get total duration via ffprobe
            probe = subprocess.run(
                ["ffprobe", "-v", "error",
                 "-show_entries", "format=duration",
                 "-of", "default=noprint_wrappers=1:nokey=1",
                 str(audio_path)],
                capture_output=True, text=True,
            )
            total_dur = float(probe.stdout.strip())

            # Split into chunks
            chunk_sec = chunk_minutes * 60
            with st.status(f"Step 2 — Splitting into {chunk_minutes}-min chunks …") as s:
                subprocess.run(
                    ["ffmpeg", "-y", "-i", str(audio_path),
                     "-f", "segment", "-segment_time", str(chunk_sec),
                     "-c", "copy", str(tmp / "chunk_%03d.mp3")],
                    capture_output=True,
                )
                chunk_files = sorted(tmp.glob("chunk_*.mp3"))
                chunks = []
                for idx, cf in enumerate(chunk_files):
                    start   = idx * chunk_sec
                    end     = min((idx + 1) * chunk_sec, total_dur)
                    size_mb = cf.stat().st_size / 1_000_000
                    chunks.append({
                        "idx":      idx,
                        "start":    start,
                        "end":      end,
                        "duration": end - start,
                        "label":    (
                            f"Chunk {idx+1}:  {fmt_time(start)} – {fmt_time(end)}"
                            f"  ({(end-start)/60:.1f} min,  {size_mb:.1f} MB)"
                        ),
                        "bytes":    cf.read_bytes(),
                        "name":     cf.name,
                    })
                st.session_state.chunks = chunks
                st.session_state.stem   = Path(uploaded.name).stem
                s.update(label=f"Step 2 — Split into {len(chunks)} chunk(s)", state="complete")

        st.session_state.stage = "select"
        st.rerun()

# ═══════════════════════════════════════════════════════════════
# STAGE 2 — Select chunks to transcribe
# ═══════════════════════════════════════════════════════════════
elif st.session_state.stage == "select":
    chunks = st.session_state.chunks

    st.subheader(f"Select chunks to transcribe  ({len(chunks)} total)")

    # Initialise checkbox defaults (all selected)
    for c in chunks:
        if f"chunk_{c['idx']}" not in st.session_state:
            st.session_state[f"chunk_{c['idx']}"] = True

    col1, col2 = st.columns(2)
    if col1.button("☑ Select All"):
        for c in chunks:
            st.session_state[f"chunk_{c['idx']}"] = True
        st.rerun()
    if col2.button("☐ Deselect All"):
        for c in chunks:
            st.session_state[f"chunk_{c['idx']}"] = False
        st.rerun()

    st.divider()
    for chunk in chunks:
        st.checkbox(chunk["label"], key=f"chunk_{chunk['idx']}")

    selected     = [c for c in chunks if st.session_state.get(f"chunk_{c['idx']}", True)]
    total_sel_s  = sum(c["duration"] for c in selected)

    st.divider()
    if not selected:
        st.warning("Select at least one chunk.")
        st.stop()

    # Advisory based on Groq free-tier hourly limit (7,200 audio seconds/hour)
    GROQ_HOURLY_LIMIT = 7200
    if total_sel_s > GROQ_HOURLY_LIMIT:
        hrs = total_sel_s / GROQ_HOURLY_LIMIT
        st.warning(
            f"**Selected: {total_sel_s/60:.0f} min of audio.**  "
            f"Groq free tier allows 120 min/hour — this will span "
            f"~{hrs:.1f} hour-windows. The app will pause automatically "
            f"if a rate-limit is hit and retry after 65 seconds."
        )
    else:
        st.info(f"Selected: {total_sel_s/60:.0f} min — fits within Groq's hourly limit.")

    if st.button("▶ Transcribe & Generate", type="primary"):

        # ── Transcribe each selected chunk ────────────────────
        groq_client = Groq(api_key=GROQ_API_KEY)
        parts = []

        for i, chunk in enumerate(selected):
            short = f"Chunk {chunk['idx']+1} ({fmt_time(chunk['start'])}–{fmt_time(chunk['end'])})"
            with st.status(f"Transcribing {short}  [{i+1}/{len(selected)}] …") as s:
                tf = tempfile.NamedTemporaryFile(suffix=".mp3", delete=False)
                tf.write(chunk["bytes"])
                tf.close()
                try:
                    for attempt in range(3):
                        try:
                            with open(tf.name, "rb") as f:
                                resp = groq_client.audio.transcriptions.create(
                                    file=(chunk["name"], f.read()),
                                    model="whisper-large-v3",
                                    language=None if language == "auto" else language,
                                    response_format="text",
                                )
                            parts.append(f"[{short}]\n{resp}")
                            s.update(label=f"{short} — done ✓", state="complete")
                            break
                        except Exception as e:
                            if attempt < 2 and "rate" in str(e).lower():
                                st.toast("Rate limit hit — waiting 65 s before retry …")
                                time.sleep(65)
                            else:
                                st.error(f"Transcription failed: {e}")
                                st.stop()
                finally:
                    os.unlink(tf.name)

            # Small courtesy pause between chunks
            if i < len(selected) - 1:
                time.sleep(3)

        st.session_state.transcript = "\n\n".join(parts)

        # ── Generate documents via Claude ─────────────────────
        claude     = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
        transcript = st.session_state.transcript

        def claude_call(system_prompt: str, user_prompt: str) -> str:
            msg = claude.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=4000,
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

        if do_summary:
            with st.status("Generating summary …") as s:
                st.session_state.summary = claude_call(
                    base_system,
                    f"""Write a concise executive summary (5–7 bullet points) of the transcript below.
Focus on key observations, issues identified, and outcomes.

Transcript:
\"\"\"{transcript}\"\"\"
""",
                )
                s.update(label="Summary done", state="complete")

        if do_mom:
            with st.status("Generating minutes of meeting …") as s:
                st.session_state.mom = claude_call(
                    base_system,
                    f"""Generate structured Minutes of Meeting from the transcript below.

CRITICAL RULES — follow these strictly:
1. Only capture what was actually said in the transcript. Do NOT add, infer, embellish, or assume anything not explicitly stated.
2. Use the speaker's own words and sentences as closely as possible. Avoid paraphrasing or rewriting.
3. If information for any section is absent from the transcript, write exactly: "[Not mentioned/discussed in this meeting]" — never leave a section blank or fabricate content.
4. AGENDA: Extract ONLY what the meeting host/initiator stated as the purpose of the meeting — verbatim or near-verbatim. Do NOT include responses, additions, or elaborations made by other participants. The agenda is set by the host alone.
5. DISCUSSION: Capture what each participant said using their actual words as closely as possible. Attribute statements to the correct speaker.
6. OBSERVATIONS & FINDINGS: Only include observations explicitly stated in the transcript. No inferences.
7. DECISIONS: Only include decisions explicitly stated or agreed upon. Do not infer decisions from discussion.
8. ACTION ITEMS: Only include tasks explicitly assigned or volunteered in the transcript. Do not derive action items from discussion.

Use this format exactly:

## Minutes of Meeting
**Date:** [exact date if stated in transcript; otherwise write: Not mentioned in the meeting]
**Prepared by:** {sender_name}, {sender_title}, {company}
**Attendees:** [names if mentioned in transcript; otherwise write: Not mentioned in the meeting]

### Agenda
[Only the host/initiator's stated purpose — verbatim or near-verbatim. No other participant's input here.]

### Discussion
[Key points by all participants using their actual words as closely as possible, attributed to the correct speaker]

### Observations & Findings
[Only observations explicitly stated in the transcript. If none: Not mentioned in the meeting.]

### Decisions
[Only decisions explicitly stated or agreed upon. If none: Not mentioned in the meeting.]

### Action Items
| # | Action | Owner | Due |
|---|--------|-------|-----|
[Only tasks explicitly assigned or agreed upon in the transcript. If none, add a single row: Not discussed in the meeting | — | —]

Transcript:
\"\"\"{transcript}\"\"\"
""",
                )
                s.update(label="Minutes done", state="complete")

        if do_email:
            with st.status("Generating feedback email …") as s:
                st.session_state.email = claude_call(
                    base_system,
                    f"""Convert the transcript below into a professional feedback email.

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
                s.update(label="Email done", state="complete")

        st.session_state.stage = "done"
        st.rerun()

# ═══════════════════════════════════════════════════════════════
# STAGE 3 — Outputs
# ═══════════════════════════════════════════════════════════════
elif st.session_state.stage == "done":
    st.success("Done!")
    stem       = st.session_state.stem
    summary    = st.session_state.summary
    mom        = st.session_state.mom
    email      = st.session_state.email
    transcript = st.session_state.transcript

    tab_labels = []
    if do_summary and summary: tab_labels.append("📋 Summary")
    if do_mom     and mom:     tab_labels.append("📝 Minutes of Meeting")
    if do_email   and email:   tab_labels.append("📧 Feedback Email")
    tab_labels.append("🗒️ Transcript")

    tabs = st.tabs(tab_labels)
    idx  = 0

    if do_summary and summary:
        with tabs[idx]:
            st.markdown(summary)
            c1, c2 = st.columns(2)
            c1.download_button("⬇ Download Summary (.md)", summary,
                               f"{stem}_summary.md", mime="text/markdown")
            c2.download_button("⬇ Download Summary (.docx)",
                               markdown_to_docx(summary, "Summary"),
                               f"{stem}_summary.docx", mime=DOCX_MIME)
        idx += 1

    if do_mom and mom:
        with tabs[idx]:
            st.markdown(mom)
            c1, c2 = st.columns(2)
            c1.download_button("⬇ Download MoM (.md)", mom,
                               f"{stem}_mom.md", mime="text/markdown")
            c2.download_button("⬇ Download MoM (.docx)",
                               markdown_to_docx(mom, "Minutes of Meeting"),
                               f"{stem}_mom.docx", mime=DOCX_MIME)
        idx += 1

    if do_email and email:
        with tabs[idx]:
            st.markdown(email)
            c1, c2 = st.columns(2)
            c1.download_button("⬇ Download Email (.md)", email,
                               f"{stem}_feedback_email.md", mime="text/markdown")
            c2.download_button("⬇ Download Email (.docx)",
                               markdown_to_docx(email, "Feedback Email"),
                               f"{stem}_feedback_email.docx", mime=DOCX_MIME)
        idx += 1

    with tabs[idx]:
        st.text_area("Raw transcript", transcript, height=300)
        st.download_button("⬇ Download Transcript", transcript,
                           f"{stem}_transcript.txt", mime="text/plain")

    st.divider()
    if st.button("↩ Process Another File"):
        for k, v in _defaults.items():
            st.session_state[k] = v
        st.rerun()
