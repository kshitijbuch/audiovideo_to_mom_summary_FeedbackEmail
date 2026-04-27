"""
app.py
Audio / Video → Executive Summary + Minutes of Meeting + Feedback Email
Transasia Biomedicals Ltd.
"""

import os
import tempfile
import subprocess
from pathlib import Path

import streamlit as st
import anthropic
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
    st.subheader("Document metadata")
    sender_name      = st.text_input("Your name",    "Kshitij Buch")
    sender_title     = st.text_input("Your title",   "Digital Projects & Technical Support Manager")
    company          = st.text_input("Company",      "Transasia Biomedicals Ltd.")
    recipient_name   = st.text_input("Recipient",    "Kshitij Buch")
    email_subject    = st.text_input("Email subject","Feedback – XL200 Troubleshooting Agent CRU Error")
    feedback_context = st.text_input("Context",      "XL200 Troubleshooting Agent Response Feedback")

# ── File upload ───────────────────────────────────────────────
uploaded = st.file_uploader(
    "Upload a video or audio file",
    type=["mkv","mp4","avi","mov","webm","mp3","m4a","wav","ogg","flac","aac"],
)

if not uploaded:
    st.info("Upload a file above to get started.")
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
st.success("All documents generated!")

tab1, tab2, tab3, tab4 = st.tabs(
    ["📋 Summary", "📝 Minutes of Meeting", "📧 Feedback Email", "🗒️ Transcript"]
)

with tab1:
    st.markdown(summary)
    st.download_button("⬇ Download Summary", summary,
                       f"{stem}_summary.md", mime="text/markdown")

with tab2:
    st.markdown(mom)
    st.download_button("⬇ Download MoM", mom,
                       f"{stem}_mom.md", mime="text/markdown")

with tab3:
    st.markdown(email)
    st.download_button("⬇ Download Email", email,
                       f"{stem}_feedback_email.md", mime="text/markdown")

with tab4:
    st.text_area("Raw transcript", transcript, height=300)
    st.download_button("⬇ Download Transcript", transcript,
                       f"{stem}_transcript.txt", mime="text/plain")
