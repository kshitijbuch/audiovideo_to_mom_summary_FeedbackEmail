"""
video_to_email.py
-----------------
XL200 / Transasia Biomedicals – Video/Audio-to-Email Feedback Tool
Converts a video or audio file → transcription → polished feedback email.

Supported input formats:
    Video : .mkv  .mp4  .avi  .mov  .webm
    Audio : .mp3  .m4a  .wav  .ogg  .flac  .aac

Folder structure:
    input/       ← place video/audio files here
    transcript/  ← _transcript.txt files saved here
    output/      ← _feedback_email.md files saved here
    .env         ← ANTHROPIC_API_KEY

Requirements:
    conda install -c conda-forge ffmpeg -y
    pip install faster-whisper anthropic python-dotenv

Usage:
    python video_to_email.py
"""

import os
import sys
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"  # workaround for duplicate OpenMP runtime
sys.stdout.reconfigure(encoding="utf-8")
import subprocess
from pathlib import Path
import anthropic
from faster_whisper import WhisperModel
from dotenv import load_dotenv

load_dotenv()

# ─────────────────────────────────────────────
# PATHS — derived from script location
# ─────────────────────────────────────────────
BASE_DIR        = Path(__file__).parent
INPUT_DIR       = BASE_DIR / "input"
TRANSCRIPT_DIR  = BASE_DIR / "transcript"
OUTPUT_DIR      = BASE_DIR / "output"

for _dir in [INPUT_DIR, TRANSCRIPT_DIR, OUTPUT_DIR]:
    _dir.mkdir(exist_ok=True)

# ─────────────────────────────────────────────
# CONFIGURATION — Edit these before running
# ─────────────────────────────────────────────
INPUT_FILENAME   = "CRU Error Feedback2.mkv"        # ← filename inside input/ (video or audio)
INPUT_PATH       = INPUT_DIR / INPUT_FILENAME
AUDIO_OUTPUT     = BASE_DIR / "audio_temp.wav"       # temp file, auto-deleted
WHISPER_MODEL    = "base"                            # tiny / base / small / medium / large
LANGUAGE         = "en"                              # "en" / "hi" / None (auto-detect)
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]  # loaded from .env

# Email metadata
SENDER_NAME      = "Kshitij Buch"
SENDER_TITLE     = "Digital Projects & Technical Support Manager"
COMPANY          = "Transasia Biomedicals Ltd."
RECIPIENT_NAME   = "Kshitij Buch"
EMAIL_SUBJECT    = "Feedback – XL200 Troubleshooting Agent CRU Error"
FEEDBACK_CONTEXT = "XL200 Troubleshooting Agent Response Feedback"


# ─────────────────────────────────────────────
# STEP 1: Extract Audio from Video using FFmpeg
# ─────────────────────────────────────────────
def extract_audio(video_path: Path, audio_path: Path) -> None:
    print(f"\n[1/3] Extracting audio from: {video_path}")
    cmd = [
        "ffmpeg", "-y",
        "-i", str(video_path),
        "-vn",                   # no video
        "-acodec", "pcm_s16le",  # WAV format
        "-ar", "16000",          # 16kHz – optimal for Whisper
        "-ac", "1",              # mono
        str(audio_path)
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"FFmpeg error:\n{result.stderr}")
    print(f"    ✓ Audio saved to: {audio_path}")


# ─────────────────────────────────────────────
# STEP 2: Transcribe Audio using Faster-Whisper
# ─────────────────────────────────────────────
def transcribe_audio(audio_path: Path, model_size: str = "base", language: str = "en") -> str:
    print(f"\n[2/3] Transcribing audio (model: {model_size}) ...")
    print("    This may take 1–3 minutes depending on video length and your hardware.")

    model = WhisperModel(model_size, device="cpu", compute_type="int8")
    segments, info = model.transcribe(
        str(audio_path),
        language=language if language else None,
        beam_size=5,
        vad_filter=True,
        vad_parameters=dict(min_silence_duration_ms=500)
    )

    print(f"    ✓ Detected language: {info.language} (confidence: {info.language_probability:.0%})")

    transcript_lines = []
    for segment in segments:
        line = f"[{segment.start:.1f}s → {segment.end:.1f}s]  {segment.text.strip()}"
        print(f"    {line}")
        transcript_lines.append(segment.text.strip())

    full_transcript = " ".join(transcript_lines)
    print(f"\n    ✓ Transcription complete. Total words: {len(full_transcript.split())}")
    return full_transcript


# ─────────────────────────────────────────────
# STEP 3: Format Transcript into Feedback Email
# ─────────────────────────────────────────────
def format_as_email(
    transcript: str,
    api_key: str,
    sender_name: str,
    sender_title: str,
    company: str,
    recipient_name: str,
    subject: str,
    context: str,
) -> str:
    print(f"\n[3/3] Formatting transcript into feedback email via Claude ...")

    client = anthropic.Anthropic(api_key=api_key)

    system_prompt = f"""You are an expert technical writer for a medical diagnostics company.
Your task is to convert a raw transcription of a video/audio recording into a professional,
structured feedback email. The email should be clear, actionable, and professional in tone.

Guidelines:
- Preserve ALL technical observations, issues, and findings from the transcript.
- Organize content logically: observations → findings → suggested actions.
- Use bullet points for issues and recommendations where appropriate.
- Correct any transcription errors or awkward phrasing while preserving meaning.
- Do NOT fabricate or add information not present in the transcript.
- The sender is: {sender_name}, {sender_title}, {company}
- The recipient is: {recipient_name}
"""

    user_prompt = f"""Please convert the following raw video transcription into a professional feedback email.

Context: This is a {context}.

Raw Transcription:
\"\"\"{transcript}\"\"\"

Format the output as a complete email with:
- Subject line: {subject}
- Professional greeting
- Clear body with organized sections
- Action items or next steps (if applicable)
- Professional sign-off from {sender_name}
"""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1500,
        messages=[{"role": "user", "content": user_prompt}],
        system=system_prompt,
    )

    email_text = message.content[0].text
    print("    ✓ Email formatted successfully.\n")
    return email_text


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  VIDEO → TRANSCRIPT → FEEDBACK EMAIL")
    print("  Transasia Biomedicals Ltd.")
    print("=" * 60)

    stem = INPUT_PATH.stem  # e.g. "CRU Error Feedback2"
    transcript_file = TRANSCRIPT_DIR / f"{stem}_transcript.txt"
    email_file      = OUTPUT_DIR / f"{stem}_feedback_email.md"

    try:
        # Step 1 – Extract audio
        extract_audio(INPUT_PATH, AUDIO_OUTPUT)

        # Step 2 – Transcribe
        transcript = transcribe_audio(AUDIO_OUTPUT, WHISPER_MODEL, LANGUAGE)

        # Save raw transcript
        with open(transcript_file, "w", encoding="utf-8") as f:
            f.write(transcript)
        print(f"    ✓ Raw transcript saved to: {transcript_file}")

        # Step 3 – Format email
        email_content = format_as_email(
            transcript=transcript,
            api_key=ANTHROPIC_API_KEY,
            sender_name=SENDER_NAME,
            sender_title=SENDER_TITLE,
            company=COMPANY,
            recipient_name=RECIPIENT_NAME,
            subject=EMAIL_SUBJECT,
            context=FEEDBACK_CONTEXT,
        )

        # Save email
        with open(email_file, "w", encoding="utf-8") as f:
            f.write(email_content)

        print("=" * 60)
        print("GENERATED FEEDBACK EMAIL:")
        print("=" * 60)
        print(email_content)
        print("=" * 60)
        print(f"\n✓ Email saved to: {email_file}")
        print("  You can now copy-paste this into Outlook or Gmail.\n")

    finally:
        # Cleanup temp audio
        if AUDIO_OUTPUT.exists():
            AUDIO_OUTPUT.unlink()
            print(f"  ✓ Temp audio file cleaned up.")


if __name__ == "__main__":
    main()
