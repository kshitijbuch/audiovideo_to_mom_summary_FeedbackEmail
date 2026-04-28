# Audio / Video → Summary, MOM & Feedback Email

A Streamlit web app that takes any audio or video recording and produces three professional documents — an **Executive Summary**, **Minutes of Meeting (MoM)**, and a **Feedback Email** — all downloadable as Markdown or Word (.docx).

Built for Transasia Biomedicals Ltd.

---

## How It Works

```
Upload file → Extract audio (FFmpeg) → Transcribe (Groq Whisper) → Generate docs (Claude AI)
```

1. **Upload** a video or audio file (MKV, MP4, AVI, MOV, WebM, MP3, M4A, WAV, OGG, FLAC, AAC)
2. **FFmpeg** strips and converts the audio to MP3 32 kbps mono — small enough for Groq's 25 MB limit
3. Long recordings are **split into configurable chunks** (10–30 min); you can select which chunks to transcribe
4. **Groq Whisper large-v3** transcribes each chunk via API, with automatic rate-limit retry
5. **Anthropic Claude Sonnet** turns the transcript into Summary, MoM, and Feedback Email
6. Every document is downloadable as **.md** and **.docx** directly from the browser

---

## Tech Stack

| Layer | Technology |
|---|---|
| UI | Streamlit |
| Audio processing | FFmpeg (extraction + chunking) |
| Transcription | Groq API — `whisper-large-v3` |
| AI document generation | Anthropic API — `claude-sonnet-4-6` |
| Word export | `python-docx` |
| Secrets management | `python-dotenv` / Streamlit secrets |

---

## GitHub & CI/CD Pipeline

The repository is hosted at **[github.com/kshitijbuch/audiovideo_to_mom_summary_FeedbackEmail](https://github.com/kshitijbuch/audiovideo_to_mom_summary_FeedbackEmail)** and wired into the following GitHub-native pipeline:

- **GitHub Codespaces** — the `.devcontainer/devcontainer.json` config spins up a fully ready Python 3.11 environment in the cloud. On attach it automatically installs all dependencies and launches the Streamlit app on port 8501, so you can run and test the app without any local setup.
- **VS Code in the browser** — Codespaces opens `README.md` and `app.py` side-by-side on launch.
- **Dependency install on container build** — `packages.txt` (system packages via `apt`) and `requirements.txt` (Python packages via `pip`) are applied automatically.

To open the app in Codespaces: click **Code → Codespaces → Create codespace on master** in the GitHub UI.

---

## Local Setup

### Prerequisites

- Python 3.10+
- FFmpeg on your PATH (`conda install -c conda-forge ffmpeg -y` or via system package manager)

### Install

```bash
git clone git@github.com:kshitijbuch/audiovideo_to_mom_summary_FeedbackEmail.git
cd audiovideo_to_mom_summary_FeedbackEmail
pip install -r requirements.txt
```

### Configure API Keys

Create a `.env` file in the project root:

```
ANTHROPIC_API_KEY=sk-ant-...
GROQ_API_KEY=gsk_...
```

For Streamlit Cloud deployment, add these to `.streamlit/secrets.toml` instead.

### Run

```bash
streamlit run app.py
```

App opens at `http://localhost:8501`.

---

## Project Structure

```
app.py                  # Streamlit web app (main entry point)
video_to_email.py       # Legacy CLI script (local faster-whisper)
requirements.txt        # Python dependencies
packages.txt            # System-level packages for Codespaces
.devcontainer/          # GitHub Codespaces configuration
.streamlit/             # Streamlit config (theme, CORS settings)
input/                  # Place source video/audio files here (gitignored)
transcript/             # Raw transcript output (gitignored)
output/                 # Generated documents output (gitignored)
```

---

## Outputs

Each run produces up to three documents (selectable via sidebar checkboxes):

| Document | Description | Formats |
|---|---|---|
| Executive Summary | 5–7 bullet-point summary of key observations and outcomes | .md, .docx |
| Minutes of Meeting | Structured MoM with agenda, discussion, findings, decisions, and action items table | .md, .docx |
| Feedback Email | Professional email ready to paste into Outlook or Gmail | .md, .docx |

The raw transcript is also available to download as plain text.

---

## Configuration (Sidebar)

| Setting | Description |
|---|---|
| Transcription language | `en`, `hi`, or `auto` |
| Chunk size | Split audio into 10–30 min segments |
| Output selection | Toggle Summary, MoM, Email independently |
| Document metadata | Sender name/title, company, recipient, email subject, context |

---

## CLI Script (Legacy)

`video_to_email.py` is the original command-line version that uses **faster-whisper** running locally (no Groq API needed). Edit the configuration block at the top of the file, drop a file in `input/`, and run:

```bash
python video_to_email.py
```

This requires the `faster-whisper` package (`pip install faster-whisper`) in addition to the standard requirements.
