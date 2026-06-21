# Claude_Context.md — Handoff File

> Paste this whole file into a new Claude session to resume with zero information loss.

## Project

**Name:** Audio/Video to Summary / MOM / Feedback Email
**Local path (Windows):** `C:\Users\Kshitij Buch\OneDrive\Documents\TBM 2026 Onwards\Projects\Project2 Transcription to Text\Project2 AudioVideo_to_Summ_MOM_FbEmail`
**GitHub repo:** `git@github.com:kshitijbuch/audiovideo_to_mom_summary_FeedbackEmail.git`
**Deployment:** Live on Streamlit Community Cloud (share.streamlit.io), auto-deploys from `master` on push.
**Python interpreter (local):** `C:\anaconda3\python.exe`

### What the app does
Upload an audio/video file, extract + chunk audio (ffmpeg), transcribe via Groq Whisper API, generate Summary / Minutes of Meeting / Feedback Email via Claude API, download as `.md` or `.docx`.

### Two scripts in the repo
- **`app.py`** — Streamlit web app, the actively maintained one. All recent work was here.
- **`video_to_email.py`** — older CLI-only predecessor (faster-whisper, hardcoded filename/config, single output type). Still present but not modified this session; flagged earlier as a candidate for removal/deprecation, no decision made yet.

---

## What we did this session

1. Explored the repo and identified `app.py` as the active app, `video_to_email.py` as a legacy CLI tool.
2. Rewrote all three Claude prompts (Summary, Minutes of Meeting, Feedback Email) inside `app.py` to enforce strict verbatim/accurate capture: no fabrication, no padding, no inference. Explicit, repeated user requirement: whatever is said in the audio transcription must be the same thing captured in the Summary and MOM — nothing not said should be made up to fill space. The output should serve the purpose, not please the user.
3. Added Anthropic prompt caching (`anthropic-beta: prompt-caching-2024-07-31`). System prompt + transcript are cached as ephemeral blocks, shared across the 3 Claude calls (Summary/MOM/Email), cutting input-token cost roughly 60–70% on calls 2 and 3.
4. Fixed `markdown_to_docx()` to support nested bullets/numbered lists (2-space indent maps to "List Bullet 2", 4-space to "List Bullet 3"), since MOM output often has sub-items that previously rendered flat in Word.
5. Moved hardcoded metadata (sender name/title, company, recipient, email subject, context) out of code into env/secrets-driven defaults. First attempt used the user's own details as defaults; corrected per explicit user feedback to generic placeholders in code ("Your Name", "Your Title", "Your Company", "Recipient Name", "Meeting / Session Feedback", "Brief context of the recording") so the public app works for any user out of the box. Personal values are opt-in only, via Streamlit Cloud Secrets or local `.env` — never hardcoded in source.
6. Unified API-key/secrets loading: checks `st.secrets` first (Streamlit Cloud), falls back to `os.environ` (local `.env`). Same pattern used for both API keys and metadata defaults via a small `_secret()` helper function.
7. Improved rate-limit UX: Groq 429 retry now shows a live per-second countdown in the `st.status` widget instead of a silent `time.sleep(65)`.
8. Added `.env.example` documenting all required (`ANTHROPIC_API_KEY`, `GROQ_API_KEY`) and optional (metadata) keys.
9. Committed and pushed 2 commits this session (see Git status below). Streamlit Cloud auto-redeployed after each push.

### Key decision/correction mid-session
First attempt at metadata defaults pulled the user's own details (Kshitij Buch / Transasia Biomedicals Ltd.) into the `.env`/fallback defaults. User corrected this: the app must ship with generic, anonymous placeholders so any user can drop in their own details — personal info must not be baked into the shared/public codebase. Fixed in commit `17d6978`.

---

## File paths touched this session

- `app.py` — main Streamlit app (all prompt/caching/UX/metadata changes)
- `.env` — local secrets file (gitignored, NOT pushed). Currently contains the user's real API keys plus their own metadata values — fine locally since it is git-ignored.
- `.env.example` — new file, pushed to GitHub, generic placeholder template for anyone cloning the repo.
- `video_to_email.py` — read/reviewed only, not modified.
- `.gitignore` — reviewed only, not modified (already correctly ignores `.env`, `input/`, `output/`, `transcript/`).
- `requirements.txt` — reviewed only, not modified.

---

## Git / commit status

**Branch:** `master`, up to date with `origin/master`. Working tree clean (no uncommitted changes) as of end of session.

Commits made this session (newest first):
- `17d6978` Use generic sidebar defaults; read metadata from st.secrets on Streamlit Cloud
- `10dcaf1` Tighten prompts for verbatim capture; add prompt caching and UX improvements

(`3b2e448` and earlier predate this session.)

All pushed to `origin/master` successfully. No open branches, no PRs, no merge conflicts.

---

## Critical code reference (current state of app.py)

### Secrets/metadata loading pattern (top of file)

    ANTHROPIC_API_KEY = (
        st.secrets.get("ANTHROPIC_API_KEY")
        or os.environ.get("ANTHROPIC_API_KEY", "")
    )
    GROQ_API_KEY = (
        st.secrets.get("GROQ_API_KEY")
        or os.environ.get("GROQ_API_KEY", "")
    )

    def _secret(key: str, fallback: str = "") -> str:
        return st.secrets.get(key) or os.environ.get(key, fallback)

    _DEF_SENDER_NAME  = _secret("SENDER_NAME",       "Your Name")
    _DEF_SENDER_TITLE = _secret("SENDER_TITLE",      "Your Title")
    _DEF_COMPANY      = _secret("COMPANY",           "Your Company")
    _DEF_RECIPIENT    = _secret("RECIPIENT_NAME",    "Recipient Name")
    _DEF_SUBJECT      = _secret("EMAIL_SUBJECT",     "Meeting / Session Feedback")
    _DEF_CONTEXT      = _secret("FEEDBACK_CONTEXT",  "Brief context of the recording")

### Claude call pattern with prompt caching

    def claude_call(task_prompt: str) -> str:
        msg = claude.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4000,
            system=[
                {
                    "type": "text",
                    "text": base_system,
                    "cache_control": {"type": "ephemeral"},
                }
            ],
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": f"TRANSCRIPT:\n\"\"\"\n{transcript}\n\"\"\"",
                            "cache_control": {"type": "ephemeral"},
                        },
                        {"type": "text", "text": task_prompt},
                    ],
                }
            ],
            extra_headers={"anthropic-beta": "prompt-caching-2024-07-31"},
        )
        return msg.content[0].text

### Core prompt rule baked into base_system (enforces verbatim capture)

    ABSOLUTE RULES — apply to every output you produce:
    1. Capture ONLY what was explicitly said in the transcript. Never infer, extrapolate, or add anything not directly stated.
    2. Use the speaker's own words and phrasing as closely as possible. Do not paraphrase to make it sound better.
    3. Do not add pleasantries, filler phrases, or padding sentences to meet a length target.
    4. If a section has no content in the transcript, write exactly: [Not mentioned] — never leave it blank or fabricate content.
    5. Do not validate, endorse, or editorialize the speaker's observations. Record them as stated.
    6. Omit any section entirely if the transcript contains zero content for it.

---

## Exact next steps (numbered, in priority order)

1. Verify Streamlit Cloud redeploy — open the live app URL, hard-refresh, confirm sidebar now shows generic placeholders ("Your Name", "Your Title", etc.) and not the previous hardcoded Transasia values.
2. Optionally set personal metadata in Streamlit Cloud Secrets for the user's own deployment only — App Settings → Secrets, add `SENDER_NAME`, `SENDER_TITLE`, `COMPANY`, `RECIPIENT_NAME`, `EMAIL_SUBJECT`, `FEEDBACK_CONTEXT` with personal values, so the user's own deployment pre-fills while the public repo stays generic.
3. Decide fate of `video_to_email.py` — either remove it (superseded by `app.py`) or convert it into a thin CLI wrapper sharing logic with `app.py`. Not yet decided.
4. Test prompt-caching savings empirically — run one real file through the app and check the Anthropic usage/cost dashboard to confirm cache-hit behavior on the MOM/Email calls (cache only kicks in if the transcript is roughly 1024+ tokens for Sonnet).
5. Consider single-call generation (not yet implemented) — combining Summary+MOM+Email into one Claude call was discussed as an alternative to the current 3-separate-calls-with-caching approach, but rejected in favor of caching since it is simpler to parse. Revisit only if cost/latency becomes an issue.
6. Review the markdown_to_docx nested-bullet fix on a real MOM output — fix was made but not yet visually verified against a live .docx download in Word.

---

## Blockers / open issues

- None currently blocking. All code is committed, pushed, and presumed auto-deployed.
- Unverified: whether Streamlit Cloud has actually picked up the latest push and finished redeploying — user should confirm visually before considering this fully closed.
- Local `.env` still contains the user's personal info and real API keys — fine since it is gitignored, but worth remembering it is intentionally different from the generic `.env.example` and the in-code fallbacks.

---

## Earlier session context (improvements suggested but not all acted on)

From the initial exploration, these were flagged but not yet implemented:
- Nested bullets/blockquotes edge cases beyond what was fixed (blockquotes using `>` are not handled at all).
- `input/` folder had sample test media files (.mkv, .mp4, .mp3, .docx) — confirmed already gitignored, no action needed.
- No automated tests exist for `app.py` or `video_to_email.py`.
