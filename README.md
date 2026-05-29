# job-apply — Corey's Job Application Agent

A Claude-powered web app (and CLI) that takes a job posting and produces a
tailored resume, ATS resume, and cover letter in under 2 minutes.

**Live app:** https://job-apply-corey.fly.dev/

---

## Features

- **Tailored resume** — styled DOCX with brand colors, targeted bullets, competency grid
- **ATS resume** — plain single-column DOCX, no tables or text boxes, parser-safe
- **Cover letter** — voice-matched DOCX tailored to the role and hiring manager
- **Google Drive sync** — all output files uploaded automatically to your Drive folder
- **Past Runs** — collapsible list of previous applications, refreshes in real-time after each run
- **Interview Prep** — generates a 9-section two-column reference card (role fit map, gap bridges, anchor stories, likely questions, differentiating edge, and more) tailored to the specific interviewer and round type
- **Auth** — account-based; each user gets their own resume, profile guide, and Drive silo
- **SSE progress streaming** — live log output while the agent runs

---

## Project Structure

```
job-apply/
├── CLAUDE.md              ← Agent workflow instructions (Claude Code reads this)
├── README.md              ← This file
├── api.py                 ← FastAPI backend (auth, run queue, SSE, Drive proxy)
├── apply.py               ← Core workflow engine + CLI entry point
├── profile.md             ← Corey's voice, stories, metrics, do-not-use phrases
├── frontend/
│   ├── index.html         ← Main SPA (form, progress log, results, past runs, prep)
│   ├── login.html
│   ├── register.html
│   └── profile.html
├── scripts/
│   ├── storage.py         ← Tigris S3 adapter (users, resumes, profiles)
│   └── office/            ← DOCX unpack / pack / validate
├── resumes/
│   └── master.docx        ← Source-of-truth resume (never use an output file)
├── jobs/
│   └── job.txt            ← Drop job postings here for CLI runs
├── output/                ← Generated files land here (gitignored)
├── Dockerfile
└── fly.toml
```

---

## Web App Usage

1. Go to https://job-apply-corey.fly.dev/
2. Register with your email, upload `master.docx`, and paste your `profile.md`
3. Paste a job posting, enter company + role, hit **Generate**
4. Watch the live log; download all three output files when done
5. Use **Past Runs** to revisit previous applications
6. Use **Interview Prep** to generate a reference card for an upcoming round — enter the interviewer name and round type for a fully tailored output

---

## CLI Usage

The CLI runs the same workflow locally without the web server.

```bash
# Install dependencies
pip install -r requirements.txt
npm install

# Set your API key
export ANTHROPIC_API_KEY=sk-ant-...

# Paste a job posting
pbpaste > jobs/job.txt    # or vim jobs/job.txt

# Run
python apply.py --job jobs/job.txt --company "Acme" --role "Solutions Engineer"

# With hiring manager name
python apply.py --job jobs/job.txt --company "Acme" --role "SE" --contact "Jane Smith"

# Debug mode (keeps unpacked/ and gen scripts for inspection)
python apply.py --job jobs/job.txt --company "Acme" --role "SE" --debug
```

Output files land in `output/[Company]_[Role]/`:
- `Resume_CoreyLaverdiere_[Company]_[Role].docx`
- `Resume_CoreyLaverdiere_[Company]_[Role]_ATS.docx`
- `CoverLetter_CoreyLaverdiere_[Company]_[Role].docx`

---

## Local Development

```bash
# Install dependencies
pip install -r requirements.txt
npm install

# Set required env vars
export ANTHROPIC_API_KEY=sk-ant-...
export SESSION_SECRET=any-random-string

# Optional: Tigris S3 for persistent user storage
export AWS_ACCESS_KEY_ID=...
export AWS_SECRET_ACCESS_KEY=...
export AWS_ENDPOINT_URL_S3=https://fly.storage.tigris.dev
export BUCKET_NAME=job-apply-corey

# Start the server
uvicorn api:app --reload --port 8000

# Open the app
open http://localhost:8000
```

If Tigris is not configured, user accounts fall back to local filesystem storage.
If `gdrive_credentials.json` is absent, the Drive upload step is skipped.

---

## Google Drive Setup (one-time)

```bash
# 1. Download OAuth credentials
#    console.cloud.google.com → APIs & Services → Credentials
#    → Create Credentials → OAuth client ID → Desktop app → Download JSON
#    Save as: gdrive_credentials.json (project root)

# 2. Authorize
python3 setup_gdrive.py
```

The token is cached at `~/.config/job-apply/gdrive_token.json` and refreshed
automatically. On Fly.io, the Drive credentials and token are mounted as secrets.

---

## Deployment (Fly.io)

```bash
fly deploy --app job-apply-corey
```

Secrets required on Fly.io:
| Secret | Description |
|--------|-------------|
| `ANTHROPIC_API_KEY` | Claude API key |
| `SESSION_SECRET` | HMAC signing key for session tokens |
| `AWS_ACCESS_KEY_ID` | Tigris key |
| `AWS_SECRET_ACCESS_KEY` | Tigris secret |
| `AWS_ENDPOINT_URL_S3` | `https://fly.storage.tigris.dev` |
| `BUCKET_NAME` | Tigris bucket name |
| `RESEND_API_KEY` | For password-change emails (optional) |

---

## Maintaining the Agent

### If XML replacements start failing
Run with `--debug` to inspect `unpacked/word/document.xml`. Section text drifts
when `master.docx` is edited in Word — update known strings in `profile.md`.

### If cover letter voice drifts
Edit `profile.md` → "Voice & Tone Rules" and "DO NOT" sections.

### If framing angle is consistently wrong for a role type
Edit `CLAUDE.md` → "Common Role Type → Framing Angle Reference" table.

---

## API

See `JobApply.postman_collection.json` for the full request/response reference.
Quick overview:

| Method | Path | Auth | Description |
|--------|------|------|-------------|
| GET | `/api/health` | — | Liveness check |
| POST | `/api/auth/register` | — | Create account |
| POST | `/api/auth/login` | — | Get session cookie |
| POST | `/api/auth/logout` | cookie | Clear session |
| GET | `/api/auth/me` | cookie | Current user info |
| GET | `/api/profile` | cookie | Get profile + resume metadata |
| PUT | `/api/profile` | cookie | Update display name or profile text |
| POST | `/api/profile/resume` | cookie | Replace master resume |
| POST | `/api/profile/password` | cookie | Change password |
| POST | `/api/run` | cookie | Start a resume generation run |
| GET | `/api/run/{id}/stream` | cookie | SSE stream of run progress |
| GET | `/api/run/{id}/status` | cookie | Poll run status |
| GET | `/api/run/{id}/files/{name}` | cookie | Download output file |
| GET | `/api/gdrive/runs` | cookie | List Drive run folders |
| GET | `/api/gdrive/runs/{folder_id}/job_posting` | cookie | Fetch saved JD from Drive |
| POST | `/api/prep` | cookie | Start an interview prep run |
| GET | `/api/prep/{id}/stream` | cookie | SSE stream of prep progress |
| GET | `/api/prep/{id}/status` | cookie | Poll prep status |
| GET | `/api/prep/{id}/files/{name}` | cookie | Download prep DOCX |
