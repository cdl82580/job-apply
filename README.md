# job-apply — Corey's Job Application Agent

A Claude-powered CLI that takes a job posting and produces a tailored resume +
cover letter DOCX in under 2 minutes.

---

## Quick Start

```bash
# 1. Install dependencies
pip install anthropic
npm install -g docx

# 2. Set your API key
export ANTHROPIC_API_KEY=sk-ant-...

# 3. Paste the job posting
vim jobs/job.txt   # or pbpaste > jobs/job.txt

# 4. Run
python apply.py --job jobs/job.txt --company "Bluehost" --role "Integration Engineer"

# 5. Collect your files
open output/
```

---

## Project Structure

```
job-apply/
├── CLAUDE.md              ← Agent instructions (Claude Code reads this)
├── README.md              ← This file
├── apply.py               ← CLI entry point
├── profile.md             ← Your voice, stories, metrics, preferences
├── resumes/
│   └── master.docx        ← Source-of-truth resume (NEVER modify outputs)
├── scripts/
│   └── office/            ← DOCX unpack/pack/validate scripts
├── jobs/
│   └── job.txt            ← Paste the current job posting here
└── output/                ← Generated files land here
```

---

## Usage

### Basic
```bash
python apply.py --job jobs/job.txt --company "Acme" --role "Solutions Engineer"
```

### With hiring manager name
```bash
python apply.py --job jobs/job.txt --company "Acme" --role "Solutions Engineer" --contact "Jane Smith"
```

### Faster (use Sonnet instead of Opus)
```bash
python apply.py --job jobs/job.txt --company "Acme" --role "SE" --model "claude-sonnet-4-5-20251101"
```

### Debug mode (keeps unpacked/ and gen scripts for inspection)
```bash
python apply.py --job jobs/job.txt --company "Acme" --role "SE" --debug
```

---

## Outputs

Files are written to `output/` with this naming convention:
- `Resume_[Company]_[Role].docx`
- `CoverLetter_[Company]_[Role].docx`

---

## Updating Your Master Resume

When your experience changes, update `resumes/master.docx` directly — this is
the source every run starts from. After updating:
1. Open it in Word and make your edits
2. Save as `resumes/master.docx` (overwrite)
3. Test with a quick run to confirm the edits land correctly

Do NOT update `master.docx` by starting from an output file — always edit the
master directly.

---

## Using with Claude Code (Option A — no API key needed)

Instead of running `apply.py`, you can run the agent interactively in Claude Code:

1. Open Claude Code in this directory (`claude` in your terminal)
2. Claude will read `CLAUDE.md` automatically
3. Say: `Paste the job posting into jobs/job.txt, then run the full workflow for Company X, Role Y`
4. Claude Code will execute all steps and produce the output files

This approach is slightly slower but lets you inspect and redirect mid-workflow.

---

## Maintaining the Agent

### If replacements start failing
The master resume XML changes if you edit `master.docx` in Word. Run with `--debug`
to inspect `unpacked/word/document.xml` and update the known strings in `profile.md`
or `CLAUDE.md` if section text has drifted.

### If the cover letter voice drifts
Edit `profile.md` → "Voice & Tone Rules" and "DO NOT" sections. The analysis
prompt passes the full profile to Claude on every run.

### If the framing angle is consistently wrong for a role type
Edit `CLAUDE.md` → "Common Role Type → Framing Angle Reference" table.

---

## Google Drive Upload (optional)

Output files are uploaded automatically to your `Job Applications` Drive folder after each run.
Files go from disk → Drive API directly — nothing routes through Claude.

### One-time setup

```bash
# 1. Install the Drive API packages
pip install google-api-python-client google-auth-oauthlib google-auth-httplib2

# 2. Download OAuth credentials
#    console.cloud.google.com → APIs & Services → Credentials
#    → Create Credentials → OAuth client ID → Desktop app → Download JSON
#    Save the file as: gdrive_credentials.json  (in this project root)

# 3. Authorize and test
python3 setup_gdrive.py
```

A browser window opens once for Google authorization. The token is cached at
`~/.config/job-apply/gdrive_token.json` and refreshed automatically on future runs.

If `gdrive_credentials.json` is absent, the upload step is skipped and a warning is printed —
all other output files are still produced normally.

---

## Environment

Tested on macOS and Linux. Requires:
- Python 3.8+
- Node.js 16+ with `docx` npm package (`npm install -g docx`)
- `extract-text` (comes with pandoc: `brew install pandoc` or `apt install pandoc`)
- `ANTHROPIC_API_KEY` environment variable
