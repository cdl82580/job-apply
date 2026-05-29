#!/usr/bin/env python3
"""
apply.py — Job Application Agent for Corey Laverdiere

Public API (for a UI layer):
    from apply import run_workflow, WorkflowConfig, WorkflowResult, WorkflowError

    result = run_workflow(
        job_posting="...",
        company="Acme",
        role="Solutions Engineer",
        config=WorkflowConfig(progress=my_callback),
    )

CLI:
    python apply.py --job jobs/job.txt --company "Acme" --role "Solutions Engineer"
    python apply.py --job jobs/job.txt --company "Acme" --role "SE" --dry-run
"""

import argparse
import html
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
import textwrap
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Callable

# Load .env before anything else so BRANDFETCH_API_KEY / ANTHROPIC_API_KEY are set
_env_path = Path(__file__).parent / ".env"
if _env_path.exists():
    for _line in _env_path.read_text(encoding="utf-8").splitlines():
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _k, _v = _line.split("=", 1)
            os.environ.setdefault(_k.strip(), _v.strip())

from scripts.brand_color import get_brand_color

try:
    import anthropic
except ImportError as _e:
    raise ImportError("anthropic package not installed. Run: pip install anthropic") from _e

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError as _e:
    raise ImportError("Pillow not installed. Run: pip install Pillow") from _e

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

APPLICANT_NAME = "CoreyLaverdiere"

# Single source of truth for contact info — used by cover letter and ATS resume
APPLICANT_CONTACT_LINE = (
    "978-790-4272  |  cdl825@gmail.com  |  Sterling, MA  |  linkedin.com/in/coreydlaverdiere"
)
APPLICANT_CONTACT_LINE_ATS = APPLICANT_CONTACT_LINE + "  |  Open to Remote"

MASTER_RESUME = Path("resumes/master.docx")
PROFILE_FILE  = Path("profile.md")
UNPACK_DIR    = Path("unpacked")
OUTPUT_DIR    = Path("output")
SCRIPTS_DIR   = Path("scripts/office")

DEFAULT_MODEL = "claude-opus-4-5-20251101"

ROUND_TYPES = (
    "Phone Screen",
    "Hiring Manager",
    "Peer",
    "Technical",
    "Executive",
    "Panel",
)

# ---------------------------------------------------------------------------
# WorkflowError / WorkflowConfig / WorkflowResult
# ---------------------------------------------------------------------------

class WorkflowError(Exception):
    """Raised when the workflow cannot continue due to an unrecoverable error."""


@dataclass
class WorkflowConfig:
    """Runtime settings for a single workflow run."""
    model:         str                    = DEFAULT_MODEL
    progress:      Callable[[str], None]  = field(default=print)
    debug:         bool                   = False
    dry_run:       bool                   = False
    # Per-user overrides — set by the server for multi-user deployments.
    # CLI single-user runs leave these as None and fall back to module constants.
    master_resume: Path | None            = None
    profile_text:  str | None             = None
    # User identity — used to scope output dirs and Drive folders.
    # CLI runs leave these None (outputs go to output/ directly).
    user_id:       str | None             = None   # UUID, used for local path
    user_label:    str | None             = None   # email, used for Drive folder name


@dataclass
class WorkflowResult:
    """Paths and metadata produced by a completed workflow run."""
    run_dir:           Path
    resume_path:       Path
    ats_path:          Path
    cover_letter_path: Path
    framing_angle:     str
    folder_url:        str | None = None


@dataclass
class InterviewPrepConfig:
    """Settings for a single interview-prep run."""
    round_type:    str
    focus:         str
    model:         str                    = DEFAULT_MODEL
    progress:      Callable[[str], None]  = field(default=print)
    profile_text:  str | None             = None
    master_resume: Path | None            = None
    user_id:       str | None             = None
    user_label:    str | None             = None
    interviewer:   str                    = ""


@dataclass
class InterviewPrepResult:
    """Paths produced by a completed interview-prep run."""
    prep_path:  Path
    run_dir:    Path
    folder_url: str | None = None

# ---------------------------------------------------------------------------
# Anthropic client — lazy init so import never fails on missing API key
# ---------------------------------------------------------------------------

_client: anthropic.Anthropic | None = None


def _get_client() -> anthropic.Anthropic:
    global _client
    if _client is None:
        _client = anthropic.Anthropic()
    return _client


def claude(system: str, user: str, max_tokens: int = 4096,
           config: WorkflowConfig | None = None) -> str:
    """Single-turn Claude call. Returns the text response."""
    model = config.model if config else DEFAULT_MODEL
    response = _get_client().messages.create(
        model=model,
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    return response.content[0].text

# ---------------------------------------------------------------------------
# Tagline width validation
# ---------------------------------------------------------------------------

# Calibri regular from the Word app bundle — same font Word uses to render the resume
_CALIBRI_PATH = (
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts/Calibri.ttf"
)
_MEASURE_PT = 110
_MASTER_TAGLINE = (
    "Delivering AI-Powered Integrations, Workflow Automations "
    "& Agentic Solutions Across the Full Enterprise Stack"
)
_MAX_TAGLINE_PX: float | None = None


def _measure_width(text: str) -> float:
    """Return rendered pixel width of text at Calibri _MEASURE_PT."""
    font = ImageFont.truetype(_CALIBRI_PATH, size=_MEASURE_PT)
    img  = Image.new("RGB", (1, 1))
    draw = ImageDraw.Draw(img)
    bbox = draw.textbbox((0, 0), text, font=font)
    return float(bbox[2] - bbox[0])


def tagline_fits(text: str) -> bool:
    """Return True if text fits on one printed line at Calibri 11pt.
    Returns True without checking when Calibri is unavailable (no Word install)."""
    global _MAX_TAGLINE_PX
    try:
        if _MAX_TAGLINE_PX is None:
            _MAX_TAGLINE_PX = _measure_width(_MASTER_TAGLINE)
        return _measure_width(text) <= _MAX_TAGLINE_PX
    except OSError:
        return True  # Calibri font not found — skip validation

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def run(
    cmd: str | list,
    check: bool = True,
    config: WorkflowConfig | None = None,
) -> subprocess.CompletedProcess:
    """Run a command. Raises WorkflowError on failure when check=True.
    Pass a list for safe argument handling; strings run through shell=True."""
    shell = isinstance(cmd, str)
    result = subprocess.run(cmd, shell=shell, capture_output=True, text=True)
    if check and result.returncode != 0:
        cmd_display = cmd if shell else " ".join(shlex.quote(str(c)) for c in cmd)
        # Build a detailed error that includes all output so the cause is always visible
        detail_parts = [f"Command failed: {cmd_display}"]
        if result.stdout.strip():
            detail_parts.append(result.stdout.strip())
        if result.stderr.strip():
            detail_parts.append(result.stderr.strip())
        detail = "\n\n".join(detail_parts)
        progress = config.progress if config else print
        progress(detail)
        raise WorkflowError(detail)
    return result


def read_file(path: Path) -> str:
    return path.read_text(encoding="utf-8")

def write_file(path: Path, content: str):
    path.write_text(content, encoding="utf-8")

def safe_filename(s: str) -> str:
    """Strip characters that are not safe for use in filenames."""
    return re.sub(r"[^A-Za-z0-9_-]", "", s)

def print_step(n: str | int, title: str, config: WorkflowConfig | None = None):
    progress = config.progress if config else print
    progress(f"\n{'='*60}")
    progress(f"  STEP {n}: {title}")
    progress(f"{'='*60}")


def extract_resume_text(config: WorkflowConfig | None = None) -> str:
    """Extract plain text from the master resume using pandoc."""
    resume = (config.master_resume if config and config.master_resume else MASTER_RESUME)
    result = run(["pandoc", str(resume), "-t", "plain"], config=config)
    return result.stdout


def read_document_xml() -> str:
    return (UNPACK_DIR / "word" / "document.xml").read_text(encoding="utf-8")

def write_document_xml(content: str):
    (UNPACK_DIR / "word" / "document.xml").write_text(content, encoding="utf-8")

# ---------------------------------------------------------------------------
# Step 1: Read inputs
# ---------------------------------------------------------------------------

def step1_read_inputs(
    job_posting: str,
    config: WorkflowConfig,
) -> tuple[str, str, str]:
    """Validate prerequisites, read master resume and profile.
    Returns (job_posting, resume_text, profile)."""
    print_step(1, "Reading Inputs", config)

    resume = config.master_resume if config.master_resume else MASTER_RESUME
    if not resume.exists():
        raise WorkflowError(f"Master resume not found at {resume}")
    if config.profile_text is None and not PROFILE_FILE.exists():
        raise WorkflowError(f"Profile not found at {PROFILE_FILE} and no profile_text provided")
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise WorkflowError("ANTHROPIC_API_KEY environment variable not set")

    resume_text = extract_resume_text(config)
    profile     = config.profile_text if config.profile_text is not None else read_file(PROFILE_FILE)

    config.progress(f"  ✓ Job posting ({len(job_posting)} chars)")
    config.progress(f"  ✓ Master resume ({len(resume_text)} chars)")
    config.progress(f"  ✓ Profile ({len(profile)} chars)")

    return job_posting, resume_text, profile

# ---------------------------------------------------------------------------
# Step 1b: Extract static sections (ProdPerfect / Applause / Fidelity)
# ---------------------------------------------------------------------------

def step1b_extract_static_sections(resume_text: str, config: WorkflowConfig) -> dict:
    """Parse static employer bullets from the master resume text."""
    system = """\
You are a resume parser. Extract bullet points for specific employers from the resume text.
Return ONLY valid JSON with no preamble or markdown fences.
"""
    prompt = f"""
Resume text:
---
{resume_text}
---

Extract the bullet points for each of these three employers:
1. ProdPerfect
2. Applause (may also appear as "Applause App Quality")
3. Fidelity Investments (may also appear as "Fidelity")

Return this JSON structure:
{{
  "prodperfect_bullets": ["exact bullet text 1", "exact bullet text 2"],
  "applause_bullets": ["exact bullet text 1", "exact bullet text 2"],
  "fidelity_bullets": ["exact bullet text 1", "exact bullet text 2"]
}}

Return ONLY valid JSON.
"""
    raw = claude(system, prompt, max_tokens=2000, config=config)
    raw = re.sub(r"^```json\s*", "", raw.strip())
    raw = re.sub(r"\s*```$", "", raw.strip())

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        config.progress(f"\n⚠️  Failed to parse static bullets: {e} — using empty fallback.")
        data = {}

    static = {
        "ehealth": {
            "company": "eHealth Technologies",
            "dates": "February 2025 – March 2026",
        },
        "hsp": {
            "company": "HSP Group",
            "title": "Senior Business Solutions Development Engineer",
            "dates": "March 2023 – November 2024",
        },
        "prodperfect": {
            "company": "ProdPerfect",
            "title": "Customer Experience & Business Solutions Engineer",
            "dates": "January 2021 – June 2022",
            "bullets": data.get("prodperfect_bullets", []),
        },
        "applause": {
            "company": "Applause App Quality",
            "title": "Senior TPM / Delivery Operations Support Manager",
            "dates": "December 2011 – January 2020",
            "bullets": data.get("applause_bullets", []),
        },
        "fidelity": {
            "company": "Fidelity Investments",
            "title": "Senior Software QA Engineer",
            "dates": "September 2003 – October 2011",
            "bullets": data.get("fidelity_bullets", []),
        },
        "education": [
            {"degree": "MBA, Management Information Systems", "school": "Clark University, Worcester, MA"},
            {"degree": "B.S. (Commonwealth Honors College)", "school": "University of Massachusetts, Amherst, MA"},
            {"degree": "Graduate Certificate, Geographic Information Systems", "school": "Penn State, World Campus"},
        ],
        "certifications": [
            "Tray Build Practitioner & Foundations — Tray.ai",
            "Associate Flow Essentials — Boomi",
            "Professional Flow Developer — Boomi",
            "Microsoft Certified: Azure Fundamentals (AZ-900)",
            "Microsoft Certified: Power Platform Fundamentals (PL-900)",
            "ServiceNow Flow Designer Micro-Certification",
            "Lean Six Sigma Green Belt",
        ],
    }

    config.progress(
        f"  ✓ Static sections extracted "
        f"(ProdPerfect: {len(static['prodperfect']['bullets'])} bullets, "
        f"Applause: {len(static['applause']['bullets'])} bullets, "
        f"Fidelity: {len(static['fidelity']['bullets'])} bullets)"
    )
    return static

# ---------------------------------------------------------------------------
# Step 2: Analysis
# ---------------------------------------------------------------------------

ANALYSIS_SYSTEM = """\
You are a professional resume strategist specializing in tailoring resumes and
cover letters for technical roles (Integration Engineer, Solutions Engineer,
Professional Services, Technical Account Manager, AI/ML roles, etc.).

You will be given:
1. A job posting
2. The candidate's master resume (extracted text)
3. A profile guide describing the candidate's voice, stories, and preferences

Your job is to produce a structured analysis that will drive all downstream
resume and cover letter edits. Be specific and opinionated. Do not hedge.

CRITICAL TAGLINE CONSTRAINT: The resume tagline must fit on a single printed line.
Keep it under 100 characters. Do not use long compound phrases or stack multiple
clauses — one punchy sentence that captures the framing angle is enough.

CRITICAL COVER LETTER CONSTRAINT: The cover letter must fit on a single page.
Write each paragraph (P1-P5) to be 2-4 sentences maximum. Total word count
across all five paragraphs must not exceed 300 words. Be concise and punchy -
every sentence must earn its place. Cut anything that restates what the previous
sentence already said.
"""


def step2_analyze(
    job_posting: str,
    resume_text: str,
    profile: str,
    company: str,
    role: str,
    contact: str | None,
    config: WorkflowConfig,
) -> dict:
    """Run the analysis pass. Returns a structured dict driving all downstream edits."""
    print_step(2, "Analysis", config)

    prompt = f"""
Job Posting:
---
{job_posting}
---

Master Resume:
---
{resume_text}
---

Profile Guide:
---
{profile}
---

Company: {company}
Role: {role}

Produce a JSON object with exactly these keys:
{{
  "role_type": "string - one of: PS/Delivery, Solutions Engineer, TAM, Integration Engineer, Agent Platform, AI Solutions, Customer Success, Forward Deployed Engineer, Other",
  "framing_angle": "string - 1-2 sentences describing the single narrative thread to run through the entire resume and cover letter",
  "tagline": "string - new tagline for the resume header (1 sentence, punchy, matches framing angle, MUST be under 100 characters)",
  "top_jd_requirements": ["string", "string", "string", "string", "string"],
  "competencies": ["14 strings, one per cell, in order: row1col1, row1col2, row1col3, row1col4, row1col5, row2col1, row2col2, row2col3, row2col4, row2col5, row3col1, row3col2, row3col3, row3col4"],
  "ehealth_title_subtitle": "string - the subtitle bar text for eHealth (e.g. 'AI Solutions & Integration Engineer  |  Subtitle  |  Tray.ai Platform Owner')",
  "ehealth_bullets": ["6 strings - complete bullet text for each of the 6 eHealth bullets"],
  "hsp_bullets": ["4 strings - complete bullet text for each of the 4 HSP Group bullets"],
  "summary": "string - full professional summary text (4-5 sentences, written in Corey's voice per profile.md)",
  "cover_letter_hook": "string - the opening angle for the cover letter P1 (what JD language to echo, what story to lead with)",
  "cover_letter_p1": "string - full text of P1 (max 3 sentences)",
  "cover_letter_p2": "string - full text of P2, primary evidence, most quantified (max 4 sentences)",
  "cover_letter_p3": "string - full text of P3, secondary evidence (max 3 sentences)",
  "cover_letter_p4": "string - full text of P4, differentiator specific to this role/company (max 3 sentences)",
  "cover_letter_p5": "string - full text of P5, short close (1-2 sentences only)",
  "contact_name": "string - hiring manager name if determinable from the posting, otherwise 'Hiring Team'"
}}

Return ONLY valid JSON. No preamble, no markdown fences, no commentary.
"""
    raw = claude(ANALYSIS_SYSTEM, prompt, max_tokens=6000, config=config)
    raw = re.sub(r"^```json\s*", "", raw.strip())
    raw = re.sub(r"\s*```$", "", raw.strip())

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise WorkflowError(
            f"Failed to parse analysis JSON: {e}\n\nRaw response:\n{raw[:2000]}"
        )

    # Caller-supplied contact overrides anything the model inferred
    if contact:
        data["contact_name"] = contact

    # Validate tagline width — retry up to 2 times if it overflows one line
    for attempt in range(2):
        tagline = data.get("tagline", "")
        if tagline_fits(tagline):
            break
        try:
            ratio = _measure_width(tagline) / _MAX_TAGLINE_PX if _MAX_TAGLINE_PX else 1.0
        except OSError:
            ratio = 1.0
        config.progress(f"\n  ⚠  Tagline too wide ({len(tagline)} chars, {ratio:.0%} of max line):")
        config.progress(f"     {tagline}")
        config.progress(f"     Requesting shorter version (attempt {attempt + 1}/2)...")
        shortened = claude(
            "You are a resume copywriter. Shorten the given tagline so it fits on one "
            "printed line of a resume. Keep the core meaning and active voice. "
            "Return only the shortened tagline — no quotes, no explanation.",
            f'Tagline to shorten: {tagline}\n\nConstraint: must be under 100 characters.',
            max_tokens=120,
            config=config,
        )
        data["tagline"] = shortened.strip().strip('"').strip("'")
    else:
        tagline = data.get("tagline", "")
        if not tagline_fits(tagline):
            config.progress(f"  ⚠  Tagline still too wide after 2 retries — proceeding anyway.")
            config.progress(f"     {tagline}")

    config.progress(f"\n  Role type:      {data.get('role_type')}")
    config.progress(f"  Framing angle:  {data.get('framing_angle')}")
    config.progress(f"  Tagline:        {data.get('tagline')}")
    config.progress(f"\n  Top JD requirements:")
    for i, req in enumerate(data.get("top_jd_requirements", []), 1):
        config.progress(f"    {i}. {req}")

    return data

# ---------------------------------------------------------------------------
# Step 2b: Brand colors
# ---------------------------------------------------------------------------

def step2b_brand_colors(company: str, config: WorkflowConfig) -> dict:
    print_step("2b", "Fetching Brand Colors", config)
    return get_brand_color(company)

# ---------------------------------------------------------------------------
# Steps 3–5: Resume build
# ---------------------------------------------------------------------------

def step3_unpack(config: WorkflowConfig):
    print_step(3, "Unpacking Master Resume", config)
    resume = config.master_resume if config.master_resume else MASTER_RESUME
    if UNPACK_DIR.exists():
        shutil.rmtree(UNPACK_DIR)
    run(
        ["python3", str(SCRIPTS_DIR / "unpack.py"), str(resume), str(UNPACK_DIR) + "/"],
        config=config,
    )
    config.progress("  ✓ Unpacked")


def apply_brand_colors(xml: str, colors: dict) -> str:
    """Replace the three hardcoded palette hex values with the brand colors."""
    xml = xml.replace('w:val="1A3C5E"',  f'w:val="{colors["primary"]}"')
    xml = xml.replace('w:color="1A3C5E"', f'w:color="{colors["primary"]}"')
    xml = xml.replace('w:color="2B6CB0"', f'w:color="{colors["border"]}"')
    xml = xml.replace('w:fill="EEF4FB"',  f'w:fill="{colors["fill"]}"')
    return xml


def _xml_escape(text: str) -> str:
    """Escape text for safe insertion as XML character data.
    Resolves any pre-escaped entities first to avoid double-encoding, then
    re-escapes cleanly — so Claude can write & or &amp; and both work."""
    for entity, char in [("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
                         ("&apos;", "'"), ("&quot;", '"')]:
        text = text.replace(entity, char)
    return html.escape(text, quote=False)


def step4_apply_edits(
    analysis: dict,
    resume_text: str,
    colors: dict | None,
    config: WorkflowConfig,
) -> int:
    """Apply all content edits and brand colors to the unpacked XML.
    Returns the number of successful replacements."""
    print_step(4, "Applying Resume Edits", config)

    xml = read_document_xml()
    total_success = 0
    total_attempted = 0

    system = """\
You are a DOCX XML editor. You will be given:
1. The current document.xml content (extracted text form of the resume)
2. The new values for each section

Your job: produce a JSON array of replacement operations. Each operation is:
  {"old": "exact string currently in the XML", "new": "replacement string"}

Rules:
- "old" must be the EXACT text currently in the XML - copy it character-for-character
- "new" is the new text to substitute
- Use \\u2014 for em dash (-), \\u2019 for right single quote ('), \\u2013 for en dash (-)
- Do NOT use XML entities (&amp;, &lt;) in the "new" values - write & and < directly;
  the packing script handles XML encoding
- The XML uses &amp; for & in the stored content - when writing "old", match exactly
  including &amp; if present
- Return ONLY a valid JSON array. No preamble, no markdown fences.
"""

    user = f"""
Current resume text (for context):
---
{resume_text[:8000]}
---

Desired new values:
- tagline: {analysis['tagline']}
- summary: {analysis['summary']}
- ehealth_title_subtitle: {analysis['ehealth_title_subtitle']}
- ehealth_bullets (6): {json.dumps(analysis['ehealth_bullets'], indent=2)}
- hsp_bullets (4): {json.dumps(analysis['hsp_bullets'], indent=2)}
- competencies (14, in row order): {json.dumps(analysis['competencies'], indent=2)}

Produce the JSON array of replacement operations. For each section, find the
current text in the resume and produce an exact old->new pair.

The competency cells currently contain these values (in order):
Row 1: Agentic AI & LLM Systems | RAG Pipelines & Prompt Engineering | REST, SOAP & GraphQL APIs | Tray.ai / iPaaS Platform Ownership | Solution Architecture & Delivery
Row 2: End-to-End Integration & Automation Delivery | Salesforce CRM & Administration | Microsoft 365 / Graph API | POC-to-Production Deployment | Six Sigma Green Belt
Row 3: JavaScript / JSON / SQL | Workday & Okta Integration | Stakeholder Enablement & AI Literacy | Technical Documentation & ROI Reporting | Cross-functional Collaboration

(Note: row 2 has 5 cells but only 4 competencies are active - the 5th cell "Cross-functional Collaboration" is the last one in row 3, not row 2.)

The tagline currently contains:
"Delivering AI-Powered Integrations, Workflow Automations & Agentic Solutions Across the Full Enterprise Stack"

Return ONLY valid JSON array.
"""

    raw = claude(system, user, max_tokens=8000, config=config)
    raw = re.sub(r"^```json\s*", "", raw.strip())
    raw = re.sub(r"\s*```$", "", raw.strip())

    try:
        ops = json.loads(raw)
    except json.JSONDecodeError as e:
        raise WorkflowError(f"Failed to parse replacement ops JSON: {e}\n\n{raw[:2000]}")

    for op in ops:
        old = op.get("old", "")
        new = op.get("new", "")
        total_attempted += 1
        safe_new = _xml_escape(new)

        if old in xml:
            # Exact match (Claude already used &amp; in the old value)
            xml = xml.replace(old, safe_new, 1)
            total_success += 1
            config.progress(f"  ✓ Replaced: {old[:60]}...")
        else:
            # Claude read plain text (bare &) but the XML has &amp; — try normalised form
            xml_old = _xml_escape(old)
            if xml_old != old and xml_old in xml:
                xml = xml.replace(xml_old, safe_new, 1)
                total_success += 1
                config.progress(f"  ✓ Replaced (normalised): {old[:60]}...")
            else:
                config.progress(f"  ✗ NOT FOUND: {old[:60]}...")

    if colors:
        xml = apply_brand_colors(xml, colors)
        config.progress(f"  ✓ Brand colors applied (primary=#{colors['primary']})")

    write_document_xml(xml)
    config.progress(f"\n  Result: {total_success}/{total_attempted} replacements succeeded")

    if total_success < total_attempted * 0.7:
        config.progress(f"\n⚠️  Warning: fewer than 70% of replacements succeeded.")
        config.progress(f"   Check the XML manually or re-run with --debug flag.")

    return total_success


def step5_pack(resume_out: Path, config: WorkflowConfig):
    print_step(5, "Packing Resume", config)
    resume = config.master_resume if config.master_resume else MASTER_RESUME
    run(
        ["python3", str(SCRIPTS_DIR / "pack.py"), str(UNPACK_DIR) + "/",
         str(resume_out), "--original", str(resume)],
        config=config,
    )
    config.progress(f"  ✓ Resume written to {resume_out}")


def step7_cleanup(config: WorkflowConfig):
    print_step(7, "Cleanup", config)
    if not config.debug and UNPACK_DIR.exists():
        shutil.rmtree(UNPACK_DIR)

# ---------------------------------------------------------------------------
# JS string escaping (shared by ATS resume and cover letter builders)
# ---------------------------------------------------------------------------

def escape_js_string(s: str) -> str:
    """Escape a string for embedding in a JS double-quoted string."""
    s = s.replace("\\", "\\\\")
    s = s.replace("`", "\\`")
    s = s.replace("${", "\\${")
    s = s.replace('"', '\\"')
    return s

# ---------------------------------------------------------------------------
# Step 5b: ATS Resume
# ---------------------------------------------------------------------------

def step5b_ats_resume(
    analysis: dict,
    static_sections: dict,
    company: str,
    role: str,
    output_path: Path,
    config: WorkflowConfig,
):
    """Generate a clean, ATS-optimized single-column DOCX."""
    print_step("5b", "Generating ATS Resume", config)

    paras: list[str] = []

    def tr(text: str, bold: bool = False, italic: bool = False, size: int = 22) -> str:
        cleaned = " ".join(text.split())
        escaped = escape_js_string(cleaned)
        props = [f'text: "{escaped}"', 'font: "Calibri"', f'size: {size}', 'color: "000000"']
        if bold:
            props.append("bold: true")
        if italic:
            props.append("italic: true")
        return "new TextRun({ " + ", ".join(props) + " })"

    def add(children_strs: list, before: int = 0, after: int = 80, left: int = 0):
        spacing = f"before: {before}, after: {after}"
        indent  = f", indent: {{ left: {left} }}" if left else ""
        paras.append(
            f"      new Paragraph({{ spacing: {{ {spacing} }}{indent}, "
            f"children: [{', '.join(children_strs)}] }})"
        )

    def heading(text: str):
        add([tr(text, bold=True, size=24)], before=240, after=60)

    def body(text: str, after: int = 80):
        add([tr(text)], after=after)

    def bullet(text: str):
        add([tr("•  " + text)], after=40, left=360)

    def job_header(company_name: str, title: str, dates: str):
        children = [tr(company_name, bold=True)]
        if dates:
            children.append(tr("  |  " + dates))
        add(children, before=200, after=0)
        if title:
            add([tr(title, italic=True)], after=40)

    # Name + contact
    add([tr("COREY LAVERDIERE", bold=True, size=40)], after=0)
    add([tr(APPLICANT_CONTACT_LINE_ATS, size=20)], after=120)

    # Tagline
    add([tr(analysis.get("tagline", ""), italic=True)], after=160)

    # Professional Summary
    heading("Professional Summary")
    body(analysis.get("summary", ""), after=0)

    # Core Competencies
    heading("Core Competencies")
    comps = analysis.get("competencies", [])
    for i in range(0, len(comps), 5):
        body(" | ".join(comps[i:i+5]), after=40)

    # Professional Experience
    heading("Professional Experience")

    ehealth = static_sections["ehealth"]
    job_header("eHealth Technologies", analysis.get("ehealth_title_subtitle", ""), ehealth["dates"])
    for b in analysis.get("ehealth_bullets", []):
        bullet(b)

    hsp = static_sections["hsp"]
    job_header(hsp["company"], hsp["title"], hsp["dates"])
    for b in analysis.get("hsp_bullets", []):
        bullet(b)

    for key in ("prodperfect", "applause", "fidelity"):
        section = static_sections[key]
        job_header(section["company"], section["title"], section["dates"])
        for b in section.get("bullets", []):
            bullet(b)

    # Education
    heading("Education")
    for edu in static_sections.get("education", []):
        children = [tr(edu["degree"], bold=True)]
        if edu.get("school"):
            children.append(tr("  —  " + edu["school"]))
        add(children, before=60, after=40)

    # Certifications
    heading("Certifications")
    for cert in static_sections.get("certifications", []):
        body(cert, after=40)

    children_js  = ",\n".join(paras)
    out_path_str = str(output_path).replace("\\", "/")

    js = f"""\
const {{ Document, Packer, Paragraph, TextRun }} = require('docx');
const fs = require('fs');

const doc = new Document({{
  styles: {{ default: {{ document: {{ run: {{ font: "Calibri", size: 22 }} }} }} }},
  sections: [{{
    properties: {{
      page: {{
        size: {{ width: 12240, height: 15840 }},
        margin: {{ top: 720, right: 1080, bottom: 720, left: 1080 }}
      }}
    }},
    children: [
{children_js}
    ]
  }}]
}});

Packer.toBuffer(doc).then(buffer => {{
  fs.writeFileSync('{out_path_str}', buffer);
  console.log('ATS resume written.');
}});
"""

    js_path = Path("ats_resume_gen.js")
    write_file(js_path, js)
    result = run(["node", str(js_path)], check=False, config=config)
    js_path.unlink(missing_ok=True)
    if result.returncode != 0:
        raise WorkflowError(f"ATS resume JS failed:\n{result.stderr}")

    config.progress(f"  ✓ ATS resume written to {output_path}")

# ---------------------------------------------------------------------------
# Step 6: Cover letter
# ---------------------------------------------------------------------------

# docx size units are half-points: 22 = 11pt, 24 = 12pt, 40 = 20pt
# twip spacing: 80 = tight gap, 240 = 1.0 line height, 720 = 0.5in margin

COVER_LETTER_JS_TEMPLATE = """\
const {{ Document, Packer, Paragraph, TextRun, BorderStyle }} = require('docx');
const fs = require('fs');

const doc = new Document({{
  styles: {{ default: {{ document: {{ run: {{ font: "Calibri", size: 22 }} }} }} }},
  sections: [{{
    properties: {{
      page: {{
        size: {{ width: 12240, height: 15840 }},
        margin: {{ top: 1080, right: 1080, bottom: 1080, left: 1080 }}
      }}
    }},
    children: [
      new Paragraph({{
        spacing: {{ after: 0 }},
        children: [new TextRun({{ text: "COREY LAVERDIERE", font: "Calibri", size: 40, bold: true, color: "{primary_color}" }})]
      }}),
      new Paragraph({{
        border: {{ bottom: {{ style: BorderStyle.SINGLE, size: 6, color: "{border_color}", space: 4 }} }},
        spacing: {{ after: 160 }},
        children: [new TextRun({{
          text: "{contact_line}",
          font: "Calibri", size: 20, color: "6B7280"
        }})]
      }}),
      new Paragraph({{
        spacing: {{ before: 160, after: 60 }},
        children: [new TextRun({{ text: "{today}", font: "Calibri", size: 22, color: "111827" }})]
      }}),
      new Paragraph({{
        spacing: {{ after: 0 }},
        children: [new TextRun({{ text: "{contact_name}", font: "Calibri", size: 22, color: "111827" }})]
      }}),
      new Paragraph({{
        spacing: {{ after: 60 }},
        children: [new TextRun({{ text: "{company}", font: "Calibri", size: 22, color: "111827" }})]
      }}),
      new Paragraph({{
        spacing: {{ before: 60, after: 160 }},
        children: [new TextRun({{ text: "Re: {role}", font: "Calibri", size: 22, bold: true, color: "111827" }})]
      }}),
      new Paragraph({{
        spacing: {{ after: 160 }},
        children: [new TextRun({{ text: "Dear {salutation},", font: "Calibri", size: 22, color: "111827" }})]
      }}),
      {body_paragraphs}
      new Paragraph({{
        spacing: {{ after: 40 }},
        children: [new TextRun({{ text: "Sincerely,", font: "Calibri", size: 22, color: "111827" }})]
      }}),
      new Paragraph({{
        spacing: {{ after: 40 }},
        children: [new TextRun({{ text: "Corey Laverdiere", font: "Calibri", size: 22, bold: true, color: "{primary_color}" }})]
      }}),
      new Paragraph({{
        spacing: {{ after: 0 }},
        children: [new TextRun({{ text: "{sign_off_contact}", font: "Calibri", size: 22, color: "6B7280" }})]
      }})
    ]
  }}]
}});

Packer.toBuffer(doc).then(buffer => {{
  fs.writeFileSync('{output_path}', buffer);
  console.log('Cover letter written.');
}});
"""


def step6_cover_letter(
    analysis: dict,
    company: str,
    role: str,
    output_path: Path,
    config: WorkflowConfig,
    colors: dict | None = None,
):
    print_step(6, "Generating Cover Letter", config)

    palette      = colors or {"primary": "1A3C5E", "border": "2B6CB0"}
    today        = date.today().strftime("%B %-d, %Y")
    contact_name = analysis.get("contact_name", "Hiring Team")
    salutation   = contact_name if contact_name != "Hiring Team" else "Hiring Team"

    body_keys = ["cover_letter_p1", "cover_letter_p2", "cover_letter_p3",
                 "cover_letter_p4", "cover_letter_p5"]

    body_paragraphs = []
    for i, key in enumerate(body_keys):
        text    = analysis.get(key, "")
        escaped = escape_js_string(text)
        after   = 200 if i == len(body_keys) - 1 else 160
        body_paragraphs.append(
            f'      new Paragraph({{\n'
            f'        spacing: {{ after: {after} }},\n'
            f'        children: [new TextRun({{ text: "{escaped}", font: "Calibri", size: 22, color: "111827" }})]\n'
            f'      }}),'
        )

    # Sign-off uses only phone | email (not full contact line)
    sign_off_contact = escape_js_string(
        "  |  ".join(APPLICANT_CONTACT_LINE.split("  |  ")[:2])
    )

    js = COVER_LETTER_JS_TEMPLATE.format(
        today=today,
        contact_name=escape_js_string(contact_name),
        company=escape_js_string(company),
        role=escape_js_string(role),
        salutation=escape_js_string(salutation),
        body_paragraphs="\n".join(body_paragraphs),
        output_path=str(output_path).replace("\\", "/"),
        primary_color=palette["primary"],
        border_color=palette["border"],
        contact_line=escape_js_string(APPLICANT_CONTACT_LINE),
        sign_off_contact=sign_off_contact,
    )

    js_path = Path("cover_letter_gen.js")
    write_file(js_path, js)
    result = run(["node", str(js_path)], check=False, config=config)
    js_path.unlink(missing_ok=True)  # always clean up, even on failure
    if result.returncode != 0:
        raise WorkflowError(f"Cover letter JS failed:\n{result.stderr}")

    config.progress(f"  ✓ Cover letter written to {output_path}")

# ---------------------------------------------------------------------------
# Step 8: Google Drive upload
# ---------------------------------------------------------------------------

GDRIVE_PARENT_FOLDER_ID = os.environ.get("GDRIVE_PARENT_FOLDER_ID", "")
GDRIVE_TOKEN_PATH       = Path.home() / ".config" / "job-apply" / "gdrive_token.json"
GDRIVE_CREDS_PATH       = Path(__file__).parent / "gdrive_credentials.json"
_MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
_SCOPES    = ["https://www.googleapis.com/auth/drive.file"]


def _seed_gdrive_token() -> None:
    """Write the GDRIVE_TOKEN_JSON env var to disk if the token file isn't there yet.

    On the server the OAuth browser flow can't run, so we store the token as a
    Fly.io secret (GDRIVE_TOKEN_JSON) and materialize it at runtime.  The token
    includes client_id + client_secret, so the Google SDK can refresh it
    automatically without needing the original gdrive_credentials.json file.
    """
    token_json = os.environ.get("GDRIVE_TOKEN_JSON", "").strip()
    if token_json and not GDRIVE_TOKEN_PATH.exists():
        GDRIVE_TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)
        GDRIVE_TOKEN_PATH.write_text(token_json)


def _gdrive_service(config: WorkflowConfig):
    """Return an authenticated Drive v3 service, or None if credentials are missing."""
    try:
        from googleapiclient.discovery import build
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
    except ImportError:
        config.progress("  ⚠ google-api-python-client not installed — skipping Drive upload")
        return None

    _seed_gdrive_token()

    creds = None
    if GDRIVE_TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(GDRIVE_TOKEN_PATH), _SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                GDRIVE_TOKEN_PATH.write_text(creds.to_json())
            except Exception as refresh_err:
                # invalid_grant means the token is permanently revoked — remove it
                # so the next run doesn't hit the same error, and tell the user.
                GDRIVE_TOKEN_PATH.unlink(missing_ok=True)
                config.progress(f"  ⚠ Drive token expired/revoked: {refresh_err}")
                config.progress("    To fix: run locally then update the secret:")
                config.progress("      rm ~/.config/job-apply/gdrive_token.json")
                config.progress("      python3 setup_gdrive.py")
                config.progress('      fly secrets set GDRIVE_TOKEN_JSON="$(cat ~/.config/job-apply/gdrive_token.json)"')
                return None
        elif GDRIVE_CREDS_PATH.exists():
            flow  = InstalledAppFlow.from_client_secrets_file(str(GDRIVE_CREDS_PATH), _SCOPES)
            creds = flow.run_local_server(port=0)
            GDRIVE_TOKEN_PATH.parent.mkdir(parents=True, exist_ok=True)
            GDRIVE_TOKEN_PATH.write_text(creds.to_json())
        else:
            config.progress("  ⚠ Drive upload skipped — set GDRIVE_TOKEN_JSON secret to enable")
            return None

    return build("drive", "v3", credentials=creds)


def _gdrive_get_or_create_folder(service, name: str, parent_id: str) -> tuple[str, str, bool]:
    """Return (folder_id, webViewLink, created) for a named subfolder.

    created=True when the folder was just made; False when it already existed.
    """
    existing = service.files().list(
        q=(
            f"name='{name}' and '{parent_id}' in parents and "
            "mimeType='application/vnd.google-apps.folder' and trashed=false"
        ),
        fields="files(id, webViewLink)",
        pageSize=1,
    ).execute().get("files", [])

    if existing:
        return existing[0]["id"], existing[0]["webViewLink"], False

    created = service.files().create(
        body={"name": name, "mimeType": "application/vnd.google-apps.folder",
              "parents": [parent_id]},
        fields="id, webViewLink",
    ).execute()
    return created["id"], created["webViewLink"], True


def _set_link_viewer(service, folder_id: str, progress: callable) -> None:
    """Grant 'anyone with the link' viewer access to a Drive folder.

    Silently ignores errors — the most common cause is the permission
    already existing (Drive returns a 409 in that case).
    """
    try:
        service.permissions().create(
            fileId=folder_id,
            body={"type": "anyone", "role": "reader"},
            fields="id",
        ).execute()
        progress("  ✓ Drive folder set to 'anyone with the link' viewer access")
    except Exception as exc:
        # 409 = permission already exists; any other error is non-fatal
        progress(f"  ⚠ Could not set Drive folder permissions: {exc}")


def _convert_docx_to_pdf_via_drive(
    service,
    docx_path: Path,
    pdf_name: str,
    folder_id: str,
    progress: callable,
) -> None:
    """Convert a local DOCX to PDF using Drive's conversion pipeline.

    Steps:
      1. Upload the DOCX with mimeType=Google Doc — Drive converts on ingest.
      2. Export the resulting Google Doc as PDF bytes.
      3. Upload the PDF to the run folder.
      4. Delete the temporary Google Doc.

    Best-effort: any exception is logged and swallowed so the caller is
    never blocked by a PDF conversion failure.
    """
    try:
        import io
        from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload

        # Step 1: upload DOCX as Google Doc (Drive handles the conversion)
        gdoc = service.files().create(
            body={"name": f"_tmp_{docx_path.stem}",
                  "mimeType": "application/vnd.google-apps.document"},
            media_body=MediaFileUpload(str(docx_path), mimetype=_MIME_DOCX),
            fields="id",
        ).execute()
        gdoc_id = gdoc["id"]

        try:
            # Step 2: export as PDF
            pdf_bytes = service.files().export(
                fileId=gdoc_id,
                mimeType="application/pdf",
            ).execute()

            # Step 3: upload PDF to the run folder
            service.files().create(
                body={"name": pdf_name, "parents": [folder_id]},
                media_body=MediaIoBaseUpload(
                    io.BytesIO(pdf_bytes), mimetype="application/pdf"
                ),
                fields="id",
            ).execute()
            progress(f"  ✓ Generated PDF: {pdf_name}")

        finally:
            # Step 4: always clean up the temp Google Doc
            try:
                service.files().delete(fileId=gdoc_id).execute()
            except Exception:
                pass

    except Exception as exc:
        progress(f"  ⚠ PDF generation skipped: {exc}")


def step8_upload(
    run_dir: Path,
    company_safe: str,
    role_safe: str,
    config: WorkflowConfig,
) -> str | None:
    """Upload output files to Google Drive. Returns the run folder URL or None.

    Drive structure:
      Job Applications/
        {user_label}/          ← created when config.user_label is set
          {Company}_{Role}/
    """
    try:
        from googleapiclient.http import MediaFileUpload
    except ImportError:
        config.progress("  ⚠ google-api-python-client not installed — skipping Drive upload")
        return None

    print_step(8, "Uploading to Google Drive", config)

    try:
        service = _gdrive_service(config)
        if service is None:
            return None

        # Resolve the parent folder (user subfolder, or root if CLI/no user)
        if config.user_label:
            user_folder_id, _, user_created = _gdrive_get_or_create_folder(
                service, config.user_label, GDRIVE_PARENT_FOLDER_ID
            )
            config.progress(f"  ✓ Drive user folder: {config.user_label}")
            if user_created:
                _set_link_viewer(service, user_folder_id, config.progress)
            run_parent_id = user_folder_id
        else:
            run_parent_id = GDRIVE_PARENT_FOLDER_ID

        run_folder_name = f"{company_safe}_{role_safe}"
        run_folder_id, folder_url, _ = _gdrive_get_or_create_folder(
            service, run_folder_name, run_parent_id
        )
        config.progress(f"  ✓ Drive run folder: {run_folder_name}")

        for f in sorted(run_dir.iterdir()):
            if f.name.startswith("~$"):
                continue
            if f.suffix == ".docx":
                mime = _MIME_DOCX
            elif f.suffix == ".pdf":
                mime = "application/pdf"
            elif f.name == "job_posting.txt":
                mime = "text/plain; charset=utf-8"
            else:
                continue
            media = MediaFileUpload(str(f), mimetype=mime, resumable=False)
            service.files().create(
                body={"name": f.name, "parents": [run_folder_id]},
                media_body=media,
                fields="id",
            ).execute()
            config.progress(f"  ✓ Uploaded {f.name}")

        # Convert the styled (non-ATS) resume to PDF via Drive
        styled_resume = run_dir / f"Resume_{APPLICANT_NAME}_{company_safe}_{role_safe}.docx"
        if styled_resume.exists():
            _convert_docx_to_pdf_via_drive(
                service,
                styled_resume,
                f"Resume_{APPLICANT_NAME}_{company_safe}_{role_safe}.pdf",
                run_folder_id,
                config.progress,
            )

        return folder_url

    except Exception as exc:
        config.progress(f"  ⚠ Drive upload failed: {exc}")
        config.progress("    Files are still available for download below.")
        return None

# ---------------------------------------------------------------------------
# Drive: targeted single-file upload (used by interview prep)
# ---------------------------------------------------------------------------

def _upload_single_to_drive(
    file_path: Path,
    folder_name: str,
    config: WorkflowConfig,
) -> str | None:
    """Upload one file into the correct user → run subfolder in Drive."""
    try:
        from googleapiclient.http import MediaFileUpload
    except ImportError:
        config.progress("  ⚠ google-api-python-client not installed — skipping Drive upload")
        return None

    try:
        service = _gdrive_service(config)
        if service is None:
            return None

        if config.user_label:
            user_folder_id, _, user_created = _gdrive_get_or_create_folder(
                service, config.user_label, GDRIVE_PARENT_FOLDER_ID
            )
            if user_created:
                _set_link_viewer(service, user_folder_id, config.progress)
            run_parent_id = user_folder_id
        else:
            run_parent_id = GDRIVE_PARENT_FOLDER_ID

        folder_id, folder_url, _ = _gdrive_get_or_create_folder(
            service, folder_name, run_parent_id
        )
        config.progress(f"  ✓ Drive folder: {folder_name}")

        media = MediaFileUpload(str(file_path), mimetype=_MIME_DOCX, resumable=False)
        service.files().create(
            body={"name": file_path.name, "parents": [folder_id]},
            media_body=media,
            fields="id",
        ).execute()
        config.progress(f"  ✓ Uploaded {file_path.name}")
        return folder_url

    except Exception as exc:
        config.progress(f"  ⚠ Drive upload failed: {exc}")
        config.progress("    File is still available for download below.")
        return None


# ---------------------------------------------------------------------------
# Drive: list run folders + fetch job posting (used by /api/gdrive/runs)
# ---------------------------------------------------------------------------

_FOLDER_MIME = "application/vnd.google-apps.folder"


def list_gdrive_run_folders(user_label: str, config: WorkflowConfig) -> list[dict]:
    """Return all run folders visible to this user from Google Drive.

    Checks two locations:
      1. Job Applications/{user_label}/  — current per-user structure
      2. Job Applications/ root          — legacy flat runs (skips email-named subfolders)

    Each entry: {name, id, web_view_link, source ("user" | "legacy")}
    Returns [] if Drive is not configured or an error occurs.
    """
    service = _gdrive_service(config)
    if service is None:
        return []

    results: list[dict] = []
    seen_ids: set[str]  = set()

    try:
        # ── 1. User's personal subfolder ──────────────────────────────
        user_roots = service.files().list(
            q=(
                f"name='{user_label}' and '{GDRIVE_PARENT_FOLDER_ID}' in parents and "
                f"mimeType='{_FOLDER_MIME}' and trashed=false"
            ),
            fields="files(id)",
            pageSize=1,
        ).execute().get("files", [])

        if user_roots:
            user_root_id = user_roots[0]["id"]
            for f in service.files().list(
                q=f"'{user_root_id}' in parents and mimeType='{_FOLDER_MIME}' and trashed=false",
                fields="files(id, name, webViewLink)",
                orderBy="modifiedTime desc",
                pageSize=100,
            ).execute().get("files", []):
                results.append({
                    "name":          f["name"],
                    "id":            f["id"],
                    "web_view_link": f.get("webViewLink", ""),
                    "source":        "user",
                })
                seen_ids.add(f["id"])

        # ── 2. Legacy flat root ────────────────────────────────────────
        for f in service.files().list(
            q=(
                f"'{GDRIVE_PARENT_FOLDER_ID}' in parents and "
                f"mimeType='{_FOLDER_MIME}' and trashed=false"
            ),
            fields="files(id, name, webViewLink)",
            orderBy="modifiedTime desc",
            pageSize=100,
        ).execute().get("files", []):
            if f["id"] in seen_ids:
                continue
            # Skip user account folders (named like emails)
            if "@" in f["name"]:
                continue
            results.append({
                "name":          f["name"],
                "id":            f["id"],
                "web_view_link": f.get("webViewLink", ""),
                "source":        "legacy",
            })

    except Exception:
        pass  # best-effort; return whatever we collected

    return results


def get_gdrive_job_posting(folder_id: str, config: WorkflowConfig) -> str | None:
    """Fetch the text of job_posting.txt from a Drive folder. Returns None if absent."""
    service = _gdrive_service(config)
    if service is None:
        return None
    try:
        files = service.files().list(
            q=f"name='job_posting.txt' and '{folder_id}' in parents and trashed=false",
            fields="files(id)",
            pageSize=1,
        ).execute().get("files", [])
        if not files:
            return None
        content = service.files().get_media(fileId=files[0]["id"]).execute()
        return content.decode("utf-8") if isinstance(content, bytes) else str(content)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Public workflow entry point
# ---------------------------------------------------------------------------

def run_workflow(
    job_posting: str,
    company: str,
    role: str,
    contact: str | None = None,
    config: WorkflowConfig | None = None,
) -> WorkflowResult:
    """
    Run the full job-application workflow.

    Args:
        job_posting: Full text of the job posting.
        company:     Company name (used in filenames and cover letter).
        role:        Role title.
        contact:     Hiring manager name, or None to let analysis infer it.
        config:      WorkflowConfig for model, progress callback, debug, dry_run.

    Returns:
        WorkflowResult with paths to generated files and optional Drive URL.

    Raises:
        WorkflowError on any unrecoverable error.
    """
    if config is None:
        config = WorkflowConfig()

    OUTPUT_DIR.mkdir(exist_ok=True)

    company_safe = safe_filename(company)
    role_safe    = safe_filename(role)
    # Scope to user subfolder when running via the server; CLI runs go to output/ directly.
    if config.user_id:
        run_dir = OUTPUT_DIR / safe_filename(config.user_id) / f"{company_safe}_{role_safe}"
    else:
        run_dir = OUTPUT_DIR / f"{company_safe}_{role_safe}"
    run_dir.mkdir(parents=True, exist_ok=True)

    # Persist job posting so interview prep can retrieve it later
    (run_dir / "job_posting.txt").write_text(job_posting, encoding="utf-8")

    resume_out = run_dir / f"Resume_{APPLICANT_NAME}_{company_safe}_{role_safe}.docx"
    ats_out    = run_dir / f"Resume_{APPLICANT_NAME}_{company_safe}_{role_safe}_ATS.docx"
    cover_out  = run_dir / f"CoverLetter_{APPLICANT_NAME}_{company_safe}_{role_safe}.docx"

    config.progress(f"\n\U0001f680 Job Application Agent")
    config.progress(f"   Company : {company}")
    config.progress(f"   Role    : {role}")
    config.progress(f"   Run dir : {run_dir}")
    config.progress(f"   Outputs : {resume_out.name}, {ats_out.name}, {cover_out.name}")

    # Step 1
    job_posting, resume_text, profile = step1_read_inputs(job_posting, config)
    static_sections = step1b_extract_static_sections(resume_text, config)

    # Step 2
    analysis = step2_analyze(job_posting, resume_text, profile, company, role, contact, config)

    if config.dry_run:
        config.progress("\n  [dry-run] Skipping file generation — analysis complete.")
        return WorkflowResult(
            run_dir=run_dir,
            resume_path=resume_out,
            ats_path=ats_out,
            cover_letter_path=cover_out,
            framing_angle=analysis.get("framing_angle", ""),
        )

    # Step 2b
    colors = step2b_brand_colors(company, config)

    # Steps 3–5: styled resume
    step3_unpack(config)
    step4_apply_edits(analysis, resume_text, colors, config)
    step5_pack(resume_out, config)

    # Step 5b: ATS resume
    step5b_ats_resume(analysis, static_sections, company, role, ats_out, config)

    # Step 6: cover letter
    step6_cover_letter(analysis, company, role, cover_out, config, colors=colors)

    # Step 7: cleanup
    step7_cleanup(config)

    # Step 8: Drive upload
    folder_url = step8_upload(run_dir, company_safe, role_safe, config)

    return WorkflowResult(
        run_dir=run_dir,
        resume_path=resume_out,
        ats_path=ats_out,
        cover_letter_path=cover_out,
        framing_angle=analysis.get("framing_angle", ""),
        folder_url=folder_url,
    )

# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Interview Prep
# ---------------------------------------------------------------------------

PREP_SYSTEM = """\
You are an expert interview coach preparing Corey Laverdiere for a specific interview round.
You know his background deeply: integration engineering, AI/ML solutions delivery,
professional services, and customer-facing technical roles.

Your job is to produce a dense, specific interview reference card with 9 sections.
All content must be in Corey's voice: direct, specific, first-person, no corporate filler.
Every prepared answer must be specific enough that it couldn't apply to any other candidate.
Be ruthlessly specific — name tools, quote numbers, reference real projects.

Return ONLY valid JSON. No preamble, no markdown fences.
"""


def _build_prep_docx_js(
    data: dict,
    company: str,
    role: str,
    round_type: str,
    focus: str,
    interviewer: str,
    output_path: Path,
    colors: dict,
) -> str:
    """Return a Node.js script that produces the interview prep reference card DOCX."""

    # Color palette — navy + teal as primary, light fills for table rows
    NAVY   = "1F4E79"
    TEAL   = "00695C"
    FILL_B = "D6E4F0"   # blue row fill
    FILL_T = "D0ECEA"   # teal row fill
    FILL_W = "FFF9C4"   # warning/gap row fill
    FILL_H = "EBF5FB"   # header row fill (slightly darker blue)
    WHITE  = "FFFFFF"
    DARK   = "1A1A1A"

    def esc(text: str) -> str:
        return escape_js_string(" ".join(str(text).split()))

    def tr(text: str, bold: bool = False, italic: bool = False,
           size: int = 18, color: str = DARK) -> str:
        props = [f'text: "{esc(text)}"', 'font: "Arial"',
                 f'size: {size}', f'color: "{color}"']
        if bold:   props.append("bold: true")
        if italic: props.append("italic: true")
        return "new TextRun({ " + ", ".join(props) + " })"

    def para(children: list[str], before: int = 0, after: int = 60,
             left: int = 0, border_bottom_color: str = "") -> str:
        spacing = f"before: {before}, after: {after}"
        indent  = f", indent: {{ left: {left} }}" if left else ""
        border  = (
            f', border: {{ bottom: {{ style: BorderStyle.SINGLE, size: 4, '
            f'color: "{border_bottom_color}", space: 2 }} }}'
        ) if border_bottom_color else ""
        return (f'new Paragraph({{ spacing: {{ {spacing} }}{indent}{border}, '
                f'children: [{", ".join(children)}] }})')

    def section_header(title: str, color: str = NAVY) -> str:
        return para([tr(title, bold=True, size=19, color=color)],
                    before=100, after=50, border_bottom_color=color)

    def cell(children_paras: list[str], fill: str = WHITE,
             width: int = 4680, top_border: bool = False) -> str:
        borders_inner = (
            '{ style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }'
        )
        top = f'top: {borders_inner}, ' if top_border else f'top: {borders_inner}, '
        borders = (
            f'top: {borders_inner}, bottom: {borders_inner}, '
            f'left: {borders_inner}, right: {borders_inner}'
        )
        children_str = ",\n          ".join(children_paras)
        return (
            f'new TableCell({{\n'
            f'  width: {{ size: {width}, type: WidthType.DXA }},\n'
            f'  borders: {{ {borders} }},\n'
            f'  shading: {{ fill: "{fill}", type: ShadingType.CLEAR }},\n'
            f'  margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},\n'
            f'  children: [\n          {children_str}\n        ]\n'
            f'}})'
        )

    def row(cells: list[str]) -> str:
        cells_str = ",\n      ".join(cells)
        return f'new TableRow({{ children: [\n      {cells_str}\n    ] }})'

    def table(rows: list[str], col_widths: list[int]) -> str:
        total = sum(col_widths)
        widths_str = ", ".join(str(w) for w in col_widths)
        rows_str   = ",\n    ".join(rows)
        return (
            f'new Table({{\n'
            f'  width: {{ size: {total}, type: WidthType.DXA }},\n'
            f'  columnWidths: [{widths_str}],\n'
            f'  rows: [\n    {rows_str}\n  ]\n'
            f'}})'
        )

    # -------------------------------------------------------------------------
    # Page geometry (US Letter, 0.6" margins)
    # Content width = 12240 - 2*(0.6*1440) = 12240 - 1728 = 10512 DXA
    # Two main columns: ~50/50 split with small gutter
    # -------------------------------------------------------------------------
    MARGIN   = 864   # 0.6 inches
    CONTENT  = 10512
    COL_L    = 5160
    COL_R    = 5352  # slightly wider right column
    assert COL_L + COL_R == CONTENT

    # =========================================================================
    # HEADER BAND — full-width title row
    # =========================================================================
    interviewer_label = interviewer or "Hiring Team"
    focus_label       = focus.strip() or ""
    ctx_parts         = [f"INTERVIEW PREP · {role} · {company} · {round_type}"]
    if focus_label:
        ctx_parts.append(f"Focus: {focus_label}")

    header_paras = [
        para([tr("COREY LAVERDIERE", bold=True, size=28, color=WHITE),
              tr(f"  ·  {APPLICANT_CONTACT_LINE}", size=16, color="D6E4F0")],
             before=60, after=40),
        para([tr(ctx_parts[0], bold=True, size=17, color="D6E4F0")],
             before=0, after=(30 if focus_label else 60)),
    ]
    if focus_label:
        header_paras.append(
            para([tr(f"Focus: {focus_label}", italic=True, size=16, color="AED6F1")],
                 before=0, after=60)
        )

    header_row = row([cell(header_paras, fill=NAVY, width=CONTENT)])
    header_tbl = table([header_row], [CONTENT])

    # =========================================================================
    # LEFT COLUMN CONTENT
    # =========================================================================
    left_paras: list[str] = []

    # Section 1 — Know Your Interviewer
    left_paras.append(section_header(f"1 · Know Your Interviewer — {interviewer_label}", NAVY))
    bullets = data.get("know_your_interviewer", [])
    for b in bullets:
        left_paras.append(
            para([tr(f"•  {b}", size=17)], before=40, after=40, left=200)
        )

    # Section 2 — Role Fit Map
    left_paras.append(section_header("2 · Role Fit Map", NAVY))
    fit_rows = data.get("role_fit_map", [])
    if fit_rows:
        # header row
        fit_table_rows = [
            row([
                cell([para([tr("They Want", bold=True, size=17, color=WHITE)], before=40, after=40)],
                     fill=NAVY, width=2400),
                cell([para([tr("I Have", bold=True, size=17, color=WHITE)], before=40, after=40)],
                     fill=NAVY, width=3360),
            ])
        ]
        for i, item in enumerate(fit_rows):
            fill = FILL_B if i % 2 == 0 else WHITE
            fit_table_rows.append(row([
                cell([para([tr(item.get("they_want", ""), size=16)], before=40, after=40)],
                     fill=fill, width=2400),
                cell([para([tr(item.get("i_have", ""), size=16)], before=40, after=40)],
                     fill=fill, width=3360),
            ]))
        left_paras.append(table(fit_table_rows, [2400, 3360]))

    # Section 3 — Gap Bridge
    left_paras.append(section_header("3 · Gap Bridge — Proactive Reframes", TEAL))
    gaps = data.get("gap_bridge", [])
    if gaps:
        gap_rows = [
            row([
                cell([para([tr("Gap", bold=True, size=17, color=WHITE)], before=40, after=40)],
                     fill=TEAL, width=1680),
                cell([para([tr("How I'll Address It (say this proactively)", bold=True, size=17, color=WHITE)],
                           before=40, after=40)], fill=TEAL, width=3480),
            ])
        ]
        for g in gaps:
            gap_rows.append(row([
                cell([para([tr(g.get("gap", ""), size=16)], before=40, after=40)],
                     fill=FILL_W, width=1680),
                cell([para([tr(g.get("reframe", ""), size=16)], before=40, after=40)],
                     fill=FILL_W, width=3480),
            ]))
        left_paras.append(table(gap_rows, [1680, 3480]))

    # =========================================================================
    # RIGHT COLUMN CONTENT
    # =========================================================================
    right_paras: list[str] = []

    # Section 4 — Development Framework
    right_paras.append(section_header("4 · My Development Framework in a Nutshell", TEAL))
    fw = data.get("framework_summary", {})
    short_ver = fw.get("short_version", "")
    if short_ver:
        right_paras.append(
            para([tr(f'“{short_ver}”', italic=True, size=17, color=NAVY)],
                 before=40, after=80, left=120)
        )

    steps = fw.get("steps", [])
    if steps:
        step_rows = [
            row([
                cell([para([tr("Step", bold=True, size=17, color=WHITE)], before=40, after=40)],
                     fill=TEAL, width=1520),
                cell([para([tr("What I Do + Proof Point", bold=True, size=17, color=WHITE)],
                           before=40, after=40)], fill=TEAL, width=3832),
            ])
        ]
        for i, s in enumerate(steps):
            fill = FILL_T if i % 2 == 0 else WHITE
            step_cell_paras = [
                para([tr(s.get("what", ""), size=16)], before=40, after=20),
                para([tr(f"Proof: {s.get('proof', '')}", italic=True, size=15,
                         color="555555")], before=0, after=40),
            ]
            step_rows.append(row([
                cell([para([tr(s.get("name", ""), bold=True, size=16)], before=40, after=40)],
                     fill=fill, width=1520),
                cell(step_cell_paras, fill=fill, width=3832),
            ]))
        right_paras.append(table(step_rows, [1520, 3832]))

    # Section 5 — Anchor Stories
    right_paras.append(section_header("5 · Anchor Stories (STAR-Ready)", NAVY))
    stories = data.get("anchor_stories", [])
    if stories:
        story_rows = [
            row([
                cell([para([tr("Story Name", bold=True, size=17, color=WHITE)], before=40, after=40)],
                     fill=NAVY, width=2200),
                cell([para([tr("Key Signal It Demonstrates", bold=True, size=17, color=WHITE)],
                           before=40, after=40)], fill=NAVY, width=3152),
            ])
        ]
        for i, s in enumerate(stories):
            fill = FILL_B if i % 2 == 0 else WHITE
            story_rows.append(row([
                cell([para([tr(s.get("story_name", ""), bold=True, size=16)], before=40, after=40)],
                     fill=fill, width=2200),
                cell([para([tr(s.get("key_signal", ""), size=16)], before=40, after=40)],
                     fill=fill, width=3152),
            ]))
        right_paras.append(table(story_rows, [2200, 3152]))

    # =========================================================================
    # MAIN 2-COLUMN TABLE
    # =========================================================================
    main_row  = row([
        cell(left_paras,  fill=WHITE, width=COL_L),
        cell(right_paras, fill=WHITE, width=COL_R),
    ])
    main_tbl  = table([main_row], [COL_L, COL_R])

    # =========================================================================
    # BOTTOM BAND
    # =========================================================================

    # Divider header
    band_hdr_row = row([
        cell([para([tr("QUICK REFERENCE · Questions · Edge · Closing",
                       bold=True, size=17, color=WHITE)],
                   before=60, after=60)], fill=NAVY, width=CONTENT)
    ])
    band_hdr_tbl = table([band_hdr_row], [CONTENT])

    # Section 6 — Likely Questions (full width, 2-col table inside)
    q_header = section_header("6 · Likely Questions + Quick Answers", NAVY)
    qs = data.get("likely_questions", [])
    q_cell_paras = [q_header]
    if qs:
        q_rows = [
            row([
                cell([para([tr("Question", bold=True, size=17, color=WHITE)], before=40, after=40)],
                     fill=NAVY, width=3800),
                cell([para([tr("Pre-Loaded Answer (2–3 sentences, then stop)",
                               bold=True, size=17, color=WHITE)], before=40, after=40)],
                     fill=NAVY, width=6712),
            ])
        ]
        for i, item in enumerate(qs):
            fill = FILL_B if i % 2 == 0 else WHITE
            q_rows.append(row([
                cell([para([tr(item.get("question", ""), size=16)], before=40, after=40)],
                     fill=fill, width=3800),
                cell([para([tr(item.get("answer", ""), size=16)], before=40, after=40)],
                     fill=fill, width=6712),
            ]))
        q_cell_paras.append(table(q_rows, [3800, 6712]))

    q_band_row = row([cell(q_cell_paras, fill=WHITE, width=CONTENT)])
    q_band_tbl = table([q_band_row], [CONTENT])

    # Sections 7 + 8 side by side, then 9 full width
    qta_paras: list[str] = [section_header("7 · Questions to Ask", TEAL)]
    for q in data.get("questions_to_ask", []):
        qta_paras.append(para([tr(f"•  {q}", size=16)], before=30, after=30, left=160))

    edge_paras: list[str] = [section_header("8 · My Differentiating Edge", NAVY)]
    for b in data.get("differentiating_edge", []):
        edge_paras.append(para([tr(f"•  {b}", size=16)], before=30, after=30, left=160))

    half = CONTENT // 2
    row_78 = row([
        cell(qta_paras,  fill=WHITE, width=half),
        cell(edge_paras, fill=WHITE, width=CONTENT - half),
    ])
    tbl_78 = table([row_78], [half, CONTENT - half])

    # Section 9 — Closing Line
    closing = data.get("closing_line", "")
    row_9   = row([
        cell([
            section_header(f"9 · Closing Line — ‘Why {company}?’ (deploy verbatim)", TEAL),
            para([tr(f'“{closing}”', italic=True, size=17, color=NAVY)],
                 before=60, after=60, left=120),
        ], fill=FILL_T, width=CONTENT)
    ])
    tbl_9 = table([row_9], [CONTENT])

    # =========================================================================
    # Assemble children list
    # =========================================================================
    children_js = f"""
      {header_tbl},
      new Paragraph({{ spacing: {{ before: 0, after: 0 }}, children: [] }}),
      {main_tbl},
      new Paragraph({{ spacing: {{ before: 60, after: 0 }}, children: [] }}),
      {band_hdr_tbl},
      {q_band_tbl},
      {tbl_78},
      {tbl_9}
    """

    out_path_str = str(output_path).replace("\\", "/")

    return f"""\
const {{
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  BorderStyle, WidthType, ShadingType
}} = require('docx');
const fs = require('fs');

const doc = new Document({{
  styles: {{ default: {{ document: {{ run: {{ font: "Arial", size: 18 }} }} }} }},
  sections: [{{
    properties: {{
      page: {{
        size: {{ width: 12240, height: 15840 }},
        margin: {{ top: {MARGIN}, right: {MARGIN}, bottom: {MARGIN}, left: {MARGIN} }}
      }}
    }},
    children: [
      {children_js}
    ]
  }}]
}});

Packer.toBuffer(doc).then(buffer => {{
  fs.writeFileSync('{out_path_str}', buffer);
  console.log('Interview prep reference card written.');
}});
"""


def generate_interview_prep(
    job_posting: str,
    company: str,
    role: str,
    config: InterviewPrepConfig,
) -> InterviewPrepResult:
    """
    Generate a tailored interview prep DOCX.

    Args:
        job_posting: Full text of the job posting.
        company:     Company name.
        role:        Role title.
        config:      InterviewPrepConfig — round type, focus, model, profile, resume path.

    Returns:
        InterviewPrepResult with the path to the generated DOCX.

    Raises:
        WorkflowError on any unrecoverable error.
    """
    wfc = WorkflowConfig(
        model=config.model,
        progress=config.progress,
        master_resume=config.master_resume,
        profile_text=config.profile_text,
        user_id=config.user_id,
        user_label=config.user_label,
    )

    OUTPUT_DIR.mkdir(exist_ok=True)
    company_safe = safe_filename(company)
    role_safe    = safe_filename(role)
    round_safe   = safe_filename(config.round_type.replace(" ", ""))
    if config.user_id:
        run_dir = OUTPUT_DIR / safe_filename(config.user_id) / f"{company_safe}_{role_safe}"
    else:
        run_dir = OUTPUT_DIR / f"{company_safe}_{role_safe}"
    run_dir.mkdir(parents=True, exist_ok=True)

    prep_out = run_dir / (
        f"InterviewPrep_{APPLICANT_NAME}_{company_safe}_{role_safe}_{round_safe}.docx"
    )

    config.progress(f"\n\U0001f4cb Interview Prep Generator")
    config.progress(f"   Company : {company}")
    config.progress(f"   Role    : {role}")
    config.progress(f"   Round   : {config.round_type}")
    if config.focus:
        config.progress(f"   Focus   : {config.focus}")

    # Step 1: Read inputs
    print_step(1, "Reading Inputs", wfc)
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise WorkflowError("ANTHROPIC_API_KEY environment variable not set")
    resume_text = extract_resume_text(wfc)
    profile     = wfc.profile_text if wfc.profile_text is not None else read_file(PROFILE_FILE)
    config.progress(
        f"  ✓ Inputs loaded "
        f"({len(resume_text)} chars resume, {len(profile)} chars profile)"
    )

    # Step 2: Generate content with Claude
    print_step(2, "Generating Interview Prep Content", wfc)

    focus_note   = config.focus or "General — cover the most likely topics for this round type"
    interviewer  = config.interviewer or "Hiring Team"
    prompt = f"""
Job Posting:
---
{job_posting}
---

Candidate Resume:
---
{resume_text[:6000]}
---

Profile & Voice Guide:
---
{profile}
---

Company: {company}
Role: {role}
Interviewer: {interviewer}
Interview Round: {config.round_type}
Focus / Slant: {focus_note}

Produce a JSON object with EXACTLY these keys (no extras, no omissions):
{{
  "know_your_interviewer": [
    "string — bullet on how to frame answers for this specific person's role/priorities. What do they care about? 1-2 sentences each. Exactly 4 bullets."
  ],
  "role_fit_map": [
    {{
      "they_want": "string — specific JD requirement (under 12 words)",
      "i_have": "string — Corey's specific matching experience (under 20 words, name tools/numbers)"
    }}
  ],
  "gap_bridge": [
    {{
      "gap": "string — the most likely experience gap this interviewer will notice (under 10 words)",
      "reframe": "string — the proactive reframe answer Corey should give (2-3 sentences, specific, uses certs/analogous tools/transferable patterns)"
    }}
  ],
  "framework_summary": {{
    "short_version": "string — 30-second version of Corey's development framework. First person. Under 60 words.",
    "steps": [
      {{
        "name": "string — step name (3-5 words)",
        "what": "string — what Corey does at this step (1-2 sentences)",
        "proof": "string — a real proof point from his resume (1 sentence, italicized in output)"
      }}
    ]
  }},
  "anchor_stories": [
    {{
      "story_name": "string — short name for the story (3-6 words)",
      "key_signal": "string — the competency or theme this demonstrates (under 15 words)"
    }}
  ],
  "likely_questions": [
    {{
      "question": "string — likely question from THIS interviewer in THIS round (under 25 words)",
      "answer": "string — tight 2-3 sentence pre-loaded answer, first person, specific"
    }}
  ],
  "questions_to_ask": [
    "string — sharp question tailored to this interviewer's level and priorities (under 30 words)"
  ],
  "differentiating_edge": [
    "string — bullet on what makes Corey the strongest candidate specifically for this role at this company (1 sentence, specific)"
  ],
  "closing_line": "string — a pre-built verbatim 'Why {company}?' answer Corey can deploy. 3-5 sentences. Specific to this company and role. First person."
}}

Constraints:
- know_your_interviewer: exactly 4 bullets calibrated to {interviewer} and round: {config.round_type}
- role_fit_map: exactly 6 rows, covering the most critical JD requirements
- gap_bridge: exactly 1-2 items (only real gaps, not invented ones)
- framework_summary.steps: exactly 5 steps matching Corey's development framework
- anchor_stories: exactly 5 stories drawn from actual resume content
- likely_questions: exactly 6 questions weighted toward {config.round_type} and focus: {focus_note}
- questions_to_ask: exactly 5 items calibrated for {interviewer} at {config.round_type} level
- differentiating_edge: exactly 5 bullets

Round-specific guidance for "{config.round_type}":
- Phone Screen: culture fit, career motivation, logistics, high-level experience. QTA: team structure, 90-day success, next steps.
- Hiring Manager: role vision, leadership alignment, team dynamics, growth. QTA: biggest current challenges, how success is measured, what the team needs now.
- Peer: collaboration style, day-to-day workflow, technical problem-solving. QTA: team dynamics, tooling, what they wish they'd known before joining.
- Technical: system design, architecture tradeoffs, specific technical depth. QTA: stack decisions, engineering culture, biggest technical challenges.
- Executive: strategic impact, ROI, company direction, big-picture fit. QTA: company priorities, how AI/automation fits the roadmap, 3-year bet.
- Panel: multiple angles — mix role-fit, technical, and cultural questions.

Return ONLY valid JSON. No preamble, no markdown fences.
"""

    raw = claude(PREP_SYSTEM, prompt, max_tokens=4096, config=wfc)
    raw = re.sub(r"^```json\s*", "", raw.strip())
    raw = re.sub(r"\s*```$",     "", raw.strip())

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise WorkflowError(f"Failed to parse prep JSON: {e}\n\nRaw:\n{raw[:2000]}")

    config.progress(
        f"  ✓ Generated: {len(data.get('likely_questions', []))} questions, "
        f"{len(data.get('anchor_stories', []))} stories, "
        f"{len(data.get('role_fit_map', []))} role-fit rows"
    )

    # Step 2b: Brand colors
    print_step("2b", "Fetching Brand Colors", wfc)
    colors = get_brand_color(company)

    # Step 3: Build DOCX
    print_step(3, "Building Interview Prep DOCX", wfc)
    js      = _build_prep_docx_js(
        data, company, role, config.round_type, config.focus,
        config.interviewer, prep_out, colors
    )
    js_path = Path(f"interview_prep_gen_{os.urandom(4).hex()}.js")
    write_file(js_path, js)
    result  = run(["node", str(js_path)], check=False, config=wfc)
    js_path.unlink(missing_ok=True)
    if result.returncode != 0:
        raise WorkflowError(f"Interview prep JS failed:\n{result.stderr}")
    config.progress(f"  ✓ Interview prep written to {prep_out}")

    # Step 4: Upload to Drive
    print_step(8, "Uploading to Google Drive", wfc)
    folder_url = _upload_single_to_drive(prep_out, f"{company_safe}_{role_safe}", wfc)

    return InterviewPrepResult(
        prep_path=prep_out,
        run_dir=run_dir,
        folder_url=folder_url,
    )


def _print_result(result: WorkflowResult):
    print(f"\n{'='*60}")
    print(f"  DONE")
    print(f"{'='*60}")
    print(f"\n  \U0001f4c1 Output folder : {result.run_dir}")
    print(f"  \U0001f4c4 Resume (DOCX) : {result.resume_path.name}")
    print(f"  \U0001f916 ATS Resume    : {result.ats_path.name}")
    print(f"  \U0001f4dd Cover letter  : {result.cover_letter_path.name}")
    if result.folder_url:
        print(f"  ☁️  Drive folder  : {result.folder_url}")
    print(f"\n  Framing angle used:")
    print(f"  {textwrap.fill(result.framing_angle, width=56, initial_indent='  ', subsequent_indent='  ')}")
    print()


def main():
    parser = argparse.ArgumentParser(description="Job Application Agent - Corey Laverdiere")
    parser.add_argument("--job",      required=True,        help="Path to job posting text file")
    parser.add_argument("--company",  required=True,        help="Company name (used in filenames)")
    parser.add_argument("--role",     required=True,        help="Role title (used in filenames and cover letter)")
    parser.add_argument("--contact",  default=None,         help="Hiring manager name if known")
    parser.add_argument("--model",    default=DEFAULT_MODEL, help=f"Anthropic model (default: {DEFAULT_MODEL})")
    parser.add_argument("--debug",    action="store_true",  help="Keep unpacked/ and gen scripts for inspection")
    parser.add_argument("--dry-run",  action="store_true",  help="Run analysis only; skip file generation")
    args = parser.parse_args()

    job_path = Path(args.job)
    if not job_path.exists():
        print(f"❌ Job file not found at {job_path}")
        sys.exit(1)

    config = WorkflowConfig(
        model=args.model,
        debug=args.debug,
        dry_run=args.dry_run,
    )

    try:
        result = run_workflow(
            job_posting=job_path.read_text(encoding="utf-8"),
            company=args.company,
            role=args.role,
            contact=args.contact,
            config=config,
        )
        _print_result(result)
    except WorkflowError as e:
        print(f"\n❌ {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
