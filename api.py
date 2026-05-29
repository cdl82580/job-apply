"""
api.py — FastAPI backend for the Job Application Agent.

Auth:    Session cookie (HMAC-signed JWT-style token, stateless — works across machines).
Storage: Tigris S3 for user accounts, resumes, and profiles (see scripts/storage.py).

Public endpoints (no session required):
  POST /api/auth/register   Create account + upload resume + profile
  POST /api/auth/login      Returns session cookie
  GET  /api/health

Protected endpoints:
  POST /api/auth/logout
  GET  /api/auth/me
  GET  /api/profile
  PUT  /api/profile
  POST /api/profile/resume
  POST /api/profile/password
  POST /api/run
  GET  /api/run/{id}/stream
  GET  /api/run/{id}/status
  GET  /api/run/{id}/files/{name}
"""

from __future__ import annotations

import asyncio
import base64
import hashlib
import hmac
import json
import os
import secrets
import tempfile
import threading
import time
import urllib.request
import uuid
from pathlib import Path
from queue import Empty, Queue
from typing import Any

from fastapi import FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from scripts import storage
from apply import (
    DEFAULT_MODEL,
    MASTER_RESUME,
    OUTPUT_DIR,
    PROFILE_FILE,
    ROUND_TYPES,
    InterviewPrepConfig,
    InterviewPrepResult,
    WorkflowConfig,
    WorkflowError,
    WorkflowResult,
    generate_interview_prep,
    get_gdrive_job_posting,
    list_gdrive_run_folders,
    run_workflow,
    safe_filename,
)

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

FLY_MACHINE_ID = os.environ.get("FLY_MACHINE_ID", "")
FLY_APP_NAME   = os.environ.get("FLY_APP_NAME", "job-apply-corey")

# Signing secret for session tokens.  Generated fresh on each deploy if unset
# (means all existing sessions are invalidated on restart) — set SESSION_SECRET
# as a Fly.io secret to get persistent sessions across restarts.
_SESSION_SECRET  = os.environ.get("SESSION_SECRET", secrets.token_hex(32))
_SESSION_COOKIE  = "session"
_SESSION_DAYS    = 30
# Bearer token for the Slack bot — set BOT_API_KEY as a Fly.io secret.
# Requests carrying this token skip cookie auth and run as the primary user account.
_BOT_API_KEY     = os.environ.get("BOT_API_KEY", "")

_NOTIFY_EMAIL = os.environ.get("APP_USER_EMAIL", "cdl825@gmail.com")

# ---------------------------------------------------------------------------
# Session helpers (stateless HMAC — works across both Fly.io machines)
# ---------------------------------------------------------------------------

def _create_session(user_id: str, email: str) -> str:
    payload = base64.urlsafe_b64encode(json.dumps({
        "user_id": user_id,
        "email":   email,
        "exp":     int(time.time()) + 86400 * _SESSION_DAYS,
    }).encode()).rstrip(b"=").decode()
    sig = hmac.new(_SESSION_SECRET.encode(), payload.encode(), hashlib.sha256).hexdigest()
    return f"{payload}.{sig}"


def _verify_session(token: str) -> dict | None:
    try:
        payload_b64, sig = token.rsplit(".", 1)
        expected = hmac.new(_SESSION_SECRET.encode(), payload_b64.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected):
            return None
        # Restore padding for base64 decoding
        padding = 4 - len(payload_b64) % 4
        data = json.loads(base64.urlsafe_b64decode(payload_b64 + "=" * padding))
        if data.get("exp", 0) < time.time():
            return None
        return data
    except Exception:
        return None


def _bot_user(request: Request) -> dict | None:
    """Return a synthetic user dict if the request carries a valid bot API key."""
    if not _BOT_API_KEY:
        return None
    auth = request.headers.get("Authorization", "")
    if not (auth.startswith("Bearer ") and hmac.compare_digest(auth[7:], _BOT_API_KEY)):
        return None
    # Resolve the primary user account so the bot runs as a real user with
    # a real resume and profile stored in Tigris.
    primary = storage.get_user_by_email(_NOTIFY_EMAIL)
    if not primary:
        return None
    return {"user_id": primary["user_id"], "email": primary["email"]}


def _current_user(request: Request) -> dict | None:
    bot = _bot_user(request)
    if bot:
        return bot
    token = request.cookies.get(_SESSION_COOKIE, "")
    return _verify_session(token)


def _require_user(request: Request) -> dict:
    user = _current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    return user

# ---------------------------------------------------------------------------
# Password hashing (stdlib scrypt — no extra deps)
# ---------------------------------------------------------------------------

def _hash_password(password: str) -> str:
    salt = os.urandom(16)
    dk   = hashlib.scrypt(password.encode(), salt=salt, n=16384, r=8, p=1)
    return f"scrypt:{salt.hex()}:{dk.hex()}"


def _verify_password(password: str, stored: str) -> bool:
    try:
        _, salt_hex, dk_hex = stored.split(":")
        dk = hashlib.scrypt(password.encode(), salt=bytes.fromhex(salt_hex), n=16384, r=8, p=1)
        return hmac.compare_digest(dk.hex(), dk_hex)
    except Exception:
        return False

# ---------------------------------------------------------------------------
# Email helper (Resend)
# ---------------------------------------------------------------------------

def _send_email(to: str, subject: str, body: str) -> bool:
    api_key = os.environ.get("RESEND_API_KEY", "")
    if not api_key:
        return False
    payload = json.dumps({
        "from":    "Job Apply <onboarding@resend.dev>",
        "to":      [to],
        "subject": subject,
        "text":    body,
    }).encode()
    req = urllib.request.Request(
        "https://api.resend.com/emails",
        data=payload,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            return 200 <= resp.status < 300
    except Exception:
        return False

# ---------------------------------------------------------------------------
# App + auth middleware
# ---------------------------------------------------------------------------

app = FastAPI(title="Job Application Agent")

_PUBLIC_PATHS = frozenset({
    "/login.html", "/register.html",
    "/api/auth/login", "/api/auth/register",
    "/api/health",
    "/favicon.ico",
})


@app.middleware("http")
async def auth_middleware(request: Request, call_next):
    if request.url.path in _PUBLIC_PATHS:
        return await call_next(request)

    user = _current_user(request)
    if not user:
        if request.url.path.startswith("/api/"):
            return JSONResponse({"detail": "Not authenticated"}, status_code=401)
        return RedirectResponse("/login.html", status_code=302)

    request.state.user = user
    return await call_next(request)

# ---------------------------------------------------------------------------
# In-memory stores
# ---------------------------------------------------------------------------

_runs:  dict[str, dict[str, Any]] = {}
_preps: dict[str, dict[str, Any]] = {}
_workflow_lock = threading.Lock()

# ---------------------------------------------------------------------------
# Request models
# ---------------------------------------------------------------------------

class LoginRequest(BaseModel):
    email: str
    password: str

class PasswordChangeRequest(BaseModel):
    current_password: str
    new_password: str

class ProfileUpdateRequest(BaseModel):
    display_name: str | None = None
    profile_text: str | None = None

class RunRequest(BaseModel):
    job_posting: str
    company: str
    role: str
    contact: str | None = None
    model: str | None = None

class PrepRequest(BaseModel):
    job_posting: str
    company: str
    role: str
    round_type: str
    focus: str | None = None
    interviewer: str | None = None
    model: str | None = None

# ---------------------------------------------------------------------------
# Health
# ---------------------------------------------------------------------------

@app.get("/api/health")
async def health():
    return {"status": "ok", "storage": storage.is_configured()}

# ---------------------------------------------------------------------------
# Auth endpoints
# ---------------------------------------------------------------------------

@app.post("/api/auth/register")
async def register(
    display_name: str      = Form(...),
    email:        str      = Form(...),
    password:     str      = Form(...),
    profile_text: str      = Form(...),
    resume:       UploadFile = File(...),
):
    email = email.strip().lower()

    if storage.get_user_by_email(email):
        raise HTTPException(400, "An account with that email already exists.")
    if len(password) < 8:
        raise HTTPException(400, "Password must be at least 8 characters.")
    if not resume.filename.lower().endswith(".docx"):
        raise HTTPException(400, "Resume must be a .docx file.")

    resume_data = await resume.read()
    if len(resume_data) < 1000:
        raise HTTPException(400, "Resume file appears to be empty or invalid.")

    user_id = str(uuid.uuid4())
    user = {
        "user_id":         user_id,
        "email":           email,
        "display_name":    display_name.strip(),
        "password_hash":   _hash_password(password),
        "created_at":      time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "resume_filename": resume.filename,
    }

    storage.save_user(user)
    storage.save_resume(user_id, resume_data)
    storage.save_profile(user_id, profile_text)

    response = JSONResponse({"ok": True, "display_name": user["display_name"]})
    token = _create_session(user_id, email)
    response.set_cookie(_SESSION_COOKIE, token, max_age=86400 * _SESSION_DAYS,
                        httponly=True, samesite="lax", secure=True)
    if FLY_MACHINE_ID:
        response.set_cookie("fly-force-instance-id", FLY_MACHINE_ID, path="/", samesite="lax")
    return response


@app.post("/api/auth/login")
async def login(req: LoginRequest):
    email = req.email.strip().lower()
    user  = storage.get_user_by_email(email)

    # Constant-time check even on miss to prevent user enumeration timing
    dummy_hash = "scrypt:00000000000000000000000000000000:00000000000000000000000000000000"
    stored = user["password_hash"] if user else dummy_hash
    ok = _verify_password(req.password, stored) and user is not None

    if not ok:
        raise HTTPException(401, "Incorrect email or password.")

    response = JSONResponse({"ok": True, "display_name": user["display_name"]})
    token = _create_session(user["user_id"], email)
    response.set_cookie(_SESSION_COOKIE, token, max_age=86400 * _SESSION_DAYS,
                        httponly=True, samesite="lax", secure=True)
    if FLY_MACHINE_ID:
        response.set_cookie("fly-force-instance-id", FLY_MACHINE_ID, path="/", samesite="lax")
    return response


@app.post("/api/auth/logout")
async def logout():
    response = JSONResponse({"ok": True})
    response.delete_cookie(_SESSION_COOKIE)
    response.delete_cookie("fly-force-instance-id")
    return response


@app.get("/api/auth/me")
async def me(request: Request):
    user_data = _require_user(request)
    record = storage.get_user_by_id(user_data["user_id"]) or {}
    return {
        "user_id":      user_data["user_id"],
        "email":        user_data["email"],
        "display_name": record.get("display_name", user_data["email"]),
        "has_resume":   storage.has_resume(user_data["user_id"]),
        "has_profile":  bool(storage.get_profile(user_data["user_id"])),
    }

# ---------------------------------------------------------------------------
# Profile endpoints
# ---------------------------------------------------------------------------

@app.get("/api/profile")
async def get_profile(request: Request):
    user_data = _require_user(request)
    record = storage.get_user_by_id(user_data["user_id"]) or {}
    profile_text = storage.get_profile(user_data["user_id"]) or ""
    return {
        "display_name":    record.get("display_name", ""),
        "email":           user_data["email"],
        "profile_text":    profile_text,
        "has_resume":      storage.has_resume(user_data["user_id"]),
        "resume_filename": record.get("resume_filename"),
    }


@app.put("/api/profile")
async def update_profile(req: ProfileUpdateRequest, request: Request):
    user_data = _require_user(request)
    record = storage.get_user_by_id(user_data["user_id"])
    if not record:
        raise HTTPException(404, "User not found.")

    if req.display_name is not None:
        record["display_name"] = req.display_name.strip()
        storage.save_user(record)

    if req.profile_text is not None:
        storage.save_profile(user_data["user_id"], req.profile_text)

    return {"ok": True}


@app.post("/api/profile/resume")
async def upload_resume(request: Request, resume: UploadFile = File(...)):
    user_data = _require_user(request)
    if not resume.filename.lower().endswith(".docx"):
        raise HTTPException(400, "Resume must be a .docx file.")
    data = await resume.read()
    if len(data) < 1000:
        raise HTTPException(400, "File appears empty or invalid.")
    record = storage.get_user_by_id(user_data["user_id"])
    if record:
        record["resume_filename"] = resume.filename
        storage.save_user(record)
    storage.save_resume(user_data["user_id"], data)
    return {"ok": True}


@app.post("/api/profile/password")
async def change_password(req: PasswordChangeRequest, request: Request):
    user_data = _require_user(request)
    record = storage.get_user_by_id(user_data["user_id"])
    if not record:
        raise HTTPException(404, "User not found.")
    if not _verify_password(req.current_password, record["password_hash"]):
        raise HTTPException(401, "Current password is incorrect.")
    if len(req.new_password) < 8:
        raise HTTPException(400, "New password must be at least 8 characters.")

    record["password_hash"] = _hash_password(req.new_password)
    storage.save_user(record)

    emailed = _send_email(
        to=user_data["email"],
        subject="Job Apply — Password Changed",
        body=(
            f"Your password was just changed.\n\n"
            f"Log in at https://{FLY_APP_NAME}.fly.dev/"
        ),
    )
    return {"ok": True, "emailed": emailed}

# ---------------------------------------------------------------------------
# Run endpoints
# ---------------------------------------------------------------------------

@app.post("/api/run")
async def create_run(req: RunRequest, request: Request, response: Response):
    user_data = _require_user(request)
    user_id   = user_data["user_id"]

    # Fetch user's resume and profile from Tigris
    resume_bytes = storage.get_resume(user_id)
    if not resume_bytes:
        raise HTTPException(400, "No master resume uploaded. Add one in your profile.")

    profile_text = storage.get_profile(user_id)
    if not profile_text:
        raise HTTPException(400, "No profile guide saved. Add one in your profile.")

    run_id = str(uuid.uuid4())
    q: Queue[dict | None] = Queue()
    _runs[run_id] = {"queue": q, "status": "queued", "result": None, "error": None,
                     "user_id": user_id}

    # Pin this browser session to the machine that owns this run's state
    if FLY_MACHINE_ID:
        response.set_cookie("fly-force-instance-id", FLY_MACHINE_ID, path="/", samesite="lax")

    def _thread():
        # Write resume to a temp file (pandoc + unpack need a real path)
        tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False, dir="/tmp")
        tmp.write(resume_bytes)
        tmp.close()
        resume_path = Path(tmp.name)

        try:
            with _workflow_lock:
                _runs[run_id]["status"] = "running"

                def progress(msg: str):
                    q.put({"type": "progress", "message": msg})

                config = WorkflowConfig(
                    model=req.model or DEFAULT_MODEL,
                    progress=progress,
                    master_resume=resume_path,
                    profile_text=profile_text,
                    user_id=user_id,
                    user_label=user_data["email"],
                )
                try:
                    result: WorkflowResult = run_workflow(
                        job_posting=req.job_posting,
                        company=req.company,
                        role=req.role,
                        contact=req.contact,
                        config=config,
                    )
                    _runs[run_id]["result"] = result
                    _runs[run_id]["status"] = "done"
                    q.put({
                        "type":          "done",
                        "run_id":        run_id,
                        "framing_angle": result.framing_angle,
                        "folder_url":    result.folder_url,
                        "files": {
                            "resume":       result.resume_path.name,
                            "ats":          result.ats_path.name,
                            "cover_letter": result.cover_letter_path.name,
                        },
                    })
                except WorkflowError as exc:
                    _runs[run_id]["status"] = "error"
                    _runs[run_id]["error"]  = str(exc)
                    q.put({"type": "error", "message": str(exc)})
                except Exception as exc:
                    msg = f"Unexpected error: {type(exc).__name__}: {exc}"
                    _runs[run_id]["status"] = "error"
                    _runs[run_id]["error"]  = msg
                    q.put({"type": "error", "message": msg})
                finally:
                    q.put(None)
        finally:
            resume_path.unlink(missing_ok=True)

    threading.Thread(target=_thread, daemon=True).start()
    return {"run_id": run_id}


@app.get("/api/run/{run_id}/stream")
async def stream_run(run_id: str, request: Request):
    _require_user(request)
    if run_id not in _runs:
        raise HTTPException(404, "Run not found")

    q    = _runs[run_id]["queue"]
    loop = asyncio.get_event_loop()

    async def generate():
        while True:
            try:
                msg = await loop.run_in_executor(None, lambda: q.get(timeout=30))
            except Empty:
                yield ": keepalive\n\n"
                continue
            if msg is None:
                break
            yield f"data: {json.dumps(msg)}\n\n"

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/api/run/{run_id}/status")
async def run_status(run_id: str, request: Request):
    _require_user(request)
    run = _runs.get(run_id)
    if not run:
        raise HTTPException(404, "Run not found")
    return {"run_id": run_id, "status": run["status"], "error": run.get("error")}


@app.get("/api/run/{run_id}/files/{filename}")
async def get_file(run_id: str, filename: str, request: Request):
    user_data = _require_user(request)
    run = _runs.get(run_id)
    if not run or run["status"] != "done" or not run.get("result"):
        raise HTTPException(404, "Run not complete")
    if run.get("user_id") != user_data["user_id"]:
        raise HTTPException(403, "Access denied")

    result: WorkflowResult = run["result"]
    file_path = (result.run_dir / filename).resolve()
    try:
        file_path.relative_to(result.run_dir.resolve())
    except ValueError:
        raise HTTPException(400, "Invalid filename")

    if not file_path.exists():
        raise HTTPException(404, "File not found")

    return FileResponse(
        file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ---------------------------------------------------------------------------
# Google Drive — run listing + job posting fetch
# ---------------------------------------------------------------------------

@app.get("/api/gdrive/runs")
async def gdrive_list_runs(request: Request):
    """List this user's run folders from Google Drive (user subfolder + legacy flat)."""
    user_data  = _require_user(request)
    user_label = user_data["email"]
    config = WorkflowConfig(progress=lambda _: None, user_label=user_label)
    try:
        folders = list_gdrive_run_folders(user_label, config)
        return {"runs": folders, "drive_configured": len(folders) >= 0}
    except Exception as exc:
        return {"runs": [], "drive_configured": False, "error": str(exc)}


@app.get("/api/gdrive/runs/{folder_id}/job_posting")
async def gdrive_get_job_posting(folder_id: str, request: Request):
    """Fetch job_posting.txt from a specific Drive folder (by Drive folder ID)."""
    _require_user(request)
    config = WorkflowConfig(progress=lambda _: None)
    text = get_gdrive_job_posting(folder_id, config)
    if text is None:
        raise HTTPException(404, "No job posting found in this Drive folder")
    return {"job_posting": text}


# ---------------------------------------------------------------------------
# Run listing (for interview prep dropdown)
# ---------------------------------------------------------------------------

@app.get("/api/runs")
async def list_runs(request: Request):
    user_data = _require_user(request)
    user_dir  = OUTPUT_DIR / safe_filename(user_data["user_id"])
    runs = []
    if user_dir.exists():
        dirs = sorted(user_dir.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True)
        for d in dirs:
            if d.is_dir():
                runs.append({
                    "folder":          d.name,
                    "has_job_posting": (d / "job_posting.txt").exists(),
                })
    return {"runs": runs}


@app.get("/api/runs/{folder}/job_posting")
async def get_run_job_posting(folder: str, request: Request):
    user_data = _require_user(request)
    user_dir  = OUTPUT_DIR / safe_filename(user_data["user_id"])
    try:
        path = (user_dir / folder / "job_posting.txt").resolve()
        path.relative_to(user_dir.resolve())
    except ValueError:
        raise HTTPException(400, "Invalid folder")
    if not path.exists():
        raise HTTPException(404, "Job posting not saved for this run")
    return {"job_posting": path.read_text(encoding="utf-8")}


# ---------------------------------------------------------------------------
# Interview prep endpoints
# ---------------------------------------------------------------------------

@app.post("/api/prep")
async def create_prep(req: PrepRequest, request: Request, response: Response):
    user_data = _require_user(request)
    user_id   = user_data["user_id"]

    if req.round_type not in ROUND_TYPES:
        raise HTTPException(400, f"round_type must be one of: {', '.join(ROUND_TYPES)}")

    resume_bytes = storage.get_resume(user_id)
    if not resume_bytes:
        raise HTTPException(400, "No master resume uploaded. Add one in your profile.")

    profile_text = storage.get_profile(user_id)
    if not profile_text:
        raise HTTPException(400, "No profile guide saved. Add one in your profile.")

    prep_id = str(uuid.uuid4())
    q: Queue[dict | None] = Queue()
    _preps[prep_id] = {"queue": q, "status": "queued", "result": None, "error": None,
                       "user_id": user_id}

    if FLY_MACHINE_ID:
        response.set_cookie("fly-force-instance-id", FLY_MACHINE_ID, path="/", samesite="lax")

    def _thread():
        tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False, dir="/tmp")
        tmp.write(resume_bytes)
        tmp.close()
        resume_path = Path(tmp.name)

        try:
            _preps[prep_id]["status"] = "running"

            def progress(msg: str):
                q.put({"type": "progress", "message": msg})

            config = InterviewPrepConfig(
                round_type=req.round_type,
                focus=req.focus or "",
                interviewer=req.interviewer or "",
                model=req.model or DEFAULT_MODEL,
                progress=progress,
                master_resume=resume_path,
                profile_text=profile_text,
                user_id=user_id,
                user_label=user_data["email"],
            )
            try:
                result: InterviewPrepResult = generate_interview_prep(
                    job_posting=req.job_posting,
                    company=req.company,
                    role=req.role,
                    config=config,
                )
                _preps[prep_id]["result"] = result
                _preps[prep_id]["status"] = "done"
                q.put({
                    "type":       "done",
                    "prep_id":    prep_id,
                    "folder_url": result.folder_url,
                    "files": {
                        "prep": result.prep_path.name,
                    },
                })
            except WorkflowError as exc:
                _preps[prep_id]["status"] = "error"
                _preps[prep_id]["error"]  = str(exc)
                q.put({"type": "error", "message": str(exc)})
            except Exception as exc:
                msg = f"Unexpected error: {type(exc).__name__}: {exc}"
                _preps[prep_id]["status"] = "error"
                _preps[prep_id]["error"]  = msg
                q.put({"type": "error", "message": msg})
            finally:
                q.put(None)
        finally:
            resume_path.unlink(missing_ok=True)

    threading.Thread(target=_thread, daemon=True).start()
    return {"prep_id": prep_id}


@app.get("/api/prep/{prep_id}/stream")
async def stream_prep(prep_id: str, request: Request):
    _require_user(request)
    if prep_id not in _preps:
        raise HTTPException(404, "Prep run not found")

    q    = _preps[prep_id]["queue"]
    loop = asyncio.get_event_loop()

    async def generate():
        while True:
            try:
                msg = await loop.run_in_executor(None, lambda: q.get(timeout=30))
            except Empty:
                yield ": keepalive\n\n"
                continue
            if msg is None:
                break
            yield f"data: {json.dumps(msg)}\n\n"

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/api/prep/{prep_id}/status")
async def prep_status(prep_id: str, request: Request):
    _require_user(request)
    prep = _preps.get(prep_id)
    if not prep:
        raise HTTPException(404, "Prep run not found")
    return {"prep_id": prep_id, "status": prep["status"], "error": prep.get("error")}


@app.get("/api/prep/{prep_id}/files/{filename}")
async def get_prep_file(prep_id: str, filename: str, request: Request):
    user_data = _require_user(request)
    prep = _preps.get(prep_id)
    if not prep or prep["status"] != "done" or not prep.get("result"):
        raise HTTPException(404, "Prep not complete")
    if prep.get("user_id") != user_data["user_id"]:
        raise HTTPException(403, "Access denied")

    result: InterviewPrepResult = prep["result"]
    file_path = (result.run_dir / filename).resolve()
    try:
        file_path.relative_to(result.run_dir.resolve())
    except ValueError:
        raise HTTPException(400, "Invalid filename")

    if not file_path.exists():
        raise HTTPException(404, "File not found")

    return FileResponse(
        file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ---------------------------------------------------------------------------
# Static frontend — mounted last; auth middleware handles redirects
# ---------------------------------------------------------------------------
app.mount("/", StaticFiles(directory="frontend", html=True), name="frontend")
