"""
api.py — FastAPI backend for the Job Application Agent.

Endpoints:
  POST   /api/run                          Start a workflow run (returns run_id)
  GET    /api/run/{run_id}/stream          SSE stream of progress events
  GET    /api/run/{run_id}/status          Lightweight status poll
  GET    /api/run/{run_id}/files/{name}    Download a generated file
  GET    /api/health                       Health / readiness check

Workflow runs are serialized by a threading.Lock — the workflow writes to a
shared UNPACK_DIR, so concurrent runs would conflict. For a single-user tool
this is fine; a second submission queues behind the first.
"""

from __future__ import annotations

import asyncio
import base64
import json
import os
import secrets
import threading
import urllib.request
import uuid
from pathlib import Path
from queue import Empty, Queue
from typing import Any

from fastapi import FastAPI, HTTPException, Request, Response
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

FLY_MACHINE_ID = os.environ.get("FLY_MACHINE_ID", "")
FLY_APP_NAME   = os.environ.get("FLY_APP_NAME", "job-apply-corey")

# HTTP Basic Auth — set APP_PASSWORD env var / Fly.io secret to enable.
# Username is ignored; any value works. /api/health is always exempt.
_APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

# Email address for password-change notifications.
_NOTIFY_EMAIL = os.environ.get("APP_USER_EMAIL", "cdl825@gmail.com")

from apply import (
    DEFAULT_MODEL,
    MASTER_RESUME,
    PROFILE_FILE,
    WorkflowConfig,
    WorkflowError,
    WorkflowResult,
    run_workflow,
)

app = FastAPI(title="Job Application Agent")


@app.middleware("http")
async def basic_auth(request: Request, call_next):
    """Require HTTP Basic Auth on every request except /api/health."""
    if not _APP_PASSWORD or request.url.path == "/api/health":
        return await call_next(request)

    auth = request.headers.get("Authorization", "")
    authenticated = False
    if auth.startswith("Basic "):
        try:
            _, pw = base64.b64decode(auth[6:]).decode().split(":", 1)
            authenticated = secrets.compare_digest(pw, _APP_PASSWORD)
        except Exception:
            pass

    if not authenticated:
        return Response(
            "Unauthorized",
            status_code=401,
            headers={"WWW-Authenticate": 'Basic realm="job-apply"'},
        )
    return await call_next(request)


# ---------------------------------------------------------------------------
# In-memory run store (sufficient for a single-user tool)
# ---------------------------------------------------------------------------
# Each entry: {queue, status, result, error}
_runs: dict[str, dict[str, Any]] = {}

# One workflow at a time — the workflow relies on a shared UNPACK_DIR
_workflow_lock = threading.Lock()

# ---------------------------------------------------------------------------
# Request models
# ---------------------------------------------------------------------------

class RunRequest(BaseModel):
    job_posting: str
    company: str
    role: str
    contact: str | None = None
    model: str | None = None

class PasswordChangeRequest(BaseModel):
    current_password: str
    new_password: str

# ---------------------------------------------------------------------------
# Helpers: email + Fly.io secret persistence
# ---------------------------------------------------------------------------

def _send_email(subject: str, body: str) -> bool:
    """Send a notification email via Resend. No-ops if RESEND_API_KEY is unset."""
    api_key = os.environ.get("RESEND_API_KEY", "")
    if not api_key:
        return False
    payload = json.dumps({
        "from":    "Job Apply <onboarding@resend.dev>",
        "to":      [_NOTIFY_EMAIL],
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


def _fly_persist_password(new_password: str) -> bool:
    """Push APP_PASSWORD to Fly.io secrets via GraphQL API.
    Requires FLY_API_TOKEN env var. Triggers a ~30s rolling restart."""
    token = os.environ.get("FLY_API_TOKEN", "")
    if not token:
        return False
    query = {
        "query": (
            "mutation ($input: SetSecretsInput!) {"
            "  setSecrets(input: $input) { release { id } } }"
        ),
        "variables": {
            "input": {
                "appId":   FLY_APP_NAME,
                "secrets": [{"key": "APP_PASSWORD", "value": new_password}],
            }
        },
    }
    req = urllib.request.Request(
        "https://api.fly.io/graphql",
        data=json.dumps(query).encode(),
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            result = json.loads(resp.read())
            return "errors" not in result
    except Exception:
        return False

# ---------------------------------------------------------------------------
# API endpoints
# ---------------------------------------------------------------------------

@app.get("/api/health")
async def health():
    return {
        "status": "ok",
        "master_resume": MASTER_RESUME.exists(),
        "profile": PROFILE_FILE.exists(),
    }


@app.post("/api/settings/password")
async def change_password(req: PasswordChangeRequest):
    global _APP_PASSWORD
    if not _APP_PASSWORD:
        raise HTTPException(status_code=400, detail="Password protection is not enabled.")
    if not secrets.compare_digest(req.current_password, _APP_PASSWORD):
        raise HTTPException(status_code=401, detail="Current password is incorrect.")
    if len(req.new_password) < 8:
        raise HTTPException(status_code=400, detail="New password must be at least 8 characters.")

    # Update in-memory immediately (all requests on this machine use it at once)
    _APP_PASSWORD = req.new_password
    os.environ["APP_PASSWORD"] = req.new_password

    # Persist to Fly.io secrets (triggers ~30s rolling restart on both machines)
    persisted = _fly_persist_password(req.new_password)

    # Email confirmation
    emailed = _send_email(
        subject="Job Apply — Password Changed",
        body=(
            f"Your Job Application Agent password was just changed.\n\n"
            f"New password:  {req.new_password}\n\n"
            f"Log in at https://{FLY_APP_NAME}.fly.dev/\n"
            + (
                "\nThe password has been saved — it will survive server restarts."
                if persisted else
                "\nNote: run the following to make it permanent across restarts:\n"
                f'  fly secrets set APP_PASSWORD="{req.new_password}" --app {FLY_APP_NAME}'
            )
        ),
    )

    return {
        "ok":        True,
        "persisted": persisted,
        "emailed":   emailed,
        "email_to":  _NOTIFY_EMAIL,
    }


@app.post("/api/run")
async def create_run(req: RunRequest, response: Response):
    run_id = str(uuid.uuid4())
    q: Queue[dict | None] = Queue()
    _runs[run_id] = {"queue": q, "status": "queued", "result": None, "error": None}

    # Pin the browser session to this machine so status/stream/file requests
    # don't land on a different Fly.io machine that has no memory of this run.
    if FLY_MACHINE_ID:
        response.set_cookie("fly-force-instance-id", FLY_MACHINE_ID, path="/", samesite="lax")

    def _thread():
        # Update status once we acquire the lock (another run may be in progress)
        with _workflow_lock:
            _runs[run_id]["status"] = "running"

            def progress(msg: str):
                q.put({"type": "progress", "message": msg})

            config = WorkflowConfig(
                model=req.model or DEFAULT_MODEL,
                progress=progress,
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
                    "type": "done",
                    "run_id": run_id,
                    "framing_angle": result.framing_angle,
                    "folder_url": result.folder_url,
                    "files": {
                        "resume": result.resume_path.name,
                        "ats": result.ats_path.name,
                        "cover_letter": result.cover_letter_path.name,
                    },
                })
            except WorkflowError as exc:
                _runs[run_id]["status"] = "error"
                _runs[run_id]["error"] = str(exc)
                q.put({"type": "error", "message": str(exc)})
            except Exception as exc:
                # Catch-all so an unexpected crash sends an error message rather
                # than silently closing the SSE stream (which shows "Connection lost").
                msg = f"Unexpected error: {type(exc).__name__}: {exc}"
                _runs[run_id]["status"] = "error"
                _runs[run_id]["error"] = msg
                q.put({"type": "error", "message": msg})
            finally:
                q.put(None)  # sentinel — signals end-of-stream

    threading.Thread(target=_thread, daemon=True).start()
    return {"run_id": run_id}


@app.get("/api/run/{run_id}/stream")
async def stream_run(run_id: str):
    if run_id not in _runs:
        raise HTTPException(status_code=404, detail="Run not found")

    q = _runs[run_id]["queue"]
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
async def run_status(run_id: str):
    run = _runs.get(run_id)
    if not run:
        raise HTTPException(status_code=404, detail="Run not found")
    return {"run_id": run_id, "status": run["status"], "error": run.get("error")}


@app.get("/api/run/{run_id}/files/{filename}")
async def get_file(run_id: str, filename: str):
    run = _runs.get(run_id)
    if not run or run["status"] != "done" or not run.get("result"):
        raise HTTPException(status_code=404, detail="Run not complete")

    result: WorkflowResult = run["result"]
    file_path = (result.run_dir / filename).resolve()

    # Prevent path traversal
    try:
        file_path.relative_to(result.run_dir.resolve())
    except ValueError:
        raise HTTPException(status_code=400, detail="Invalid filename")

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found")

    return FileResponse(
        file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ---------------------------------------------------------------------------
# Static frontend — must be mounted last so /api/* routes take precedence
# ---------------------------------------------------------------------------
app.mount("/", StaticFiles(directory="frontend", html=True), name="frontend")
