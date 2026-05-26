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
import json
import os
import threading
import uuid
from pathlib import Path
from queue import Empty, Queue
from typing import Any

from fastapi import FastAPI, HTTPException, Response
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

# When running on Fly.io each machine has a unique ID in this env var.
# We set it as a cookie on POST /api/run so Fly.io's proxy pins all follow-up
# requests (status, stream, file downloads) to the same machine — otherwise the
# load balancer can route them to a different machine that has no record of the run.
FLY_MACHINE_ID = os.environ.get("FLY_MACHINE_ID", "")

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

# ---------------------------------------------------------------------------
# In-memory run store (sufficient for a single-user tool)
# ---------------------------------------------------------------------------
# Each entry: {queue, status, result, error}
_runs: dict[str, dict[str, Any]] = {}

# One workflow at a time — the workflow relies on a shared UNPACK_DIR
_workflow_lock = threading.Lock()

# ---------------------------------------------------------------------------
# Request model
# ---------------------------------------------------------------------------

class RunRequest(BaseModel):
    job_posting: str
    company: str
    role: str
    contact: str | None = None
    model: str | None = None

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
