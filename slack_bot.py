"""
slack_bot.py — Slack bot for the Job Application Agent.

Environment variables required:
  SLACK_BOT_TOKEN       xoxb-... token from the Slack app
  SLACK_SIGNING_SECRET  signing secret from the Slack app Basic Information page
  BOT_API_KEY           must match the BOT_API_KEY set on the Fly.io app
  JOB_APPLY_API_URL     base URL of the deployed app (default: https://job-apply-corey.fly.dev)

Run locally:
  python slack_bot.py

The bot listens on port 3000 (configurable via PORT env var).
In production, run behind a reverse proxy or expose directly via Fly.io.

Slash commands handled:
  /apply    — open a modal to start a resume + cover letter run
  /prep     — open a modal to start an interview prep run
  /status   — check the status of your most recent run
"""

from __future__ import annotations

import json
import os
import threading
import time

import requests
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

SLACK_BOT_TOKEN     = os.environ["SLACK_BOT_TOKEN"]
SLACK_SIGNING_SECRET = os.environ["SLACK_SIGNING_SECRET"]
SLACK_APP_TOKEN     = os.environ.get("SLACK_APP_TOKEN", "")  # xapp-... for Socket Mode
BOT_API_KEY         = os.environ["BOT_API_KEY"]
API_BASE            = os.environ.get("JOB_APPLY_API_URL", "https://job-apply-corey.fly.dev").rstrip("/")
PORT                = int(os.environ.get("PORT", "3000"))

ROUND_TYPES = ["recruiter_screen", "hiring_manager", "technical", "panel", "final", "take_home"]

app = App(token=SLACK_BOT_TOKEN, signing_secret=SLACK_SIGNING_SECRET)

# ---------------------------------------------------------------------------
# API helpers
# ---------------------------------------------------------------------------

def _api(method: str, path: str, **kwargs) -> requests.Response:
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {BOT_API_KEY}"
    return getattr(requests, method)(f"{API_BASE}{path}", headers=headers, timeout=30, **kwargs)


def _post_run(job_posting: str, company: str, role: str, contact: str = "") -> dict:
    r = _api("post", "/api/run", json={
        "job_posting": job_posting,
        "company": company,
        "role": role,
        "contact": contact or None,
    })
    r.raise_for_status()
    return r.json()


def _post_prep(job_posting: str, company: str, role: str,
               round_type: str, focus: str = "", interviewer: str = "") -> dict:
    r = _api("post", "/api/prep", json={
        "job_posting": job_posting,
        "company": company,
        "role": role,
        "round_type": round_type,
        "focus": focus or None,
        "interviewer": interviewer or None,
    })
    r.raise_for_status()
    return r.json()


def _poll_run(run_id: str, timeout: int = 300) -> dict:
    deadline = time.time() + timeout
    while time.time() < deadline:
        r = _api("get", f"/api/run/{run_id}/status")
        r.raise_for_status()
        data = r.json()
        if data["status"] in ("done", "error"):
            return data
        time.sleep(5)
    return {"status": "timeout", "error": "Timed out waiting for run to complete"}


def _poll_prep(prep_id: str, timeout: int = 300) -> dict:
    deadline = time.time() + timeout
    while time.time() < deadline:
        r = _api("get", f"/api/prep/{prep_id}/status")
        r.raise_for_status()
        data = r.json()
        if data["status"] in ("done", "error"):
            return data
        time.sleep(5)
    return {"status": "timeout", "error": "Timed out waiting for prep to complete"}


# ---------------------------------------------------------------------------
# /apply — open modal
# ---------------------------------------------------------------------------

@app.command("/apply")
def apply_command(ack, body, client):
    ack()
    client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "apply_submit",
            "title": {"type": "plain_text", "text": "Generate Application"},
            "submit": {"type": "plain_text", "text": "Generate"},
            "close": {"type": "plain_text", "text": "Cancel"},
            "blocks": [
                {
                    "type": "input",
                    "block_id": "company",
                    "label": {"type": "plain_text", "text": "Company name"},
                    "element": {"type": "plain_text_input", "action_id": "value",
                                "placeholder": {"type": "plain_text", "text": "Acme Corp"}},
                },
                {
                    "type": "input",
                    "block_id": "role",
                    "label": {"type": "plain_text", "text": "Role title"},
                    "element": {"type": "plain_text_input", "action_id": "value",
                                "placeholder": {"type": "plain_text", "text": "Solutions Engineer"}},
                },
                {
                    "type": "input",
                    "block_id": "contact",
                    "optional": True,
                    "label": {"type": "plain_text", "text": "Hiring manager name (optional)"},
                    "element": {"type": "plain_text_input", "action_id": "value",
                                "placeholder": {"type": "plain_text", "text": "Jane Smith"}},
                },
                {
                    "type": "input",
                    "block_id": "job_posting",
                    "label": {"type": "plain_text", "text": "Job posting (paste full text)"},
                    "element": {"type": "plain_text_input", "action_id": "value",
                                "multiline": True,
                                "placeholder": {"type": "plain_text", "text": "Paste the full job description here…"}},
                },
            ],
        },
    )


@app.view("apply_submit")
def apply_view_submit(ack, body, client, view):
    ack()
    vals = view["state"]["values"]
    company     = vals["company"]["value"]["value"].strip()
    role        = vals["role"]["value"]["value"].strip()
    contact     = (vals["contact"]["value"]["value"] or "").strip()
    job_posting = vals["job_posting"]["value"]["value"].strip()
    user_id     = body["user"]["id"]
    channel     = body["user"]["id"]  # DM the user

    def _run():
        client.chat_postMessage(
            channel=channel,
            text=f":hourglass_flowing_sand: Starting application for *{role}* at *{company}*…",
        )
        try:
            run_data = _post_run(job_posting, company, role, contact)
            run_id = run_data["run_id"]
            status = _poll_run(run_id)
        except Exception as exc:
            client.chat_postMessage(channel=channel, text=f":x: Error starting run: {exc}")
            return

        if status["status"] == "done":
            r = _api("get", f"/api/run/{run_id}/status")
            # Full result is on the run object server-side; link to Drive folder or app
            drive_url = API_BASE  # fallback
            try:
                # Fetch the done event details by re-checking — status endpoint
                # doesn't return folder_url, so link to the web app instead.
                drive_url = f"{API_BASE}/"
            except Exception:
                pass

            client.chat_postMessage(
                channel=channel,
                text=(
                    f":white_check_mark: *{role} @ {company}* — done!\n"
                    f"Resume, ATS resume, and cover letter are in your Google Drive.\n"
                    f"<{API_BASE}|Open the app> to download the files."
                ),
            )
        elif status["status"] == "timeout":
            client.chat_postMessage(
                channel=channel,
                text=f":warning: Run is taking longer than expected. Check <{API_BASE}|the app> for status.",
            )
        else:
            client.chat_postMessage(
                channel=channel,
                text=f":x: Run failed: {status.get('error', 'Unknown error')}",
            )

    threading.Thread(target=_run, daemon=True).start()


# ---------------------------------------------------------------------------
# /prep — open modal
# ---------------------------------------------------------------------------

@app.command("/prep")
def prep_command(ack, body, client):
    ack()
    client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "prep_submit",
            "title": {"type": "plain_text", "text": "Interview Prep"},
            "submit": {"type": "plain_text", "text": "Generate"},
            "close": {"type": "plain_text", "text": "Cancel"},
            "blocks": [
                {
                    "type": "input",
                    "block_id": "company",
                    "label": {"type": "plain_text", "text": "Company name"},
                    "element": {"type": "plain_text_input", "action_id": "value",
                                "placeholder": {"type": "plain_text", "text": "Acme Corp"}},
                },
                {
                    "type": "input",
                    "block_id": "role",
                    "label": {"type": "plain_text", "text": "Role title"},
                    "element": {"type": "plain_text_input", "action_id": "value",
                                "placeholder": {"type": "plain_text", "text": "Solutions Engineer"}},
                },
                {
                    "type": "input",
                    "block_id": "round_type",
                    "label": {"type": "plain_text", "text": "Round type"},
                    "element": {
                        "type": "static_select",
                        "action_id": "value",
                        "placeholder": {"type": "plain_text", "text": "Select round type"},
                        "options": [
                            {"text": {"type": "plain_text", "text": rt.replace("_", " ").title()},
                             "value": rt}
                            for rt in ROUND_TYPES
                        ],
                    },
                },
                {
                    "type": "input",
                    "block_id": "interviewer",
                    "optional": True,
                    "label": {"type": "plain_text", "text": "Interviewer name (optional)"},
                    "element": {"type": "plain_text_input", "action_id": "value",
                                "placeholder": {"type": "plain_text", "text": "Jane Smith, VP Engineering"}},
                },
                {
                    "type": "input",
                    "block_id": "focus",
                    "optional": True,
                    "label": {"type": "plain_text", "text": "Focus areas (optional)"},
                    "element": {"type": "plain_text_input", "action_id": "value",
                                "placeholder": {"type": "plain_text", "text": "System design, API architecture"}},
                },
                {
                    "type": "input",
                    "block_id": "job_posting",
                    "label": {"type": "plain_text", "text": "Job posting (paste full text)"},
                    "element": {"type": "plain_text_input", "action_id": "value",
                                "multiline": True,
                                "placeholder": {"type": "plain_text", "text": "Paste the full job description here…"}},
                },
            ],
        },
    )


@app.view("prep_submit")
def prep_view_submit(ack, body, client, view):
    ack()
    vals        = view["state"]["values"]
    company     = vals["company"]["value"]["value"].strip()
    role        = vals["role"]["value"]["value"].strip()
    round_type  = vals["round_type"]["value"]["selected_option"]["value"]
    interviewer = (vals["interviewer"]["value"]["value"] or "").strip()
    focus       = (vals["focus"]["value"]["value"] or "").strip()
    job_posting = vals["job_posting"]["value"]["value"].strip()
    channel     = body["user"]["id"]

    def _run():
        client.chat_postMessage(
            channel=channel,
            text=f":hourglass_flowing_sand: Generating *{round_type.replace('_', ' ').title()}* prep for *{role}* at *{company}*…",
        )
        try:
            prep_data = _post_prep(job_posting, company, role, round_type, focus, interviewer)
            prep_id = prep_data["prep_id"]
            status = _poll_prep(prep_id)
        except Exception as exc:
            client.chat_postMessage(channel=channel, text=f":x: Error starting prep: {exc}")
            return

        if status["status"] == "done":
            client.chat_postMessage(
                channel=channel,
                text=(
                    f":white_check_mark: *{round_type.replace('_', ' ').title()} prep* for *{role} @ {company}* — done!\n"
                    f"Your prep card is in Google Drive.\n"
                    f"<{API_BASE}|Open the app> to download it."
                ),
            )
        elif status["status"] == "timeout":
            client.chat_postMessage(
                channel=channel,
                text=f":warning: Prep is taking longer than expected. Check <{API_BASE}|the app> for status.",
            )
        else:
            client.chat_postMessage(
                channel=channel,
                text=f":x: Prep failed: {status.get('error', 'Unknown error')}",
            )

    threading.Thread(target=_run, daemon=True).start()


# ---------------------------------------------------------------------------
# /jobstatus — check API health
# ---------------------------------------------------------------------------

@app.command("/jobstatus")
def jobstatus_command(ack, respond):
    ack()
    try:
        r = requests.get(f"{API_BASE}/api/health", timeout=10)
        r.raise_for_status()
        data = r.json()
        respond(f":white_check_mark: API is up. Storage configured: `{data.get('storage', '?')}`")
    except Exception as exc:
        respond(f":x: API health check failed: {exc}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if SLACK_APP_TOKEN:
        # Socket Mode — no public URL needed (good for local dev)
        handler = SocketModeHandler(app, SLACK_APP_TOKEN)
        print(f"Starting in Socket Mode")
        handler.start()
    else:
        # HTTP mode — run behind a public URL (Fly.io / ngrok)
        print(f"Starting HTTP server on port {PORT}")
        app.start(port=PORT)
