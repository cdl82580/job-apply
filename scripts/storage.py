"""
scripts/storage.py — Tigris (S3-compatible) storage for user data and resumes.

Key layout:
  users/{sha256(email)}.json          — account record
  user_ids/{user_id}.txt              — reverse lookup: user_id → email
  resumes/{user_id}/master.docx       — uploaded master resume
  profiles/{user_id}/profile.md       — voice / profile guide
"""

from __future__ import annotations

import hashlib
import json
import os
from typing import Any

try:
    import boto3
    from botocore.exceptions import ClientError
    _HAS_BOTO3 = True
except ImportError:
    _HAS_BOTO3 = False

BUCKET    = os.environ.get("BUCKET_NAME", "job-apply-storage")
_ENDPOINT = os.environ.get("AWS_ENDPOINT_URL_S3", "https://fly.storage.tigris.dev")
_REGION   = os.environ.get("AWS_REGION", "auto")


def is_configured() -> bool:
    return _HAS_BOTO3 and bool(os.environ.get("AWS_ACCESS_KEY_ID"))


def _client():
    if not _HAS_BOTO3:
        raise RuntimeError("boto3 not installed — run: pip install boto3")
    return boto3.client("s3", endpoint_url=_ENDPOINT, region_name=_REGION)


def _email_key(email: str) -> str:
    return hashlib.sha256(email.strip().lower().encode()).hexdigest()


# ---------------------------------------------------------------------------
# Low-level primitives
# ---------------------------------------------------------------------------

def put_bytes(key: str, data: bytes, content_type: str = "application/octet-stream") -> None:
    _client().put_object(Bucket=BUCKET, Key=key, Body=data, ContentType=content_type)


def put_text(key: str, text: str) -> None:
    put_bytes(key, text.encode("utf-8"), "text/plain; charset=utf-8")


def get_bytes(key: str) -> bytes | None:
    try:
        resp = _client().get_object(Bucket=BUCKET, Key=key)
        return resp["Body"].read()
    except Exception as e:
        if hasattr(e, "response") and e.response.get("Error", {}).get("Code") == "NoSuchKey":
            return None
        raise


def get_text(key: str) -> str | None:
    data = get_bytes(key)
    return data.decode("utf-8") if data is not None else None


def exists(key: str) -> bool:
    try:
        _client().head_object(Bucket=BUCKET, Key=key)
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# User accounts
# ---------------------------------------------------------------------------

def save_user(user: dict[str, Any]) -> None:
    """Persist a user record. Also writes a user_id → email index entry."""
    put_text(f"users/{_email_key(user['email'])}.json", json.dumps(user))
    put_text(f"user_ids/{user['user_id']}.txt", user["email"])


def get_user_by_email(email: str) -> dict[str, Any] | None:
    data = get_text(f"users/{_email_key(email)}.json")
    return json.loads(data) if data else None


def get_user_by_id(user_id: str) -> dict[str, Any] | None:
    email = get_text(f"user_ids/{user_id}.txt")
    return get_user_by_email(email.strip()) if email else None


# ---------------------------------------------------------------------------
# Resumes and profiles
# ---------------------------------------------------------------------------

_DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def save_resume(user_id: str, data: bytes) -> None:
    put_bytes(f"resumes/{user_id}/master.docx", data, _DOCX_MIME)


def get_resume(user_id: str) -> bytes | None:
    return get_bytes(f"resumes/{user_id}/master.docx")


def has_resume(user_id: str) -> bool:
    return exists(f"resumes/{user_id}/master.docx")


def save_profile(user_id: str, text: str) -> None:
    put_text(f"profiles/{user_id}/profile.md", text)


def get_profile(user_id: str) -> str | None:
    return get_text(f"profiles/{user_id}/profile.md")
