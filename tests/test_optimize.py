"""Tests for the Optimize Run agent — /api/optimize endpoints and the pure
apply.py helpers (_apply_optimize_edits, _parse_cover_letter_text)."""

import uuid
from unittest.mock import MagicMock, patch

import pytest

from tests.conftest import _store


# ── Seeding helpers ───────────────────────────────────────────────────────────

FOLDER_ID = "drive-folder-abc123"


def _seed_app_with_run(user_record, folder_id: str = FOLDER_ID) -> str:
    """Create an application record with a linked resume run owned by the user."""
    from scripts.applications import save_application

    app_id = str(uuid.uuid4())
    save_application(user_record["user_id"], {
        "id":          app_id,
        "company":     "Acme",
        "role_title":  "Software Engineer",
        "status":      "Applied",
        "created_at":  "2026-01-01T00:00:00Z",
        "linked_runs": [{
            "id":               str(uuid.uuid4()),
            "type":             "resume",
            "folder_name":      "Acme_SoftwareEngineer",
            "folder_url":       f"https://drive.google.com/drive/folders/{folder_id}",
            "gdrive_folder_id": folder_id,
            "linked_at":        "2026-01-01T00:00:00Z",
            "linked_by":        "system",
        }],
    })
    return app_id


def _body(app_id: str, **overrides) -> dict:
    body = {
        "app_id":      app_id,
        "folder_id":   FOLDER_ID,
        "instruction": "Emphasize Kubernetes more and shorten the summary.",
        "company":     "Acme",
        "role":        "Software Engineer",
    }
    body.update(overrides)
    return body


# ── /api/optimize endpoint tests ──────────────────────────────────────────────

class TestOptimizeEndpoint:
    def test_requires_auth(self, client):
        r = client.post("/api/optimize", json=_body("some-app"))
        assert r.status_code == 401

    def test_missing_fields_rejected(self, authed_client):
        r = authed_client.post("/api/optimize", json={"instruction": "x"})
        assert r.status_code == 422

    def test_blank_instruction_rejected(self, authed_client, user_record):
        app_id = _seed_app_with_run(user_record)
        r = authed_client.post("/api/optimize", json=_body(app_id, instruction="   "))
        assert r.status_code == 400
        assert "instruction" in r.json()["detail"].lower()

    def test_overlong_instruction_rejected(self, authed_client, user_record):
        app_id = _seed_app_with_run(user_record)
        r = authed_client.post("/api/optimize", json=_body(app_id, instruction="x" * 4001))
        assert r.status_code == 400

    def test_no_documents_selected_rejected(self, authed_client, user_record):
        app_id = _seed_app_with_run(user_record)
        r = authed_client.post("/api/optimize", json=_body(
            app_id, optimize_resume=False, optimize_cover_letter=False))
        assert r.status_code == 400

    def test_foreign_folder_rejected(self, authed_client, user_record):
        app_id = _seed_app_with_run(user_record)
        r = authed_client.post("/api/optimize", json=_body(app_id, folder_id="not-my-folder"))
        assert r.status_code == 403

    def test_no_linked_runs_rejected(self, authed_client, user_record):
        # An app with no linked runs cannot authorize any folder
        from scripts.applications import save_application
        app_id = str(uuid.uuid4())
        save_application(user_record["user_id"], {
            "id": app_id, "company": "Acme", "role_title": "SE",
            "status": "Applied", "created_at": "2026-01-01T00:00:00Z",
        })
        r = authed_client.post("/api/optimize", json=_body(app_id))
        assert r.status_code == 403

    def test_returns_optimize_id_and_machine_id(self, authed_client, user_record):
        app_id = _seed_app_with_run(user_record)
        with patch("api.threading.Thread") as mock_thread:
            mock_thread.return_value = MagicMock()
            r = authed_client.post("/api/optimize", json=_body(app_id))
        assert r.status_code == 200
        d = r.json()
        assert uuid.UUID(d["optimize_id"])
        assert "machine_id" in d

    def test_status_found_after_create(self, authed_client, user_record):
        app_id = _seed_app_with_run(user_record)
        with patch("api.threading.Thread") as mock_thread:
            mock_thread.return_value = MagicMock()
            r = authed_client.post("/api/optimize", json=_body(app_id))
        optimize_id = r.json()["optimize_id"]
        r2 = authed_client.get(f"/api/optimize/{optimize_id}/status")
        assert r2.status_code == 200
        assert r2.json()["status"] == "queued"

    def test_status_unknown_optimize(self, authed_client):
        r = authed_client.get(f"/api/optimize/{uuid.uuid4()}/status")
        assert r.status_code == 404

    def test_admin_blocked(self, admin_client, admin_record):
        app_id = _seed_app_with_run(admin_record)
        r = admin_client.post("/api/optimize", json=_body(app_id))
        assert r.status_code == 403


# ── Unit tests: _apply_optimize_edits ────────────────────────────────────────

def _noop(_msg: str) -> None:
    pass


class TestApplyOptimizeEdits:
    def test_full_match_with_entities(self):
        from apply import _apply_optimize_edits
        xml = (
            '<w:p><w:r><w:t>Agentic AI &amp; Automation</w:t></w:r></w:p>'
            '<w:p><w:r><w:t xml:space="preserve">Old summary text here</w:t></w:r></w:p>'
        )
        field_map = {
            "competency_1": "Agentic AI & Automation",
            "summary":      "Old summary text here",
        }
        edits = [
            {"field": "competency_1", "new": "LLM Pipelines & RAG"},
            {"field": "summary",      "new": "New summary text"},
        ]
        out, ok, total = _apply_optimize_edits(xml, edits, field_map, _noop)
        assert (ok, total) == (2, 2)
        assert "LLM Pipelines &amp; RAG" in out
        assert "New summary text" in out
        assert "Old summary" not in out

    def test_unknown_field_skipped(self):
        from apply import _apply_optimize_edits
        xml = "<w:t>Something</w:t>"
        out, ok, total = _apply_optimize_edits(
            xml, [{"field": "nope", "new": "x"}], {"summary": "Something"}, _noop)
        assert (ok, total) == (0, 1)
        assert out == xml

    def test_zero_matches_when_text_absent(self):
        from apply import _apply_optimize_edits
        xml = "<w:t>Completely different content</w:t>"
        out, ok, total = _apply_optimize_edits(
            xml, [{"field": "summary", "new": "x"}], {"summary": "Not in the doc"}, _noop)
        assert (ok, total) == (0, 1)
        assert out == xml

    def test_unchanged_text_skipped(self):
        from apply import _apply_optimize_edits
        xml = "<w:t>Same text</w:t>"
        out, ok, total = _apply_optimize_edits(
            xml, [{"field": "summary", "new": "Same text"}], {"summary": "Same text"}, _noop)
        assert (ok, total) == (0, 1)
        assert out == xml

    def test_page_break_split_collapsed(self):
        from apply import _apply_optimize_edits
        # Word split one logical run across a lastRenderedPageBreak
        xml = (
            '<w:p><w:r><w:t>Delivered 20+ integrations for the GateWay '
            '</w:t><w:lastRenderedPageBreak/><w:t>customer portal</w:t></w:r></w:p>'
        )
        field_map = {"job2_bullet3": "Delivered 20+ integrations for the GateWay customer portal"}
        edits = [{"field": "job2_bullet3", "new": "Shipped 25 GateWay portal integrations"}]
        out, ok, total = _apply_optimize_edits(xml, edits, field_map, _noop)
        assert (ok, total) == (1, 1)
        assert "Shipped 25 GateWay portal integrations" in out
        assert "lastRenderedPageBreak" not in out

    def test_partial_success(self):
        from apply import _apply_optimize_edits
        xml = "<w:t>Real field content</w:t>"
        field_map = {"summary": "Real field content", "tagline": "Missing from doc"}
        edits = [
            {"field": "summary", "new": "Updated content"},
            {"field": "tagline", "new": "New tagline"},
        ]
        out, ok, total = _apply_optimize_edits(xml, edits, field_map, _noop)
        assert (ok, total) == (1, 2)
        assert "Updated content" in out


# ── Unit tests: _parse_cover_letter_text ─────────────────────────────────────

SAMPLE_LETTER = """COREY LAVERDIERE

978-790-4272 | cdl825@gmail.com | Sterling, MA

June 1, 2026

Jane Doe

Acme Corp

Re: Software Engineer

Dear Jane Doe,

First paragraph of the letter body which wraps
across lines the way pandoc renders it.

Second paragraph with a proof point.

Third paragraph, different angle.

Fourth paragraph, the differentiator.

Short close.

Sincerely,

Corey Laverdiere

978-790-4272 | cdl825@gmail.com"""


class TestParseCoverLetterText:
    def test_parses_paragraphs_and_contact(self):
        from apply import _parse_cover_letter_text
        r = _parse_cover_letter_text(SAMPLE_LETTER)
        assert r["contact_name"] == "Jane Doe"
        assert len(r["paragraphs"]) == 5
        assert r["paragraphs"][0].startswith("First paragraph")
        # pandoc line-wraps are collapsed to single-line paragraphs
        assert "\n" not in r["paragraphs"][0]
        assert r["paragraphs"][4] == "Short close."

    def test_missing_salutation_raises(self):
        from apply import _parse_cover_letter_text, WorkflowError
        with pytest.raises(WorkflowError):
            _parse_cover_letter_text("No letter structure at all.\n\nJust text.")

    def test_missing_signoff_raises(self):
        from apply import _parse_cover_letter_text, WorkflowError
        with pytest.raises(WorkflowError):
            _parse_cover_letter_text("Dear Team,\n\nBody paragraph.\n\nNo sign-off here.")

    def test_hiring_team_fallback_without_re_line(self):
        from apply import _parse_cover_letter_text
        plain = "Dear Hiring Team,\n\nOnly paragraph.\n\nSincerely,\n\nCorey"
        r = _parse_cover_letter_text(plain)
        assert r["contact_name"] == "Hiring Team"
        assert r["paragraphs"] == ["Only paragraph."]


# ── Unit tests: pattern-based Drive upsert/delete (self-healing filenames) ────

class TestGdriveUpsertFileByPattern:
    """A legacy Drive file (e.g. an ALL-CAPS name from before title_case_name()
    existed) must be found by document-type pattern and renamed in place —
    not left under its old name, and not duplicated under the new one."""

    def test_renames_existing_file_matching_pattern(self, tmp_path):
        from apply import _gdrive_upsert_file_by_pattern

        local = tmp_path / "Resume_CoreyLaverdiere_Acme_SE.docx"
        local.write_bytes(b"docx bytes")

        service = MagicMock()
        service.files.return_value.list.return_value.execute.return_value = {
            "files": [{"id": "file-1", "name": "Resume_COREYLAVERDIERE_Acme_SE.docx"}]
        }
        service.files.return_value.update.return_value.execute.return_value = {"id": "file-1"}

        _gdrive_upsert_file_by_pattern(
            service, "folder-1", r"^Resume_(?!.*_ATS\.docx$).*\.docx$",
            "Resume_CoreyLaverdiere_Acme_SE.docx", local,
        )

        kwargs = service.files.return_value.update.call_args.kwargs
        assert kwargs["fileId"] == "file-1"
        assert kwargs["body"] == {"name": "Resume_CoreyLaverdiere_Acme_SE.docx"}
        service.files.return_value.create.assert_not_called()

    def test_no_rename_when_name_already_matches(self, tmp_path):
        from apply import _gdrive_upsert_file_by_pattern

        local = tmp_path / "Resume_CoreyLaverdiere_Acme_SE.docx"
        local.write_bytes(b"docx bytes")

        service = MagicMock()
        service.files.return_value.list.return_value.execute.return_value = {
            "files": [{"id": "file-1", "name": "Resume_CoreyLaverdiere_Acme_SE.docx"}]
        }
        service.files.return_value.update.return_value.execute.return_value = {"id": "file-1"}

        _gdrive_upsert_file_by_pattern(
            service, "folder-1", r"^Resume_(?!.*_ATS\.docx$).*\.docx$",
            "Resume_CoreyLaverdiere_Acme_SE.docx", local,
        )

        assert service.files.return_value.update.call_args.kwargs["body"] == {}

    def test_creates_when_no_existing_match(self, tmp_path):
        from apply import _gdrive_upsert_file_by_pattern

        local = tmp_path / "Resume_CoreyLaverdiere_Acme_SE.docx"
        local.write_bytes(b"docx bytes")

        service = MagicMock()
        service.files.return_value.list.return_value.execute.return_value = {"files": []}
        service.files.return_value.create.return_value.execute.return_value = {"id": "new-id"}

        _gdrive_upsert_file_by_pattern(
            service, "folder-1", r"^Resume_(?!.*_ATS\.docx$).*\.docx$",
            "Resume_CoreyLaverdiere_Acme_SE.docx", local,
        )

        kwargs = service.files.return_value.create.call_args.kwargs
        assert kwargs["body"] == {"name": "Resume_CoreyLaverdiere_Acme_SE.docx", "parents": ["folder-1"]}
        service.files.return_value.update.assert_not_called()

    def test_ignores_non_matching_files_in_folder(self, tmp_path):
        # An ATS resume already in the folder must not be mistaken for the
        # styled (non-ATS) resume just because both start with "Resume_".
        from apply import _gdrive_upsert_file_by_pattern

        local = tmp_path / "Resume_CoreyLaverdiere_Acme_SE.docx"
        local.write_bytes(b"docx bytes")

        service = MagicMock()
        service.files.return_value.list.return_value.execute.return_value = {
            "files": [{"id": "ats-1", "name": "Resume_COREYLAVERDIERE_Acme_SE_ATS.docx"}]
        }
        service.files.return_value.create.return_value.execute.return_value = {"id": "new-id"}

        _gdrive_upsert_file_by_pattern(
            service, "folder-1", r"^Resume_(?!.*_ATS\.docx$).*\.docx$",
            "Resume_CoreyLaverdiere_Acme_SE.docx", local,
        )

        service.files.return_value.create.assert_called_once()
        service.files.return_value.update.assert_not_called()


class TestGdriveDeleteByPattern:
    def test_deletes_only_matching_files(self):
        from apply import _gdrive_delete_by_pattern

        service = MagicMock()
        service.files.return_value.list.return_value.execute.return_value = {
            "files": [
                {"id": "old-pdf", "name": "Resume_COREYLAVERDIERE_Acme_SE.pdf"},
                {"id": "other", "name": "job_description.md"},
            ]
        }
        service.files.return_value.delete.return_value.execute.return_value = {}

        _gdrive_delete_by_pattern(service, "folder-1", r"^Resume_.*\.pdf$")

        service.files.return_value.delete.assert_called_once_with(fileId="old-pdf")
