"""
test_contacts.py
----------------
Tests for contact groups, priority scoring, and group-expanded CLI dispatch.

All tests point OUTLOOK_CONTACTS_FILE at the fixture file so they run without
a real contacts.json and without modifying the working copy.

Run:
    pytest tests/test_contacts.py -v
    pytest tests/ -v          # runs alongside test_cli.py
"""

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

SKILLS_DIR = Path(__file__).parent.parent
CLI        = str(SKILLS_DIR / "outlook_cli.py")
FIXTURE    = str(Path(__file__).parent / "fixtures" / "contacts.json")


# ---------------------------------------------------------------------------
# Helper: run CLI with both OUTLOOK_MOCK and OUTLOOK_CONTACTS_FILE set
# ---------------------------------------------------------------------------

def run(*args: str, input_text: str | None = None) -> tuple[object, int]:
    env = {
        **os.environ,
        "OUTLOOK_MOCK": "1",
        "OUTLOOK_CONTACTS_FILE": FIXTURE,
    }
    result = subprocess.run(
        [sys.executable, CLI, *args],
        capture_output=True,
        text=True,
        env=env,
        cwd=str(SKILLS_DIR),
        input=input_text,
    )
    raw = result.stdout.strip()
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        pytest.fail(
            f"Non-JSON output for args {args!r}:\n"
            f"  stdout: {raw[:400]!r}\n"
            f"  stderr: {result.stderr[:400]!r}"
        )
    return data, result.returncode


# ---------------------------------------------------------------------------
# Direct unit tests for outlook_contacts (in-process, fast)
# ---------------------------------------------------------------------------

@pytest.fixture(autouse=True)
def use_fixture_contacts(monkeypatch):
    """Point outlook_contacts at the fixture file for all in-process tests."""
    monkeypatch.setenv("OUTLOOK_CONTACTS_FILE", FIXTURE)
    # Force reload so the new env var is picked up
    import importlib
    import outlook_contacts as oc
    importlib.reload(oc)
    yield


class TestContactsLoad:
    def test_load_returns_groups(self):
        import outlook_contacts as oc
        data = oc.load()
        assert "groups" in data
        assert "direct_reports" in data["groups"]
        assert len(data["groups"]["direct_reports"]) == 2

    def test_load_returns_contacts(self):
        import outlook_contacts as oc
        data = oc.load()
        assert "alice.chen@co.com" in data["contacts"]


class TestGroupList:
    def test_returns_all_groups(self):
        import outlook_contacts as oc
        groups = oc.group_list()
        names = [g["group"] for g in groups]
        for expected in ("direct_reports", "supervisors", "peers", "program_leads"):
            assert expected in names

    def test_count_field(self):
        import outlook_contacts as oc
        groups = oc.group_list()
        dr = next(g for g in groups if g["group"] == "direct_reports")
        assert dr["count"] == 2

    def test_members_have_name_and_email(self):
        import outlook_contacts as oc
        groups = oc.group_list()
        for g in groups:
            for m in g["members"]:
                assert "email" in m
                assert "name" in m


class TestGroupShow:
    def test_show_direct_reports(self):
        import outlook_contacts as oc
        result = oc.group_show("direct_reports")
        assert result["group"] == "direct_reports"
        emails = [m["email"] for m in result["members"]]
        assert "alice.chen@co.com" in emails
        assert "bob.kim@co.com" in emails

    def test_show_case_insensitive(self):
        import outlook_contacts as oc
        result = oc.group_show("DIRECT_REPORTS")
        assert result["count"] == 2

    def test_show_invalid_group_raises(self):
        import outlook_contacts as oc
        with pytest.raises(KeyError):
            oc.group_show("NONEXISTENT_GROUP_XYZ")

    def test_members_have_title_field(self):
        import outlook_contacts as oc
        result = oc.group_show("direct_reports")
        for m in result["members"]:
            assert "title" in m


class TestGroupEmails:
    def test_returns_list_of_strings(self):
        import outlook_contacts as oc
        emails = oc.group_emails("direct_reports")
        assert isinstance(emails, list)
        assert all(isinstance(e, str) for e in emails)

    def test_correct_emails(self):
        import outlook_contacts as oc
        emails = oc.group_emails("supervisors")
        assert "manager@co.com" in emails

    def test_unknown_group_returns_empty(self):
        import outlook_contacts as oc
        emails = oc.group_emails("GHOST_GROUP")
        assert emails == []


class TestGroupAddRemove:
    def test_add_then_remove(self, tmp_path):
        """Add/remove without touching the real fixture."""
        import importlib
        import shutil
        import outlook_contacts as oc

        # Work on a copy
        tmp_file = tmp_path / "contacts.json"
        shutil.copy(FIXTURE, tmp_file)
        os.environ["OUTLOOK_CONTACTS_FILE"] = str(tmp_file)
        importlib.reload(oc)

        add_result = oc.group_add("peers", "newpeer@co.com", name="New Peer")
        assert add_result["status"] == "ok"
        assert "newpeer@co.com" in oc.group_emails("peers")

        rem_result = oc.group_remove("peers", "newpeer@co.com")
        assert rem_result["removed"] is True
        assert "newpeer@co.com" not in oc.group_emails("peers")

        # Restore fixture path
        os.environ["OUTLOOK_CONTACTS_FILE"] = FIXTURE
        importlib.reload(oc)

    def test_add_idempotent(self, tmp_path):
        """Adding the same email twice should not duplicate it."""
        import shutil
        import importlib
        import outlook_contacts as oc

        tmp_file = tmp_path / "contacts.json"
        shutil.copy(FIXTURE, tmp_file)
        os.environ["OUTLOOK_CONTACTS_FILE"] = str(tmp_file)
        importlib.reload(oc)

        oc.group_add("peers", "dup@co.com")
        oc.group_add("peers", "dup@co.com")
        emails = oc.group_emails("peers")
        assert emails.count("dup@co.com") == 1

        os.environ["OUTLOOK_CONTACTS_FILE"] = FIXTURE
        importlib.reload(oc)


class TestWhoIs:
    def test_known_person(self):
        import outlook_contacts as oc
        result = oc.who_is("alice.chen@co.com")
        assert result["name"] == "Alice Chen"
        assert "direct_reports" in result["groups"]

    def test_manager_in_supervisors(self):
        import outlook_contacts as oc
        result = oc.who_is("manager@co.com")
        assert "supervisors" in result["groups"]

    def test_unknown_person_empty_groups(self):
        import outlook_contacts as oc
        result = oc.who_is("nobody@co.com")
        assert result["groups"] == []


class TestPriorityTier:
    def test_supervisor_tier(self):
        import outlook_contacts as oc
        assert oc.priority_tier("manager@co.com") == "supervisor"

    def test_direct_report_tier(self):
        import outlook_contacts as oc
        assert oc.priority_tier("alice.chen@co.com") == "direct_report"

    def test_peer_tier(self):
        import outlook_contacts as oc
        assert oc.priority_tier("dave.wu@co.com") == "peer"

    def test_program_lead_tier(self):
        import outlook_contacts as oc
        assert oc.priority_tier("carol.jones@co.com") == "program_lead"

    def test_unknown_tier(self):
        import outlook_contacts as oc
        assert oc.priority_tier("stranger@co.com") == "unknown"

    def test_case_insensitive(self):
        import outlook_contacts as oc
        assert oc.priority_tier("MANAGER@CO.COM") == "supervisor"


class TestEnrichMessages:
    def test_adds_tier_field(self):
        import outlook_contacts as oc
        msgs = [
            {"entry_id": "M1", "sender_email": "manager@co.com",
             "subject": "Budget approval needed", "body": "Please approve",
             "received": "2026-05-23T10:00:00", "unread": True},
        ]
        enriched = oc.enrich_messages(msgs)
        assert enriched[0]["tier"] == "supervisor"

    def test_always_surface_signal(self):
        import outlook_contacts as oc
        msgs = [
            {"entry_id": "M1", "sender_email": "manager@co.com",
             "subject": "FYI", "body": "Just letting you know.",
             "received": "2026-05-23T10:00:00", "unread": True},
        ]
        enriched = oc.enrich_messages(msgs)
        assert "always_surface" in enriched[0]["priority_signals"]

    def test_blocker_flag_for_direct_report(self):
        import outlook_contacts as oc
        msgs = [
            {"entry_id": "M2", "sender_email": "alice.chen@co.com",
             "subject": "Issue on CDR", "body": "I'm blocked on the antenna trade study.",
             "received": "2026-05-23T11:00:00", "unread": True},
        ]
        enriched = oc.enrich_messages(msgs)
        assert "blocker_flag" in enriched[0]["priority_signals"]

    def test_review_flag_for_supervisor(self):
        import outlook_contacts as oc
        msgs = [
            {"entry_id": "M3", "sender_email": "manager@co.com",
             "subject": "Sign-off Required", "body": "Please provide sign-off on the subcontract.",
             "received": "2026-05-23T08:00:00", "unread": True},
        ]
        enriched = oc.enrich_messages(msgs)
        assert "review_flag" in enriched[0]["priority_signals"]

    def test_unknown_sender_no_signals(self):
        import outlook_contacts as oc
        msgs = [
            {"entry_id": "M4", "sender_email": "stranger@co.com",
             "subject": "Hello", "body": "Just saying hi.",
             "received": "2026-05-23T09:00:00", "unread": False},
        ]
        enriched = oc.enrich_messages(msgs)
        assert enriched[0]["tier"] == "unknown"
        assert enriched[0]["priority_signals"] == []

    def test_supervisor_sorted_first(self):
        import outlook_contacts as oc
        msgs = [
            {"entry_id": "M_peer", "sender_email": "dave.wu@co.com",
             "subject": "Peer msg", "body": "x",
             "received": "2026-05-23T09:00:00", "unread": True},
            {"entry_id": "M_sup", "sender_email": "manager@co.com",
             "subject": "Supervisor msg", "body": "x",
             "received": "2026-05-23T08:00:00", "unread": True},
            {"entry_id": "M_dr", "sender_email": "alice.chen@co.com",
             "subject": "DR msg", "body": "x",
             "received": "2026-05-23T07:00:00", "unread": True},
        ]
        enriched = oc.enrich_messages(msgs)
        assert enriched[0]["entry_id"] == "M_sup", "Supervisor should sort first"
        assert enriched[1]["entry_id"] == "M_dr", "Direct report should sort second"


class TestContactsResolve:
    def test_natural_language_my_team(self):
        import outlook_contacts as oc
        result = oc.resolve("my team")
        assert result["type"] == "group"
        assert result["group"] == "direct_reports"
        assert len(result["emails"]) == 2

    def test_natural_language_boss(self):
        import outlook_contacts as oc
        result = oc.resolve("boss")
        assert result["type"] == "group"
        assert "manager@co.com" in result["emails"]

    def test_group_name_direct(self):
        import outlook_contacts as oc
        result = oc.resolve("direct_reports")
        assert result["type"] == "group"

    def test_contact_alias(self):
        import outlook_contacts as oc
        result = oc.resolve("Alice")
        assert result["type"] == "contact"
        assert result["emails"] == ["alice.chen@co.com"]

    def test_contact_full_name(self):
        import outlook_contacts as oc
        result = oc.resolve("Alice Chen")
        assert result["type"] == "contact"
        assert "alice.chen@co.com" in result["emails"]

    def test_unknown_returns_type_unknown(self):
        import outlook_contacts as oc
        result = oc.resolve("Zaphod Beeblebrox")
        assert result["type"] == "unknown"
        assert result["emails"] == []


# ---------------------------------------------------------------------------
# CLI subprocess tests for group commands
# ---------------------------------------------------------------------------

class TestGroupCLI:
    def test_group_list(self):
        data, rc = run("group", "list")
        assert rc == 0
        assert isinstance(data, list)
        group_names = [g["group"] for g in data]
        assert "direct_reports" in group_names

    def test_group_show(self):
        data, rc = run("group", "show", "direct_reports")
        assert rc == 0
        assert data["group"] == "direct_reports"
        assert data["count"] == 2

    def test_group_emails_returns_list(self):
        data, rc = run("group", "emails", "direct_reports")
        assert rc == 0
        assert isinstance(data, list)
        assert "alice.chen@co.com" in data

    def test_group_who(self):
        data, rc = run("group", "who", "manager@co.com")
        assert rc == 0
        assert "supervisors" in data["groups"]
        assert data["name"] == "Jane Director"

    def test_group_tier(self):
        data, rc = run("group", "tier", "alice.chen@co.com")
        assert rc == 0
        assert data["tier"] == "direct_report"

    def test_group_unknown_shows_error(self):
        data, rc = run("group", "show", "GHOST_GROUP_XYZ")
        assert rc == 1
        assert "error" in data


class TestResolveWithContacts:
    def test_resolve_natural_language(self):
        data, rc = run("resolve", "my team")
        assert rc == 0
        assert data["type"] == "group"
        assert len(data["emails"]) == 2

    def test_resolve_contact_alias(self):
        data, rc = run("resolve", "Alice")
        assert rc == 0
        assert data["type"] == "contact"
        assert "alice.chen@co.com" in data["emails"]

    def test_resolve_falls_back_to_gal_for_unknown(self):
        # "Tom" is in the mock GAL but not in contacts.json fixture
        data, rc = run("resolve", "Tom")
        assert rc == 0
        # Should resolve via GAL (mock knows Tom)
        assert data.get("resolved") is True or data.get("type") == "contact"


class TestGroupDispatch:
    def test_draft_to_group(self):
        data, rc = run("draft",
                       "--to-group", "direct_reports",
                       "--subject", "Team Update",
                       "--body", "Hi team, quick update...")
        assert rc == 0
        assert data["status"] == "draft"
        assert "alice.chen@co.com" in data["to"] or "alice.chen" in data["to"]

    def test_draft_to_group_and_extra_recipient(self):
        data, rc = run("draft",
                       "--to", "extra@co.com",
                       "--to-group", "supervisors",
                       "--subject", "FYI",
                       "--body", "Loop-in note.")
        assert rc == 0
        assert data["status"] == "draft"
        # Both the group member and the extra should be in the to list
        assert "manager@co.com" in data["to"]
        assert "extra@co.com" in data["to"]

    def test_draft_invalid_group_errors(self):
        data, rc = run("draft",
                       "--to-group", "NONEXISTENT_GROUP_XYZ",
                       "--subject", "Test",
                       "--body", "x")
        assert rc == 1
        assert "error" in data

    def test_meeting_with_attendees_group(self):
        data, rc = run("meeting",
                       "--subject", "Team Sync",
                       "--start", "2026-06-02T09:00:00",
                       "--duration", "30",
                       "--attendees-group", "direct_reports")
        assert rc == 0
        assert data["status"] == "saved"
        assert "alice.chen@co.com" in data["attendees"]
        assert "bob.kim@co.com" in data["attendees"]

    def test_meeting_attendees_group_and_extra(self):
        data, rc = run("meeting",
                       "--subject", "Cross-team Sync",
                       "--start", "2026-06-03T10:00:00",
                       "--attendees", "extra@co.com",
                       "--attendees-group", "program_leads")
        assert rc == 0
        assert "carol.jones@co.com" in data["attendees"]
        assert "extra@co.com" in data["attendees"]

    def test_forward_to_group(self):
        data, rc = run("forward", "EMAIL001",
                       "--to-group", "supervisors",
                       "--body", "FYI — needs your review.")
        assert rc == 0
        assert data["status"] == "draft"
        assert "manager@co.com" in data["to"]


class TestEnrichCLI:
    def test_enrich_returns_list(self):
        data, rc = run("enrich", "--hours", "168")
        assert rc == 0
        assert isinstance(data, list)

    def test_enrich_adds_tier_field(self):
        data, _ = run("enrich", "--hours", "168")
        for msg in data:
            assert "tier" in msg, f"Missing 'tier' in {msg.get('subject')!r}"

    def test_enrich_adds_priority_signals(self):
        data, _ = run("enrich", "--hours", "168")
        for msg in data:
            assert "priority_signals" in msg

    def test_enrich_unread_only(self):
        data, rc = run("enrich", "--hours", "168", "--unread")
        assert rc == 0
        assert all(m.get("unread") for m in data)
