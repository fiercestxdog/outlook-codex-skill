"""
outlook_contacts.py
-------------------
Contact group management for the Outlook skill.
Pure Python — zero COM / pywin32 dependency.

Groups are stored in contacts.json (same directory by default).
Override the path with the OUTLOOK_CONTACTS_FILE env var for testing
or multi-profile setups.

Schema (contacts.json):
  {
    "contacts": {
      "<email>": {
        "name": "...",
        "title": "...",
        "aliases": ["first name", "nickname", ...],
        "notes": "..."
      }
    },
    "groups": {
      "direct_reports": ["alice@co.com", ...],
      "supervisors":    ["manager@co.com"],
      "peers":          [...],
      "program_leads":  [...]
    },
    "natural_language": {
      "my team": "direct_reports",
      "boss":    "supervisors"
    }
  }

Dependencies:
    python >= 3.8   (stdlib only)

Install:
    No install required.
"""

__version__ = "2.0.0"

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

_DEFAULT_PATH = Path(__file__).parent / "contacts.json"

_EMPTY: dict[str, Any] = {
    "contacts": {},
    "groups": {},
    "natural_language": {},
}

# Priority tiers from highest to lowest — determines brief ordering
# Groups not listed here → "unknown" (neutral)
_TIER_ORDER: list[str] = ["supervisors", "program_leads", "direct_reports", "peers"]
_TIER_NAMES: dict[str, str] = {
    "supervisors":    "supervisor",
    "program_leads":  "program_lead",
    "direct_reports": "direct_report",
    "peers":          "peer",
}


# ---------------------------------------------------------------------------
# I/O helpers
# ---------------------------------------------------------------------------

def _path() -> Path:
    override = os.environ.get("OUTLOOK_CONTACTS_FILE")
    return Path(override) if override else _DEFAULT_PATH


def load() -> dict[str, Any]:
    """Return the full contacts.json dict; empty structure if file is absent."""
    p = _path()
    if not p.exists():
        return {k: dict(v) for k, v in _EMPTY.items()}
    with open(p, encoding="utf-8") as fh:
        return json.load(fh)


def save(data: dict[str, Any]) -> None:
    p = _path()
    with open(p, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)
        fh.write("\n")


# ---------------------------------------------------------------------------
# Group queries
# ---------------------------------------------------------------------------

def group_list() -> list[dict]:
    """Return all groups with member counts and name lists."""
    data = load()
    contacts = data.get("contacts", {})
    out = []
    for gname, emails in data.get("groups", {}).items():
        members = [
            {"email": e, "name": contacts.get(e, {}).get("name", e)}
            for e in emails
        ]
        out.append({"group": gname, "count": len(emails), "members": members})
    return out


def group_show(group_name: str) -> dict:
    """Return full member details for a group."""
    data = load()
    norm = group_name.lower()
    groups = data.get("groups", {})
    if norm not in groups:
        available = list(groups.keys())
        raise KeyError(f"Group {group_name!r} not found. Available: {available}")
    contacts = data.get("contacts", {})
    members = []
    for email in groups[norm]:
        info = contacts.get(email, {})
        members.append({
            "email":  email,
            "name":   info.get("name", email),
            "title":  info.get("title", ""),
            "notes":  info.get("notes", ""),
        })
    return {"group": norm, "count": len(members), "members": members}


def group_emails(group_name: str) -> list[str]:
    """Return bare email list for a group (for --to / --attendees piping)."""
    return list(load().get("groups", {}).get(group_name.lower(), []))


def group_add(group_name: str, email: str, name: str | None = None) -> dict:
    """Add an email to a group, optionally recording a display name."""
    data = load()
    norm = group_name.lower()
    data.setdefault("groups", {}).setdefault(norm, [])
    if email not in data["groups"][norm]:
        data["groups"][norm].append(email)
    if name:
        data.setdefault("contacts", {}).setdefault(email, {})["name"] = name
    save(data)
    return {"status": "ok", "group": norm, "email": email,
            "count": len(data["groups"][norm])}


def group_remove(group_name: str, email: str) -> dict:
    """Remove an email from a group."""
    data = load()
    norm = group_name.lower()
    before = list(data.get("groups", {}).get(norm, []))
    after  = [e for e in before if e.lower() != email.lower()]
    data.setdefault("groups", {})[norm] = after
    save(data)
    return {"status": "ok", "group": norm, "email": email,
            "removed": len(after) < len(before), "count": len(after)}


def who_is(email: str) -> dict:
    """Return which groups a person belongs to and their contact metadata."""
    data = load()
    email_lower = email.lower()
    in_groups = [
        g for g, members in data.get("groups", {}).items()
        if email_lower in [m.lower() for m in members]
    ]
    info = data.get("contacts", {}).get(email, {})
    return {
        "email":  email,
        "name":   info.get("name", email),
        "title":  info.get("title", ""),
        "groups": in_groups,
        "notes":  info.get("notes", ""),
    }


# ---------------------------------------------------------------------------
# Priority scoring
# ---------------------------------------------------------------------------

def priority_tier(email: str) -> str:
    """
    Return the relationship tier for brief priority-scoring.

    Tiers (highest → lowest):
      supervisor   — your chain of command; unread always surfaces
      program_lead — PM, Chief Eng, key stakeholders
      direct_report— people you manage; flag blockers/urgency signals
      peer         — same-level colleagues
      unknown      — not in any group (neutral)
    """
    data = load()
    email_lower = email.lower()
    for group_key in _TIER_ORDER:
        members = [m.lower() for m in data.get("groups", {}).get(group_key, [])]
        if email_lower in members:
            return _TIER_NAMES[group_key]
    return "unknown"


def enrich_messages(messages: list[dict]) -> list[dict]:
    """
    Add 'tier' and 'priority_signals' to each message dict.
    Called by the skill before synthesis to enable relationship-aware sorting.

    Signals added per tier:
      supervisor    → always_surface=True
      direct_report → check body for blocker keywords → blocker_flag
      program_lead  → check for approval/review request → review_flag
    """
    _BLOCKER_WORDS = {"blocked", "blocker", "stuck", "waiting on you", "urgent",
                      "need your input", "need input", "action required"}
    _REVIEW_WORDS  = {"approval", "approve", "sign-off", "sign off", "review",
                      "decision needed"}

    enriched = []
    for msg in messages:
        m = dict(msg)
        email = (m.get("sender_email") or "").lower()
        tier  = priority_tier(email)
        m["tier"] = tier

        signals: list[str] = []
        body = (m.get("body") or "").lower()
        subject = (m.get("subject") or "").lower()
        text = f"{subject} {body}"

        if tier == "supervisor":
            signals.append("always_surface")
        if tier == "direct_report":
            if any(w in text for w in _BLOCKER_WORDS):
                signals.append("blocker_flag")
        if tier in ("supervisor", "program_lead"):
            if any(w in text for w in _REVIEW_WORDS):
                signals.append("review_flag")

        # Stale unread from anyone noteworthy
        if m.get("unread") and tier != "unknown":
            import datetime as _dt
            try:
                age_hours = (
                    _dt.datetime.now() -
                    _dt.datetime.fromisoformat(m.get("received", ""))
                ).total_seconds() / 3600
                if age_hours > 24:
                    signals.append(f"stale_{int(age_hours)}h")
            except Exception:
                pass

        m["priority_signals"] = signals
        enriched.append(m)

    # Sort: supervisor first, then program_lead, then direct_report, then rest
    _tier_rank = {"supervisor": 0, "program_lead": 1, "direct_report": 2, "peer": 3, "unknown": 4}
    enriched.sort(key=lambda m: (_tier_rank.get(m["tier"], 4),
                                  0 if m.get("priority_signals") else 1))
    return enriched


# ---------------------------------------------------------------------------
# Resolve: group name / natural-language alias / contact alias → emails
# ---------------------------------------------------------------------------

def resolve(name: str) -> dict:
    """
    Resolve a name/alias/group reference to email(s).

    Checks in order:
      1. Natural-language alias  ("my team" → direct_reports)
      2. Exact group name        ("direct_reports")
      3. Contact aliases + names ("Alice", "Alice Chen")
      4. Direct email match

    Returns:
      {"type": "group"|"contact"|"unknown",
       "emails": [...],
       "group": <name if group>,
       "name": <display name if contact>,
       "input": <original>}
    """
    data = load()
    norm = name.strip().lower()

    # 1. Natural-language alias
    nl = data.get("natural_language", {})
    if norm in nl:
        group_name = nl[norm]
        emails = list(data.get("groups", {}).get(group_name, []))
        return {"type": "group", "group": group_name, "emails": emails, "input": name}

    # 2. Exact group name
    if norm in data.get("groups", {}):
        emails = list(data["groups"][norm])
        return {"type": "group", "group": norm, "emails": emails, "input": name}

    # 3. Contact alias / name / email
    for email, info in data.get("contacts", {}).items():
        aliases  = [a.lower() for a in info.get("aliases", [])]
        disp     = info.get("name", "").lower()
        if norm in aliases or norm == disp or norm == email.lower():
            return {"type": "contact", "emails": [email],
                    "name": info.get("name", email), "input": name}

    return {"type": "unknown", "emails": [], "input": name}
