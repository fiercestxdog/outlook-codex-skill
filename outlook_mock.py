"""
outlook_mock.py
---------------
In-memory Outlook stub for testing outlook_cli.py without the COM bridge.

Activate with: OUTLOOK_MOCK=1 python outlook_cli.py <cmd>

State is module-global and resets automatically on each subprocess invocation
(each CLI call is a fresh process). For in-process test sequences call
reset_state() between cases.

All public functions match the API of outlook_helpers.py exactly so the CLI
never needs to know it is talking to a stub.

Dependencies:
    python >= 3.8   (stdlib only)
"""

from __future__ import annotations

import copy
import datetime as dt
import uuid
from typing import Any

# ---------------------------------------------------------------------------
# Pure scheduling helpers (copied from outlook_helpers.py — no COM import)
# ---------------------------------------------------------------------------

def _nth_weekday(year: int, month: int, weekday: int, n: int) -> dt.date:
    import calendar
    if n > 0:
        first = dt.date(year, month, 1)
        offset = (weekday - first.weekday()) % 7
        return first + dt.timedelta(days=offset + 7 * (n - 1))
    last_dom = calendar.monthrange(year, month)[1]
    last = dt.date(year, month, last_dom)
    offset = (last.weekday() - weekday) % 7
    return last - dt.timedelta(days=offset)


def _observed(d: dt.date) -> dt.date:
    if d.weekday() == 5:
        return d - dt.timedelta(days=1)
    if d.weekday() == 6:
        return d + dt.timedelta(days=1)
    return d


def us_federal_holidays(years) -> set:
    out: set[dt.date] = set()
    for y in years:
        for d in (dt.date(y, 1, 1), dt.date(y, 6, 19), dt.date(y, 7, 4),
                  dt.date(y, 11, 11), dt.date(y, 12, 25)):
            out.add(_observed(d))
        out.add(_nth_weekday(y, 1, 0, 3))   # MLK Day        (3rd Mon Jan)
        out.add(_nth_weekday(y, 2, 0, 3))   # Presidents Day (3rd Mon Feb)
        out.add(_nth_weekday(y, 5, 0, -1))  # Memorial Day   (last Mon May)
        out.add(_nth_weekday(y, 9, 0, 1))   # Labor Day      (1st Mon Sep)
        out.add(_nth_weekday(y, 10, 0, 2))  # Columbus Day   (2nd Mon Oct)
        out.add(_nth_weekday(y, 11, 3, 4))  # Thanksgiving   (4th Thu Nov)
    return out


WORK_SCHEDULE: dict[str, Any] = {
    "workdays": {0, 1, 2, 3, 4},
    "off_friday_anchor": dt.date(2026, 5, 29),
    "off_friday_interval_weeks": 2,
    "days_off": us_federal_holidays(range(2026, 2028)),
}


def is_off_friday(day: dt.date, schedule: dict | None = None) -> bool:
    s = schedule or WORK_SCHEDULE
    interval = s.get("off_friday_interval_weeks", 0)
    anchor = s.get("off_friday_anchor")
    if not interval or anchor is None or day.weekday() != 4:
        return False
    return (day - anchor).days % (interval * 7) == 0


def is_working_day(day: dt.date, schedule: dict | None = None) -> bool:
    s = schedule or WORK_SCHEDULE
    if day.weekday() not in s.get("workdays", {0, 1, 2, 3, 4}):
        return False
    if day in s.get("days_off", set()):
        return False
    if is_off_friday(day, s):
        return False
    return True


def working_days(start: dt.date, end: dt.date, schedule: dict | None = None) -> list[dict]:
    s = schedule or WORK_SCHEDULE
    out = []
    day = start
    while day <= end:
        if is_working_day(day, s):
            reason = "working"
        elif day.weekday() >= 5:
            reason = "weekend"
        elif is_off_friday(day, s):
            reason = "off-Friday (9/80)"
        elif day in s.get("days_off", set()):
            reason = "holiday/PTO"
        else:
            reason = "day off"
        out.append({"date": day.isoformat(), "weekday": day.strftime("%a"),
                    "working": reason == "working", "reason": reason})
        day += dt.timedelta(days=1)
    return out


# ---------------------------------------------------------------------------
# Category helpers
# ---------------------------------------------------------------------------

CATEGORY_COLORS: dict[str, int] = {
    "none": 0, "red": 1, "orange": 2, "peach": 3, "yellow": 4, "green": 5,
    "teal": 6, "olive": 7, "blue": 8, "purple": 9, "maroon": 10, "steel": 11,
    "dark steel": 12, "gray": 13, "dark gray": 14, "black": 15, "dark red": 16,
    "dark orange": 17, "dark peach": 18, "dark yellow": 19, "dark green": 20,
    "dark teal": 21, "dark olive": 22, "dark blue": 23, "dark purple": 24,
    "dark maroon": 25,
}
_COLOR_NAMES: dict[int, str] = {v: k for k, v in CATEGORY_COLORS.items()}

CATEGORY_SCHEME = [
    ("Decision needed", "red"),
    ("Waiting on reply", "orange"),
    ("Action item",     "yellow"),
    ("FYI / read later", "blue"),
    ("Delegated",       "teal"),
    ("Done",            "green"),
]


def color_name(color_int: int) -> str:
    return _COLOR_NAMES.get(color_int, str(color_int))


def _color_to_int(color: Any) -> int:
    if isinstance(color, int):
        return color
    return CATEGORY_COLORS[str(color).strip().lower()]


def _split_categories(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [c.strip() for c in value if c and str(c).strip()]
    return [c.strip() for c in str(value).split(",") if c.strip()]


def _join_categories(cats: Any) -> str:
    return ", ".join(_split_categories(cats))


def _join(recips: Any) -> str:
    if recips is None:
        return ""
    if isinstance(recips, str):
        return recips
    return "; ".join(recips)


# ---------------------------------------------------------------------------
# Fixture builder — called once at import time (fresh per subprocess)
# ---------------------------------------------------------------------------

def _ago(**kw: Any) -> str:
    return (dt.datetime.now() - dt.timedelta(**kw)).isoformat()


def _ahead(**kw: Any) -> str:
    return (dt.datetime.now() + dt.timedelta(**kw)).isoformat()


def _build_fixtures() -> dict:
    now = dt.datetime.now()

    # ── Emails ──────────────────────────────────────────────────────────────
    emails: list[dict] = [
        {
            "entry_id": "EMAIL001",
            "conversation_id": "CONV001",
            "subject": "Q2 Budget Review – Sign-off Required",
            "sender_name": "Sarah Mitchell",
            "sender_email": "sarah.mitchell@company.com",
            "to": "larry@company.com",
            "cc": "",
            "received": _ago(hours=2),
            "unread": True,
            "importance": 2,
            "has_attachments": True,
            "flagged": False,
            "categories": [],
            "attachments": ["Q2_Budget_Deck.pdf"],
            "body": (
                "Larry, please review the Q2 budget deck (attached) and provide "
                "sign-off by Friday COB. Marketing is 15% over; R&D on track. "
                "Key ask: decision on Q3 headcount add for the CNI program. Thanks, Sarah"
            ),
        },
        {
            "entry_id": "EMAIL002",
            "conversation_id": "CONV002",
            "subject": "Re: CDR Action Items – Status Update",
            "sender_name": "Tom Garza",
            "sender_email": "tom.garza@company.com",
            "to": "larry@company.com; team@company.com",
            "cc": "",
            "received": _ago(hours=26),
            "unread": True,
            "importance": 1,
            "has_attachments": False,
            "flagged": True,
            "categories": ["Waiting on reply"],
            "attachments": [],
            "body": (
                "Larry, items 3 and 7 from the CDR are still open. "
                "I need your input on the antenna trade study by end of this week. "
                "Can you send me the updated link budget spreadsheet? — Tom"
            ),
        },
        {
            "entry_id": "EMAIL003",
            "conversation_id": "CONV003",
            "subject": "All-Hands Meeting – Friday 10am",
            "sender_name": "HR Communications",
            "sender_email": "hr@company.com",
            "to": "all-staff@company.com",
            "cc": "",
            "received": _ago(hours=30),
            "unread": False,
            "importance": 1,
            "has_attachments": False,
            "flagged": False,
            "categories": ["FYI / read later"],
            "attachments": [],
            "body": (
                "All staff: quarterly all-hands is this Friday at 10am. "
                "Agenda: Q2 results, headcount plan, facility updates."
            ),
        },
        {
            "entry_id": "EMAIL004",
            "conversation_id": "CONV004",
            "subject": "Subcontract Modification – Approval Needed",
            "sender_name": "Contracts Team",
            "sender_email": "contracts@company.com",
            "to": "larry@company.com",
            "cc": "",
            "received": _ago(days=2, hours=3),
            "unread": True,
            "importance": 2,
            "has_attachments": True,
            "flagged": False,
            "categories": ["Decision needed"],
            "attachments": ["SubK_Mod_014.pdf"],
            "body": (
                "Hi Larry, the subcontract modification for Supplier XYZ is ready. "
                "Value delta: +$240K. Please sign and return by 17:00 today."
            ),
        },
        {
            "entry_id": "EMAIL005",
            "conversation_id": "CONV005",
            "subject": "Re: Antenna Test Results – 3dB Anomaly",
            "sender_name": "Dr. Priya Nair",
            "sender_email": "p.nair@company.com",
            "to": "larry@company.com",
            "cc": "tom.garza@company.com",
            "received": _ago(days=3),
            "unread": False,
            "importance": 1,
            "has_attachments": True,
            "flagged": False,
            "categories": [],
            "attachments": ["test_report_v3.xlsx"],
            "body": (
                "Larry, the 3dB drop at 14.5 GHz is reproducible. "
                "I believe it's a connector torque issue. Re-torquing and retesting Monday. "
                "Will update you after. — Priya"
            ),
        },
        {
            "entry_id": "EMAIL006",
            "conversation_id": "CONV006",
            "subject": "Travel Approval – DC Trip Jun 3–4",
            "sender_name": "Admin Portal",
            "sender_email": "admin@company.com",
            "to": "larry@company.com",
            "cc": "",
            "received": _ago(days=4),
            "unread": False,
            "importance": 1,
            "has_attachments": False,
            "flagged": False,
            "categories": ["Action item"],
            "attachments": [],
            "body": "Your travel request for Washington DC (Jun 3–4) is pending cost-center approval. Est. $1,240.",
        },
        {
            "entry_id": "EMAIL007",
            "conversation_id": "CONV007",
            "subject": "Weekly Engineering Summary",
            "sender_name": "Eng Team",
            "sender_email": "eng-updates@company.com",
            "to": "eng-leaders@company.com",
            "cc": "",
            "received": _ago(days=6),
            "unread": False,
            "importance": 0,
            "has_attachments": False,
            "flagged": False,
            "categories": ["FYI / read later"],
            "attachments": [],
            "body": "This week: CDR prep complete, lab rack delivered, two hires started. Next: EMI pre-test.",
        },
        {
            "entry_id": "EMAIL008",
            "conversation_id": "CONV001",   # reply in same thread as EMAIL001
            "subject": "Re: Q2 Budget Review – Sign-off Required",
            "sender_name": "Larry Stullich",
            "sender_email": "larry@company.com",
            "to": "sarah.mitchell@company.com",
            "cc": "",
            "received": _ago(hours=1),
            "unread": False,
            "importance": 1,
            "has_attachments": False,
            "flagged": False,
            "categories": [],
            "attachments": [],
            "body": (
                "Sarah, reviewed. Marketing overage acceptable given Q3 pipeline. "
                "Approved on headcount +2 for CNI. Looping in Finance for formal sign-off."
            ),
        },
    ]

    # ── Calendar events ──────────────────────────────────────────────────────
    def _combine(days_offset: int, hour: int, minute: int = 0) -> str:
        d = (now + dt.timedelta(days=days_offset)).date()
        return dt.datetime.combine(d, dt.time(hour, minute)).isoformat()

    events: list[dict] = [
        {
            "entry_id": "EVT001",
            "subject": "CNI Program Weekly Sync",
            "start": _combine(0, 9, 0),
            "end": _combine(0, 9, 30),
            "duration": 30,
            "location": "Teams",
            "organizer": "Larry Stullich",
            "required_attendees": "Tom Garza; Sarah Mitchell",
            "optional_attendees": "",
            "categories": [],
            "all_day": False,
            "recurring": True,
            "body": "Weekly CNI program status. Agenda: red items, schedule, action items.",
        },
        {
            "entry_id": "EVT002",
            "subject": "1:1 with Sarah Mitchell",
            "start": _combine(1, 14, 0),
            "end": _combine(1, 14, 30),
            "duration": 30,
            "location": "Sarah's Office",
            "organizer": "Sarah Mitchell",
            "required_attendees": "Larry Stullich",
            "optional_attendees": "",
            "categories": [],
            "all_day": False,
            "recurring": True,
            "body": "",   # intentionally thin agenda → triggers pulse warning
        },
        {
            "entry_id": "EVT003",
            "subject": "Q2 Budget Closeout",
            "start": _combine(2, 10, 0),
            "end": _combine(2, 11, 0),
            "duration": 60,
            "location": "Conf Room B",
            "organizer": "Sarah Mitchell",
            "required_attendees": "Larry Stullich; Finance Lead; Contracts Team",
            "optional_attendees": "HR Lead",
            "categories": [],
            "all_day": False,
            "recurring": False,
            "body": "Review and close Q2 budget actuals. Bring signed cost-center reports.",
        },
        {
            "entry_id": "EVT004",
            "subject": "Company All-Day – Office Flex Day",
            "start": _combine(3, 0, 0),
            "end": _combine(3, 23, 59),
            "duration": 1440,
            "location": "",
            "organizer": "HR Communications",
            "required_attendees": "",
            "optional_attendees": "",
            "categories": [],
            "all_day": True,
            "recurring": False,
            "body": "Work-from-home approved site-wide.",
        },
        {
            "entry_id": "EVT005",
            "subject": "Antenna Test Debrief",
            "start": _combine(4, 13, 0),
            "end": _combine(4, 14, 0),
            "duration": 60,
            "location": "RF Lab",
            "organizer": "Dr. Priya Nair",
            "required_attendees": "Larry Stullich; Dr. Priya Nair; Tom Garza",
            "optional_attendees": "Test Director",
            "categories": [],
            "all_day": False,
            "recurring": False,
            "body": "",   # thin agenda → pulse warning
        },
        {
            "entry_id": "EVT006",
            "subject": "PDR Planning – Sub-system B",
            "start": _combine(6, 9, 0),
            "end": _combine(6, 10, 30),
            "duration": 90,
            "location": "Conf Room A",
            "organizer": "Larry Stullich",
            "required_attendees": "Eng Team; Tom Garza; Dr. Priya Nair",
            "optional_attendees": "PM Office",
            "categories": [],
            "all_day": False,
            "recurring": False,
            "body": (
                "Prepare and review PDR materials for sub-system B. "
                "Bring ICD v1.2 draft, preliminary test plan, and open requirements list."
            ),
        },
    ]

    # ── Tasks ────────────────────────────────────────────────────────────────
    tasks: list[dict] = [
        {
            "entry_id": "TASK001",
            "subject": "Sign Q2 budget sign-off",
            "due": (now + dt.timedelta(days=2)).date().isoformat(),
            "body": "Received from Sarah Mitchell. See EMAIL001.",
            "categories": ["Action item"],
            "status": "open",
        },
        {
            "entry_id": "TASK002",
            "subject": "Send link budget spreadsheet to Tom Garza",
            "due": (now + dt.timedelta(days=3)).date().isoformat(),
            "body": "Requested in CDR thread. See EMAIL002.",
            "categories": ["Action item"],
            "status": "open",
        },
    ]

    # ── Master category list ──────────────────────────────────────────────────
    master_cats: list[dict] = [
        {"name": name, "color": _color_to_int(color), "color_name": color}
        for name, color in CATEGORY_SCHEME
    ]

    # ── Item store (all items by entry_id) ───────────────────────────────────
    item_store: dict[str, dict] = {}
    for e in emails:
        item_store[e["entry_id"]] = e
    for ev in events:
        item_store[ev["entry_id"]] = ev
    for t in tasks:
        item_store[t["entry_id"]] = t

    return {
        "emails": emails,
        "events": events,
        "tasks": tasks,
        "item_store": item_store,
        "master_cats": master_cats,
        "drafts": {},   # entry_id → draft dict
        "sent": [],     # list of sent entry_ids
        "folders": {
            "inbox": emails,
            "projects/cni": [emails[1], emails[4]],
            "projects/budget": [emails[0], emails[3]],
        },
    }


# Module-level state — reset happens automatically on each subprocess call
_STATE: dict = _build_fixtures()


def reset_state() -> None:
    """Re-initialize all mock state.  Call between in-process test cases."""
    global _STATE
    _STATE = _build_fixtures()


# ---------------------------------------------------------------------------
# Mail
# ---------------------------------------------------------------------------

def _message_to_dict(e: dict, include_body: bool, body_chars: int) -> dict:
    record = {k: e[k] for k in (
        "entry_id", "subject", "sender_name", "sender_email",
        "to", "received", "unread", "importance", "has_attachments",
    )}
    if include_body:
        record["body"] = (e.get("body") or "")[:body_chars]
    return record


def get_recent_messages(
    hours: int = 24,
    unread_only: bool = False,
    max_items: int = 50,
    include_body: bool = True,
    body_chars: int = 2000,
) -> list[dict]:
    cutoff = dt.datetime.now() - dt.timedelta(hours=hours)
    results = []
    for e in _STATE["emails"]:
        if dt.datetime.fromisoformat(e["received"]) < cutoff:
            continue
        if unread_only and not e["unread"]:
            continue
        results.append(_message_to_dict(e, include_body, body_chars))
        if len(results) >= max_items:
            break
    return results


def get_flagged_items(max_items: int = 50) -> list[dict]:
    results = []
    for e in _STATE["emails"]:
        if e.get("flagged"):
            results.append(_message_to_dict(e, False, 0))
            if len(results) >= max_items:
                break
    return results


def search_messages(query: str, max_items: int = 25) -> list[dict]:
    q = query.lower()
    results = []
    for e in _STATE["emails"]:
        if q in e["subject"].lower():
            results.append({
                "entry_id": e["entry_id"],
                "subject": e["subject"],
                "sender_name": e["sender_name"],
                "received": e["received"],
                "unread": e["unread"],
            })
            if len(results) >= max_items:
                break
    return results


# ---------------------------------------------------------------------------
# Calendar
# ---------------------------------------------------------------------------

def _event_to_dict(ev: dict) -> dict:
    """Return the public view of an event (entry_id included for agent use)."""
    return {k: ev.get(k) for k in (
        "entry_id",        # included so agents can reschedule/cancel by id
        "subject", "start", "end", "duration", "location",
        "organizer", "required_attendees", "optional_attendees",
        "categories", "all_day", "recurring", "body",
    )}


def get_calendar_events(days_ahead: int = 7, days_back: int = 0) -> list[dict]:
    now = dt.datetime.now()
    lo = now - dt.timedelta(days=days_back)
    hi = now + dt.timedelta(days=days_ahead)
    results = []
    for ev in _STATE["events"]:
        try:
            ev_start = dt.datetime.fromisoformat(ev["start"])
        except Exception:
            continue
        if lo <= ev_start <= hi:
            results.append(_event_to_dict(ev))
    results.sort(key=lambda e: e.get("start") or "")
    return results


def get_todays_events() -> list[dict]:
    return get_calendar_events(days_ahead=1, days_back=0)


def get_meeting_health(days_ahead: int = 1) -> list[dict]:
    out = []
    for e in get_calendar_events(days_ahead=days_ahead):
        issues = []
        if len((e.get("body") or "").strip()) < 100:
            issues.append("Thin or missing agenda")
        if issues:
            out.append({"subject": e["subject"], "start": e["start"], "issues": issues})
    return out


def get_time_spent_analytics(days_back: int = 7) -> dict:
    events = get_calendar_events(days_ahead=0, days_back=days_back)
    total_min = 0
    by_cat: dict[str, int] = {}
    for e in events:
        if e.get("all_day"):
            continue
        dur = int(e.get("duration") or 0)
        total_min += dur
        cats = e.get("categories") or ["Uncategorized"]
        for c in cats:
            by_cat[c] = by_cat.get(c, 0) + dur
    return {
        "period_days": days_back,
        "total_hours": round(total_min / 60, 1),
        "by_category_hours": {k: round(v / 60, 1) for k, v in by_cat.items()},
    }


# ---------------------------------------------------------------------------
# Folders
# ---------------------------------------------------------------------------

def list_folders(max_depth: int = 3, include_counts: bool = True) -> list[dict]:
    unread = sum(1 for e in _STATE["emails"] if e["unread"])
    folders = [
        {"name": "Inbox",       "path": "Inbox",              "depth": 0, "unread": unread},
        {"name": "Sent Items",  "path": "Sent Items",         "depth": 0, "unread": 0},
        {"name": "Drafts",      "path": "Drafts",             "depth": 0, "unread": 0},
        {"name": "Projects",    "path": "Inbox/Projects",     "depth": 1, "unread": 2},
        {"name": "CNI",         "path": "Inbox/Projects/CNI", "depth": 2, "unread": 1},
        {"name": "Budget",      "path": "Inbox/Projects/Budget", "depth": 2, "unread": 1},
    ]
    return [f for f in folders if f["depth"] <= max_depth]


def get_messages_from_folder(
    folder_path: str,
    hours: int | None = None,
    unread_only: bool = False,
    max_items: int = 50,
    include_body: bool = True,
    body_chars: int = 2000,
    base: str = "inbox",
) -> list[dict]:
    norm = folder_path.lower().replace("\\", "/").strip("/")
    emails_in_folder = _STATE["emails"]   # default: whole inbox
    for key, items in _STATE["folders"].items():
        if norm in key or key in norm:
            emails_in_folder = items
            break

    cutoff = dt.datetime.now() - dt.timedelta(hours=hours) if hours else None
    results = []
    for e in emails_in_folder:
        if cutoff and dt.datetime.fromisoformat(e["received"]) < cutoff:
            continue
        if unread_only and not e["unread"]:
            continue
        results.append(_message_to_dict(e, include_body, body_chars))
        if len(results) >= max_items:
            break
    return results


# ---------------------------------------------------------------------------
# Compose (drafts / meetings)
# ---------------------------------------------------------------------------

def create_draft_email(
    to: Any,
    subject: str,
    body: str,
    cc: Any = None,
    bcc: Any = None,
    attachments: list[str] | None = None,
    html: bool = False,
    importance: int = 1,
    send: bool = False,
    display: bool = False,
) -> dict:
    entry_id = f"DRAFT-{uuid.uuid4().hex[:8].upper()}"
    draft = {
        "entry_id": entry_id,
        "type": "draft_email",
        "to": _join(to),
        "subject": subject,
        "body": body,
        "cc": _join(cc) if cc else "",
        "status": "sent" if send else "draft",
    }
    if send:
        _STATE["sent"].append(entry_id)
    else:
        _STATE["drafts"][entry_id] = draft
    return {"status": draft["status"], "entry_id": entry_id,
            "subject": subject, "to": _join(to)}


def schedule_meeting(
    subject: str,
    start: Any,
    end: Any = None,
    duration_minutes: int = 30,
    location: str = "",
    body: str = "",
    attendees: list[str] | None = None,
    optional_attendees: list[str] | None = None,
    reminder_minutes: int | None = 15,
    busy_status: str = "busy",
    all_day: bool = False,
    send: bool = False,
    display: bool = False,
) -> dict:
    entry_id = f"EVT-{uuid.uuid4().hex[:8].upper()}"
    start_str = start if isinstance(start, str) else start.isoformat()
    start_dt = dt.datetime.fromisoformat(start_str)
    if end:
        end_dt = dt.datetime.fromisoformat(end if isinstance(end, str) else end.isoformat())
    else:
        end_dt = start_dt + dt.timedelta(minutes=duration_minutes)

    invitees = list(attendees or [])
    optionals = list(optional_attendees or [])
    sent_invites = send and bool(invitees or optionals)
    status = "invites_sent" if sent_invites else "saved"

    event_record: dict = {
        "entry_id": entry_id,
        "subject": subject,
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "duration": duration_minutes,
        "location": location,
        "organizer": "Larry Stullich",
        "required_attendees": "; ".join(invitees),
        "optional_attendees": "; ".join(optionals),
        "categories": [],
        "all_day": all_day,
        "recurring": False,
        "body": body,
    }
    _STATE["events"].append(event_record)
    _STATE["item_store"][entry_id] = event_record

    return {"status": status, "subject": subject,
            "start": start_dt.isoformat(), "attendees": invitees + optionals}


# ---------------------------------------------------------------------------
# Read one message / a whole thread
# ---------------------------------------------------------------------------

def _get_item(entry_id: str) -> dict:
    item = _STATE["item_store"].get(entry_id) or _STATE["drafts"].get(entry_id)
    if item is None:
        raise KeyError(f"No item with EntryID {entry_id!r}")
    return item


def get_message(entry_id: str, body_chars: int = 4000) -> dict:
    e = _get_item(entry_id)
    record = _message_to_dict(e, True, body_chars)
    record["cc"] = e.get("cc", "")
    record["categories"] = _split_categories(e.get("categories", []))
    record["attachments"] = list(e.get("attachments", []))
    return record


def get_conversation(entry_id: str, max_items: int = 50, body_chars: int = 1500) -> list[dict]:
    e = _get_item(entry_id)
    conv_id = e.get("conversation_id", entry_id)
    thread = [
        _message_to_dict(m, True, body_chars)
        for m in _STATE["emails"]
        if m.get("conversation_id") == conv_id
    ]
    thread.sort(key=lambda m: m.get("received") or "")
    return thread[:max_items]


def get_conversation_summary_data(entry_id: str, max_items: int = 25) -> str:
    thread = get_conversation(entry_id, max_items=max_items)
    lines = []
    for msg in thread:
        lines.append(f"SENDER: {msg['sender_name']}")
        lines.append(f"DATE: {msg['received']}")
        lines.append(f"BODY:\n{msg.get('body', '')}")
        lines.append("-" * 30)
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Reply / forward
# ---------------------------------------------------------------------------

def reply_to_message(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    html: bool = False,
    send: bool = False,
    display: bool = False,
) -> dict:
    original = _get_item(entry_id)
    reply_id = f"REPLY-{uuid.uuid4().hex[:8].upper()}"
    to = original.get("to", "") if reply_all else original.get("sender_email", "")
    draft = {
        "entry_id": reply_id,
        "type": "reply",
        "subject": f"Re: {original.get('subject', '')}",
        "to": to,
        "body": body,
        "in_reply_to": entry_id,
        "status": "sent" if send else "draft",
    }
    if send:
        _STATE["sent"].append(reply_id)
    else:
        _STATE["drafts"][reply_id] = draft
    return {"status": draft["status"], "entry_id": reply_id,
            "subject": draft["subject"], "to": to}


def forward_message(
    entry_id: str,
    to: Any,
    body: str = "",
    html: bool = False,
    send: bool = False,
    display: bool = False,
) -> dict:
    original = _get_item(entry_id)
    fwd_id = f"FWD-{uuid.uuid4().hex[:8].upper()}"
    to_str = _join(to)
    draft = {
        "entry_id": fwd_id,
        "type": "forward",
        "subject": f"Fwd: {original.get('subject', '')}",
        "to": to_str,
        "body": body,
        "forwarded_from": entry_id,
        "status": "sent" if send else "draft",
    }
    if send:
        _STATE["sent"].append(fwd_id)
    else:
        _STATE["drafts"][fwd_id] = draft
    return {"status": draft["status"], "entry_id": fwd_id,
            "subject": draft["subject"], "to": to_str}


# ---------------------------------------------------------------------------
# Triage actions
# ---------------------------------------------------------------------------

def mark_read(entry_id: str, read: bool = True) -> dict:
    item = _get_item(entry_id)
    item["unread"] = not read
    return {"status": "ok", "entry_id": entry_id, "unread": not read}


def flag_message(entry_id: str, flag: bool = True, request: str = "Follow up") -> dict:
    item = _get_item(entry_id)
    item["flagged"] = flag
    return {"status": "ok", "entry_id": entry_id, "flagged": flag}


def move_message(entry_id: str, folder_path: str, base: str = "inbox") -> dict:
    item = _get_item(entry_id)
    item["folder"] = folder_path
    return {"status": "moved", "to": folder_path}


# ---------------------------------------------------------------------------
# Categories
# ---------------------------------------------------------------------------

def get_categories(entry_id: str) -> list[str]:
    return _split_categories(_get_item(entry_id).get("categories", []))


def set_categories(entry_id: str, categories: Any) -> dict:
    item = _get_item(entry_id)
    item["categories"] = _split_categories(categories)
    return {"status": "ok", "entry_id": entry_id, "categories": item["categories"]}


def add_categories(entry_id: str, categories: Any) -> dict:
    item = _get_item(entry_id)
    current = _split_categories(item.get("categories", []))
    for c in _split_categories(categories):
        if c not in current:
            current.append(c)
    item["categories"] = current
    return {"status": "ok", "entry_id": entry_id, "categories": current}


def remove_categories(entry_id: str, categories: Any) -> dict:
    item = _get_item(entry_id)
    drop = {c.lower() for c in _split_categories(categories)}
    kept = [c for c in _split_categories(item.get("categories", [])) if c.lower() not in drop]
    item["categories"] = kept
    return {"status": "ok", "entry_id": entry_id, "categories": kept}


def list_master_categories() -> list[dict]:
    return list(_STATE["master_cats"])


def ensure_category(name: str, color: Any = None) -> dict:
    existing = {c["name"].lower() for c in _STATE["master_cats"]}
    if name.lower() in existing:
        return {"status": "exists", "name": name}
    color_int = _color_to_int(color) if color is not None else 0
    _STATE["master_cats"].append(
        {"name": name, "color": color_int, "color_name": color_name(color_int)})
    return {"status": "created", "name": name}


def init_categories(
    scheme: list | None = None,
    recolor: bool = True,
    remove_unlisted: bool = False,
) -> dict:
    scheme = CATEGORY_SCHEME if scheme is None else scheme
    wanted = {name: _color_to_int(c) for name, c in scheme}
    existing = {c["name"].lower(): c for c in _STATE["master_cats"]}

    created, recolored, removed, unchanged = [], [], [], []
    for name, want_color in wanted.items():
        cat = existing.get(name.lower())
        if cat is None:
            _STATE["master_cats"].append(
                {"name": name, "color": want_color, "color_name": color_name(want_color)})
            created.append(name)
        elif recolor and cat["color"] != want_color:
            cat["color"] = want_color
            recolored.append(name)
        else:
            unchanged.append(name)

    if remove_unlisted:
        wanted_lower = {n.lower() for n in wanted}
        _STATE["master_cats"] = [
            c for c in _STATE["master_cats"] if c["name"].lower() in wanted_lower]

    return {"created": created, "recolored": recolored,
            "removed": removed, "unchanged": unchanged}


# ---------------------------------------------------------------------------
# Free slots (same algorithm as outlook_helpers.py)
# ---------------------------------------------------------------------------

def find_free_slots(
    duration_minutes: int = 30,
    days_ahead: int = 5,
    work_start: int = 9,
    work_end: int = 17,
    slot_minutes: int = 30,
    respect_schedule: bool = True,
    ignore_all_day: bool = True,
) -> list[dict]:
    events = get_calendar_events(days_ahead=days_ahead)
    busy: list[tuple[dt.datetime, dt.datetime]] = []
    for e in events:
        if ignore_all_day and e.get("all_day"):
            continue
        try:
            busy.append((dt.datetime.fromisoformat(e["start"]),
                         dt.datetime.fromisoformat(e["end"])))
        except Exception:
            continue

    now = dt.datetime.now()
    dur = dt.timedelta(minutes=duration_minutes)
    step = dt.timedelta(minutes=slot_minutes)
    slots: list[dict] = []
    for d in range(days_ahead):
        day = (now + dt.timedelta(days=d)).date()
        if respect_schedule and not is_working_day(day):
            continue
        cursor = dt.datetime.combine(day, dt.time(work_start, 0))
        day_end = dt.datetime.combine(day, dt.time(work_end, 0))
        while cursor + dur <= day_end:
            slot_end = cursor + dur
            overlap = any(cursor < b_end and slot_end > b_start for b_start, b_end in busy)
            if cursor >= now and not overlap:
                slots.append({"start": cursor.isoformat(), "end": slot_end.isoformat()})
            cursor += step
    return slots


def auto_block_focus_time(duration_hours: int = 2, days_ahead: int = 3) -> dict:
    slots = find_free_slots(duration_minutes=duration_hours * 60, days_ahead=days_ahead)
    if not slots:
        return {"status": "error",
                "message": f"No {duration_hours}h blocks in next {days_ahead} days."}
    target = slots[0]
    return schedule_meeting(
        subject="Focus Time - Do Not Disturb",
        start=target["start"],
        end=target["end"],
        body="Reserved by AI Assistant for deep work.",
        busy_status="busy",
    )


# ---------------------------------------------------------------------------
# Tasks
# ---------------------------------------------------------------------------

def create_task(
    subject: str,
    due: Any = None,
    body: str = "",
    reminder: Any = None,
    importance: int = 1,
    categories: Any = None,
) -> dict:
    task_id = f"TASK-{uuid.uuid4().hex[:8].upper()}"
    task: dict = {
        "entry_id": task_id,
        "subject": subject,
        "body": body,
        "due": str(due) if due else None,
        "categories": _split_categories(categories or []),
        "status": "open",
    }
    _STATE["tasks"].append(task)
    _STATE["item_store"][task_id] = task
    return {"status": "created", "entry_id": task_id, "subject": subject}


# ---------------------------------------------------------------------------
# Contacts / GAL
# ---------------------------------------------------------------------------

_GAL: dict[str, dict] = {
    "sarah": {"name": "Sarah Mitchell", "email": "sarah.mitchell@company.com"},
    "sarah mitchell": {"name": "Sarah Mitchell", "email": "sarah.mitchell@company.com"},
    "tom": {"name": "Tom Garza", "email": "tom.garza@company.com"},
    "tom garza": {"name": "Tom Garza", "email": "tom.garza@company.com"},
    "priya": {"name": "Dr. Priya Nair", "email": "p.nair@company.com"},
    "priya nair": {"name": "Dr. Priya Nair", "email": "p.nair@company.com"},
    "hr": {"name": "HR Communications", "email": "hr@company.com"},
}


def resolve_recipient(name: str) -> dict:
    key = name.strip().lower()
    if key in _GAL:
        m = _GAL[key]
        return {"resolved": True, "input": name, "name": m["name"], "email": m["email"]}
    for k, m in _GAL.items():
        if key in k or k in key:
            return {"resolved": True, "input": name, "name": m["name"], "email": m["email"]}
    return {"resolved": False, "input": name}


# ---------------------------------------------------------------------------
# Approval wrapper
# ---------------------------------------------------------------------------

def propose_draft(to: Any, subject: str, body: str, **kwargs: Any) -> dict:
    kwargs["send"] = False
    result = create_draft_email(to, subject, body, **kwargs)
    result["preview"] = {"to": _join(to), "subject": subject, "body": (body or "")[:600]}
    return result


def send_draft(entry_id: str) -> dict:
    draft = _STATE["drafts"].get(entry_id)
    if draft is None:
        raise KeyError(f"No draft with EntryID {entry_id!r}")
    subject = draft.get("subject")
    del _STATE["drafts"][entry_id]
    _STATE["sent"].append(entry_id)
    return {"status": "sent", "entry_id": entry_id, "subject": subject}


# ---------------------------------------------------------------------------
# Reschedule / cancel / respond / save-attachments
# ---------------------------------------------------------------------------

def reschedule_meeting(
    entry_id: str,
    start: Any = None,
    end: Any = None,
    duration_minutes: int | None = None,
    location: str | None = None,
    send_update: bool = False,
    display: bool = False,
) -> dict:
    for ev in _STATE["events"]:
        if ev.get("entry_id") == entry_id:
            if start is not None:
                ev["start"] = start if isinstance(start, str) else start.isoformat()
            if end is not None:
                ev["end"] = end if isinstance(end, str) else end.isoformat()
            elif duration_minutes is not None:
                s_dt = dt.datetime.fromisoformat(ev["start"])
                ev["end"] = (s_dt + dt.timedelta(minutes=duration_minutes)).isoformat()
            if location is not None:
                ev["location"] = location
            return {"status": "update_sent" if send_update else "saved",
                    "subject": ev["subject"], "start": ev["start"], "end": ev["end"]}
    raise KeyError(f"No appointment with EntryID {entry_id!r}")


def cancel_meeting(entry_id: str, send_cancellation: bool = False) -> dict:
    for i, ev in enumerate(_STATE["events"]):
        if ev.get("entry_id") == entry_id:
            subject = ev["subject"]
            _STATE["events"].pop(i)
            _STATE["item_store"].pop(entry_id, None)
            return {"status": "cancelled", "subject": subject,
                    "cancellation_sent": send_cancellation}
    raise KeyError(f"No appointment with EntryID {entry_id!r}")


def respond_to_invite(
    entry_id: str,
    response: str = "accept",
    send: bool = True,
    comment: str | None = None,
) -> dict:
    for ev in _STATE["events"]:
        if ev.get("entry_id") == entry_id:
            ev["my_response"] = response
            return {"status": "responded", "response": response.lower(),
                    "sent": send, "subject": ev["subject"]}
    raise KeyError(f"No appointment with EntryID {entry_id!r}")


def save_attachments(entry_id: str, out_dir: str, name_filter: str | None = None) -> dict:
    item = _get_item(entry_id)
    attachments: list[str] = list(item.get("attachments") or [])
    if item.get("has_attachments") and not attachments:
        attachments = ["attachment_1.pdf"]  # synthetic fallback
    if name_filter:
        attachments = [a for a in attachments if name_filter.lower() in a.lower()]
    # Return what would be saved — don't write real files in the mock
    saved = [f"{out_dir}\\{a}" for a in attachments]
    return {"status": "ok", "count": len(saved), "files": saved}
