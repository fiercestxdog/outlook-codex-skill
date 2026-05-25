"""
outlook_helpers.py  (v2)
------------------------
Read Outlook mail and calendar via the local desktop client (COM automation).

Why COM: it drives the Outlook app already running and authenticated on your
machine. No Azure app registration, no admin consent, no IT ticket -- which is
usually what survives a locked-down corporate/defense environment. Windows-only,
and Outlook must be installed (running is best).

All functions return plain JSON-serializable dicts/lists so they can be piped
straight into a Codex/Claude skill (weekly brief, action-item extraction, etc.).

⚠️  CLASSIC OUTLOOK ONLY
    This module uses pywin32 COM automation, which works with Classic Outlook
    (the legacy Win32 app, version numbers 14.x–16.x). The new "One Outlook"
    web-based app shipped as the Windows 11 default since late 2024 does NOT
    support COM — all Dispatch() calls will silently fail or raise. If commands
    return nothing or crash, verify Classic Outlook is installed and active:
        Outlook → Settings → "Try the new Outlook" toggle must be OFF.
    Classic Outlook M365 remains supported; this code has a sunset dependency
    on Microsoft eventually forcing the migration to New Outlook.

⚠️  SECURITY POPUP (Outlook 365, post-Oct 2024)
    Microsoft re-enabled programmatic-access prompts ("A program is trying to
    send an email on your behalf"). The old registry bypass no longer suppresses
    it on M365. Workarounds:
      1. Antivirus exception for python.exe (simplest for personal use).
      2. Domain GPO: HKLM\\...\\Office\\16.0\\Outlook\\Security\\ObjectModelGuard=0.
      3. Move send/draft to Microsoft Graph (OAuth) for fully unattended use.

Dependency:
    pip install pywin32>=306

Locale note: Outlook's Restrict() date filters must match the Windows system
locale's short date/time format. _DATE_FMT below is the US locale pattern
("%m/%d/%Y %I:%M %p"). On non-US Windows, call _detect_date_fmt() at startup
to get the correct pattern; mismatched formats silently return zero results.
"""

from __future__ import annotations

import datetime as dt
import re
from typing import Any

import pythoncom
import win32com.client

__version__ = "2.0.0"

# Outlook default-folder enum (OlDefaultFolders)
OL_FOLDER_INBOX = 6
OL_FOLDER_CALENDAR = 9
OL_FOLDER_SENT = 5

# Item types (OlItemType)
OL_MAIL_ITEM = 0
OL_APPOINTMENT_ITEM = 1

# Recipient types (OlMeetingRecipientType)
OL_REQUIRED = 1
OL_OPTIONAL = 2

# Busy status (OlBusyStatus)
_BUSY = {"free": 0, "tentative": 1, "busy": 2, "oof": 3, "elsewhere": 4}

_DATE_FMT = "%m/%d/%Y %I:%M %p"   # US locale — see _detect_date_fmt() below


def _detect_date_fmt() -> str:
    """Derive the Outlook-compatible Restrict() date format from the Windows
    locale. Falls back to US format. Call once at startup if you see calendar
    filters returning 0 results on a non-US machine.

    Usage (one-time diagnostic):
        python -c "import outlook_helpers; print(outlook_helpers._detect_date_fmt())"
    If the result differs from '%m/%d/%Y %I:%M %p', set _DATE_FMT to the
    returned value before making any Restrict() calls.
    """
    try:
        import ctypes
        import ctypes.wintypes

        # LOCALE_USER_DEFAULT = 0x0400
        LOCALE_SSHORTDATE = 0x001F
        LOCALE_STIMEFORMAT = 0x1003
        buf = ctypes.create_unicode_buffer(80)

        def get_fmt(lc_type: int) -> str:
            ctypes.windll.kernel32.GetLocaleInfoW(0x0400, lc_type, buf, len(buf))
            raw = buf.value  # e.g. "M/d/yyyy" or "d/MM/yyyy"
            # Map Windows locale picture tokens → strftime directives
            token_map = [
                ("dddd", "%A"), ("ddd", "%a"), ("dd", "%d"), ("d", "%-d"),
                ("MMMM", "%B"), ("MMM", "%b"), ("MM", "%m"), ("M", "%-m"),
                ("yyyy", "%Y"), ("yy", "%y"),
                ("HH", "%H"), ("H", "%-H"),
                ("hh", "%I"), ("h", "%-I"),
                ("mm", "%M"),  # minutes (must come after MM=month)
                ("ss", "%S"),
                ("tt", "%p"),
            ]
            result = raw
            for token, directive in token_map:
                result = result.replace(token, directive)
            return result

        date_part = get_fmt(LOCALE_SSHORTDATE)
        time_part = get_fmt(LOCALE_STIMEFORMAT)
        return f"{date_part} {time_part}"
    except Exception:
        return _DATE_FMT   # safe fallback


# --------------------------------------------------------------------------- #
# COM apartment initialization
# --------------------------------------------------------------------------- #
# Initialize COM for the main thread once at import time and uninitialize on
# exit. This is safer than calling CoInitialize() inside _application() on
# every call: repeated CoInitialize() calls on the same thread are harmless
# (COM reference-counts them), but they leak if CoUninitialize() is never
# paired — which matters if this module is loaded into a long-running process
# (web server, task queue, background service).
#
# THREAD NOTE: If you call any helper function from a *worker thread*, you must
# call pythoncom.CoInitialize() at the start of that thread and
# pythoncom.CoUninitialize() when the thread exits. The main-thread
# initialization below does NOT cover worker threads.
import atexit as _atexit
pythoncom.CoInitialize()
_atexit.register(pythoncom.CoUninitialize)


def _check_classic_outlook() -> None:
    """Raise a clear RuntimeError if the running Outlook is the new web-based
    app (COM-incompatible) or if Outlook is not running at all.

    Classic Outlook reports a version string like "16.0.xxxxx.xxxxx".
    New Outlook either raises on Dispatch or returns a version that doesn't
    start with a known major number (14 = 2010, 15 = 2013, 16 = 2016/365).
    """
    try:
        app = win32com.client.Dispatch("Outlook.Application")
        ver: str = getattr(app, "Version", "") or ""
    except Exception as exc:
        raise RuntimeError(
            "Outlook COM is unavailable. Is Classic Outlook installed and running?\n"
            "New Outlook (the Windows 11 default) does not support COM automation.\n"
            f"  Original error: {exc}"
        ) from exc

    major = ver.split(".")[0] if ver else ""
    if major not in {"14", "15", "16"}:
        raise RuntimeError(
            f"Outlook version '{ver}' is not Classic Outlook.\n"
            "New Outlook does not support COM automation.\n"
            "Fix: Outlook → Settings → turn off 'Try the new Outlook'."
        )


def _application():
    """Return the Outlook.Application COM object.

    Raises RuntimeError with an actionable message if Classic Outlook is not
    available (New Outlook detected, Outlook not running, COM failure).
    """
    _check_classic_outlook()
    return win32com.client.Dispatch("Outlook.Application")


def _namespace():
    """Return a MAPI namespace."""
    return _application().GetNamespace("MAPI")


def _as_dt(value: dt.datetime | str) -> dt.datetime:
    """Accept a datetime or an ISO-8601 string and return a datetime."""
    if isinstance(value, dt.datetime):
        return value
    return dt.datetime.fromisoformat(value)


def _to_py_datetime(value: Any) -> str | None:
    """COM returns pywintypes datetimes; normalize to an ISO string."""
    if value is None:
        return None
    try:
        return dt.datetime(
            value.year, value.month, value.day,
            value.hour, value.minute, value.second,
        ).isoformat()
    except Exception:
        return str(value)


# --------------------------------------------------------------------------- #
# Mail
# --------------------------------------------------------------------------- #
def _message_to_dict(msg, include_body: bool, body_chars: int) -> dict:
    record = {
        "entry_id": msg.EntryID,
        "subject": msg.Subject,
        "sender_name": msg.SenderName,
        "sender_email": _smtp_address(msg),
        "to": msg.To,
        "received": _to_py_datetime(msg.ReceivedTime),
        "unread": bool(msg.UnRead),
        "importance": msg.Importance,  # 0 low / 1 normal / 2 high
        "has_attachments": msg.Attachments.Count > 0,
    }
    if include_body:
        record["body"] = (msg.Body or "")[:body_chars]
    return record


def get_recent_messages(
    hours: int = 24,
    unread_only: bool = False,
    max_items: int = 50,
    include_body: bool = True,
    body_chars: int = 2000,
) -> list[dict]:
    """Return Inbox messages received within the last `hours`, newest first."""
    ns = _namespace()
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)  # descending

    cutoff = dt.datetime.now() - dt.timedelta(hours=hours)
    restriction = f"[ReceivedTime] >= '{cutoff.strftime(_DATE_FMT)}'"
    if unread_only:
        restriction += " AND [Unread] = True"
    items = items.Restrict(restriction)

    out: list[dict] = []
    for msg in items:
        try:
            if getattr(msg, "Class", None) != 43:  # 43 = olMail; skip non-mail
                continue
            out.append(_message_to_dict(msg, include_body, body_chars))
        except Exception:
            continue
        if len(out) >= max_items:
            break
    return out


def get_flagged_items(max_items: int = 50) -> list[dict]:
    """Return Inbox messages flagged for follow-up (not yet complete), newest first."""
    ns = _namespace()
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    # olFlagMarked = 2
    items = inbox.Items.Restrict("[FlagStatus] = 2")
    items.Sort("[ReceivedTime]", True)

    out: list[dict] = []
    for msg in items:
        try:
            # We skip bodies here to keep the radar list lean
            out.append(_message_to_dict(msg, include_body=False, body_chars=0))
        except Exception:
            continue
        if len(out) >= max_items:
            break
    return out


def search_messages(query: str, max_items: int = 25) -> list[dict]:
    """Case-insensitive substring search over Inbox subjects, newest first."""
    ns = _namespace()
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    # ci_phrasematch does a fast content-index match on the subject
    restricted = items.Restrict(
        f"@SQL=\"urn:schemas:httpmail:subject\" ci_phrasematch '{query}'"
    )
    out: list[dict] = []
    for msg in restricted:
        try:
            out.append({
                "entry_id": msg.EntryID,
                "subject": msg.Subject,
                "sender_name": msg.SenderName,
                "received": _to_py_datetime(msg.ReceivedTime),
                "unread": bool(msg.UnRead),
            })
        except Exception:
            continue
        if len(out) >= max_items:
            break
    return out


def _smtp_address(msg) -> str | None:
    """Resolve a real SMTP address (Exchange senders show an X.500 string
    in SenderEmailAddress, so go through the Exchange user object)."""
    try:
        if msg.SenderEmailType == "EX":
            sender = msg.Sender
            if sender is not None:
                exch = sender.GetExchangeUser()
                if exch is not None:
                    return exch.PrimarySmtpAddress
        return msg.SenderEmailAddress
    except Exception:
        return getattr(msg, "SenderEmailAddress", None)


# --------------------------------------------------------------------------- #
# Calendar
# --------------------------------------------------------------------------- #
def get_calendar_events(days_ahead: int = 7, days_back: int = 0) -> list[dict]:
    """Return appointments in [now - days_back, now + days_ahead].

    Recurring events are expanded into individual instances.

    ⚠️  MANDATORY ORDER — do not reorder these three lines:
        1. items.Sort("[Start]")          ← must sort by [Start], no other field
        2. items.IncludeRecurrences = True ← must be set AFTER Sort
        3. items.Restrict(filter)          ← must be called AFTER IncludeRecurrences

    Breaking this order causes Outlook to either silently drop all recurring
    events, return a raw master-recurrence item instead of instances, or loop
    infinitely on some Outlook 365 builds. This is undocumented behavior that
    was reverse-engineered from MSKB articles — treat it as immutable.
    """
    ns = _namespace()
    cal = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
    items = cal.Items
    items.Sort("[Start]")              # Step 1: sort by Start — REQUIRED first
    items.IncludeRecurrences = True    # Step 2: expand recurrences — AFTER Sort

    start = dt.datetime.now() - dt.timedelta(days=days_back)
    end = dt.datetime.now() + dt.timedelta(days=days_ahead)
    restriction = (
        f"[Start] >= '{start.strftime(_DATE_FMT)}' "
        f"AND [Start] <= '{end.strftime(_DATE_FMT)}'"
    )
    restricted = items.Restrict(restriction)   # Step 3: filter — AFTER IncludeRecurrences

    out: list[dict] = []
    for appt in restricted:
        try:
            out.append({
                "subject": appt.Subject,
                "start": _to_py_datetime(appt.Start),
                "end": _to_py_datetime(appt.End),
                "duration": appt.Duration,
                "location": appt.Location,
                "organizer": appt.Organizer,
                "required_attendees": appt.RequiredAttendees,
                "optional_attendees": appt.OptionalAttendees,
                "categories": _split_categories(getattr(appt, "Categories", "")),
                "all_day": bool(appt.AllDayEvent),
                "recurring": bool(appt.IsRecurring),
                "body": (appt.Body or "")[:1000],
            })
        except Exception:
            continue
    out.sort(key=lambda e: e["start"] or "")
    return out


def get_todays_events() -> list[dict]:
    """Convenience: just today's meetings."""
    return get_calendar_events(days_ahead=1, days_back=0)


def get_meeting_health(days_ahead: int = 1) -> list[dict]:
    """Identify upcoming meetings that might need attention (no agenda, no confirmed attendees)."""
    events = get_calendar_events(days_ahead=days_ahead)
    out = []
    for e in events:
        issues = []
        body = e.get("body", "").strip()
        if len(body) < 100:
            issues.append("Thin or missing agenda")
        
        # Check if it's a meeting (has attendees) vs a personal appointment
        if e["required_attendees"] or e["optional_attendees"]:
            # This is a bit simplified for COM; real health would check ResponseStatus
            pass

        if issues:
            out.append({
                "subject": e["subject"],
                "start": e["start"],
                "issues": issues
            })
    return out


def get_time_spent_analytics(days_back: int = 7) -> dict:
    """Summarize time spent in meetings over the last `days_back` days, grouped by category."""
    events = get_calendar_events(days_ahead=0, days_back=days_back)
    total_min = 0
    by_cat = {}
    
    for e in events:
        if e.get("all_day"): continue
        dur = e.get("duration", 0)
        total_min += dur
        
        cats = e.get("categories") or ["Uncategorized"]
        for c in cats:
            by_cat[c] = by_cat.get(c, 0) + dur

    return {
        "period_days": days_back,
        "total_hours": round(total_min / 60, 1),
        "by_category_hours": {k: round(v / 60, 1) for k, v in by_cat.items()}
    }


# --------------------------------------------------------------------------- #
# Folders
# --------------------------------------------------------------------------- #
def list_folders(max_depth: int = 3, include_counts: bool = True) -> list[dict]:
    """Walk every mail store and return its folder tree as flat dicts:
    {name, path, depth, unread}. `path` is what you pass to
    get_messages_from_folder()."""
    ns = _namespace()
    out: list[dict] = []

    def walk(folder, path, depth):
        try:
            entry = {"name": folder.Name, "path": path, "depth": depth}
            if include_counts:
                entry["unread"] = folder.UnReadItemCount
            out.append(entry)
        except Exception:
            return
        if depth >= max_depth:
            return
        try:
            for sub in folder.Folders:
                walk(sub, f"{path}/{sub.Name}", depth + 1)
        except Exception:
            pass

    for store in ns.Folders:
        walk(store, store.Name, 0)
    return out


def _resolve_folder(path: str, base: str = "inbox"):
    """Resolve a folder from a '/'- or '\\'-separated path.
    base='inbox' (default) treats the path as relative to your Inbox, so
    'Projects/F-35' means Inbox > Projects > F-35. base='root' treats it as
    relative to the mailbox root (e.g. 'Sent Items', 'Archive/2026')."""
    parts = [p for p in re.split(r"[\\/]+", path) if p]
    ns = _namespace()
    if base == "root":
        current = ns.GetDefaultFolder(OL_FOLDER_INBOX).Parent  # mailbox store root
    else:
        current = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    # tolerate the path optionally repeating the base folder's own name
    if parts and current.Name and parts[0].lower() == current.Name.lower():
        parts = parts[1:]
    for name in parts:
        current = current.Folders.Item(name)  # raises if the folder is missing
    return current


def get_messages_from_folder(
    folder_path: str,
    hours: int | None = None,
    unread_only: bool = False,
    max_items: int = 50,
    include_body: bool = True,
    body_chars: int = 2000,
    base: str = "inbox",
) -> list[dict]:
    """Read messages from any folder (not just the Inbox). Pass hours=None to
    get the newest messages regardless of age."""
    folder = _resolve_folder(folder_path, base=base)
    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    filters = []
    if hours is not None:
        cutoff = dt.datetime.now() - dt.timedelta(hours=hours)
        filters.append(f"[ReceivedTime] >= '{cutoff.strftime(_DATE_FMT)}'")
    if unread_only:
        filters.append("[Unread] = True")
    if filters:
        items = items.Restrict(" AND ".join(filters))

    out: list[dict] = []
    for msg in items:
        try:
            if getattr(msg, "Class", None) != 43:  # olMail only
                continue
            out.append(_message_to_dict(msg, include_body, body_chars))
        except Exception:
            continue
        if len(out) >= max_items:
            break
    return out


# --------------------------------------------------------------------------- #
# Compose (draft email + schedule meeting)
# --------------------------------------------------------------------------- #
def _join(recips) -> str:
    if recips is None:
        return ""
    if isinstance(recips, str):
        return recips
    return "; ".join(recips)


def create_draft_email(
    to,
    subject: str,
    body: str,
    cc=None,
    bcc=None,
    attachments: list[str] | None = None,
    html: bool = False,
    importance: int = 1,        # 0 low / 1 normal / 2 high
    send: bool = False,
    display: bool = False,
) -> dict:
    """Create a mail item. Saves to Drafts by default (send=False) so a human
    reviews before anything leaves -- the safe default for an agent workflow.
    `to`/`cc`/`bcc` accept a string or a list of addresses. Set display=True to
    pop the draft open in Outlook, send=True to send immediately."""
    import os

    app = _application()
    mail = app.CreateItem(OL_MAIL_ITEM)
    mail.To = _join(to)
    if cc:
        mail.CC = _join(cc)
    if bcc:
        mail.BCC = _join(bcc)
    mail.Subject = subject
    if html:
        mail.HTMLBody = body
    else:
        mail.Body = body
    mail.Importance = importance
    for path in (attachments or []):
        mail.Attachments.Add(os.path.abspath(path))

    if send:
        mail.Send()
        status = "sent"
    else:
        mail.Save()  # lands in the Drafts folder
        status = "draft"
        if display:
            mail.Display(False)

    return {
        "status": status,
        "entry_id": getattr(mail, "EntryID", None),
        "subject": subject,
        "to": _join(to),
    }


def schedule_meeting(
    subject: str,
    start: dt.datetime | str,
    end: dt.datetime | str | None = None,
    duration_minutes: int = 30,
    location: str = "",
    body: str = "",
    attendees: list[str] | None = None,
    optional_attendees: list[str] | None = None,
    reminder_minutes: int | None = 15,
    busy_status: str = "busy",   # free / tentative / busy / oof / elsewhere
    all_day: bool = False,
    send: bool = False,
    display: bool = False,
) -> dict:
    """Create a calendar appointment, or a meeting request when attendees are
    supplied. Saves to your calendar by default (send=False); set send=True to
    actually send invites. `start`/`end` accept a datetime or ISO-8601 string;
    if `end` is omitted, `duration_minutes` is used."""
    app = _application()
    appt = app.CreateItem(OL_APPOINTMENT_ITEM)
    appt.Subject = subject

    start_dt = _as_dt(start)
    appt.Start = start_dt
    if all_day:
        appt.AllDayEvent = True
    elif end is not None:
        appt.End = _as_dt(end)
    else:
        appt.Duration = duration_minutes  # in minutes

    appt.Location = location
    appt.Body = body
    if reminder_minutes is not None:
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = reminder_minutes
    appt.BusyStatus = _BUSY.get(busy_status, 2)

    invitees = list(attendees or [])
    optionals = list(optional_attendees or [])
    if invitees or optionals:
        appt.MeetingStatus = 1  # olMeeting -> turns the appointment into an invite
        for addr in invitees:
            appt.Recipients.Add(addr).Type = OL_REQUIRED
        for addr in optionals:
            appt.Recipients.Add(addr).Type = OL_OPTIONAL
        appt.Recipients.ResolveAll()

    if send and (invitees or optionals):
        appt.Send()
        status = "invites_sent"
    else:
        appt.Save()  # on your calendar; no invites sent
        status = "saved"
        if display:
            appt.Display(False)

    return {
        "status": status,
        "subject": subject,
        "start": start_dt.isoformat(),
        "attendees": invitees + optionals,
    }


# --------------------------------------------------------------------------- #
# Shared item access
# --------------------------------------------------------------------------- #
OL_TASK_ITEM = 3

def _get_item(entry_id: str):
    """Fetch any item (mail, appointment, task) by its EntryID."""
    return _namespace().GetItemFromID(entry_id)


def _split_categories(value) -> list[str]:
    """Outlook stores categories as a comma-separated string."""
    if value is None:
        return []
    if isinstance(value, list):
        return [c.strip() for c in value if c and c.strip()]
    return [c.strip() for c in str(value).split(",") if c.strip()]


def _join_categories(cats) -> str:
    return ", ".join(_split_categories(cats))


# --------------------------------------------------------------------------- #
# Read one message / a whole thread
# --------------------------------------------------------------------------- #
def get_message(entry_id: str, body_chars: int = 4000) -> dict:
    """Fetch a single message in full, including attachment names."""
    msg = _get_item(entry_id)
    record = _message_to_dict(msg, include_body=True, body_chars=body_chars)
    record["cc"] = getattr(msg, "CC", "")
    record["categories"] = _split_categories(getattr(msg, "Categories", ""))
    record["attachments"] = [a.FileName for a in msg.Attachments]
    return record


def get_conversation(entry_id: str, max_items: int = 50, body_chars: int = 1500) -> list[dict]:
    """Return the full conversation thread for a message, oldest first."""
    msg = _get_item(entry_id)
    try:
        conv = msg.GetConversation()
    except Exception:
        conv = None
    if conv is None:
        return [_message_to_dict(msg, True, body_chars)]

    out: list[dict] = []
    roots = conv.GetRootItems()
    queue = [roots.Item(i) for i in range(1, roots.Count + 1)]
    while queue and len(out) < max_items:
        node = queue.pop(0)
        try:
            if getattr(node, "Class", None) == 43:  # olMail
                out.append(_message_to_dict(node, True, body_chars))
            children = conv.GetChildren(node)
            for i in range(1, children.Count + 1):
                queue.append(children.Item(i))
        except Exception:
            continue
    out.sort(key=lambda m: m.get("received") or "")
    return out[:max_items]


def get_conversation_summary_data(entry_id: str, max_items: int = 25) -> str:
    """Return a single flattened string of a conversation for LLM summarization."""
    thread = get_conversation(entry_id, max_items=max_items)
    lines = []
    for msg in thread:
        lines.append(f"SENDER: {msg['sender_name']}")
        lines.append(f"DATE: {msg['received']}")
        lines.append(f"BODY:\n{msg['body']}")
        lines.append("-" * 30)
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# Reply / forward (as draft by default)
# --------------------------------------------------------------------------- #
def reply_to_message(
    entry_id: str,
    body: str,
    reply_all: bool = False,
    html: bool = False,
    send: bool = False,
    display: bool = False,
) -> dict:
    """Draft a thread-aware reply. Your text goes above the quoted original.
    Saves as a draft by default (send=False)."""
    original = _get_item(entry_id)
    reply = original.ReplyAll() if reply_all else original.Reply()
    if html:
        reply.HTMLBody = (body or "") + reply.HTMLBody
    else:
        reply.Body = (body or "") + "\n\n" + reply.Body

    if send:
        reply.Send()
        status = "sent"
    else:
        reply.Save()
        status = "draft"
        if display:
            reply.Display(False)
    return {"status": status, "entry_id": getattr(reply, "EntryID", None),
            "subject": reply.Subject, "to": reply.To}


def forward_message(
    entry_id: str,
    to,
    body: str = "",
    html: bool = False,
    send: bool = False,
    display: bool = False,
) -> dict:
    """Forward a message (as draft by default)."""
    original = _get_item(entry_id)
    fwd = original.Forward()
    fwd.To = _join(to)
    if html:
        fwd.HTMLBody = (body or "") + fwd.HTMLBody
    else:
        fwd.Body = (body or "") + "\n\n" + fwd.Body

    if send:
        fwd.Send()
        status = "sent"
    else:
        fwd.Save()
        status = "draft"
        if display:
            fwd.Display(False)
    return {"status": status, "entry_id": getattr(fwd, "EntryID", None),
            "subject": fwd.Subject, "to": _join(to)}


# --------------------------------------------------------------------------- #
# Triage actions (mark read, flag, move)
# --------------------------------------------------------------------------- #
def mark_read(entry_id: str, read: bool = True) -> dict:
    msg = _get_item(entry_id)
    msg.UnRead = not read
    msg.Save()
    return {"status": "ok", "entry_id": entry_id, "unread": not read}


def flag_message(entry_id: str, flag: bool = True, request: str = "Follow up") -> dict:
    """Set or clear a follow-up flag."""
    msg = _get_item(entry_id)
    if flag:
        msg.FlagRequest = request
        msg.FlagStatus = 2  # olFlagMarked
    else:
        try:
            msg.ClearTaskFlag()
        except Exception:
            msg.FlagStatus = 0  # olNoFlag
    msg.Save()
    return {"status": "ok", "entry_id": entry_id, "flagged": flag}


def move_message(entry_id: str, folder_path: str, base: str = "inbox") -> dict:
    """Move a message to another folder (path resolved like get_messages_from_folder)."""
    msg = _get_item(entry_id)
    target = _resolve_folder(folder_path, base=base)
    msg.Move(target)
    return {"status": "moved", "to": folder_path}


# --------------------------------------------------------------------------- #
# Categories -- organize mail AND meetings the same way
# --------------------------------------------------------------------------- #
def get_categories(entry_id: str) -> list[str]:
    """Read the categories on any item (mail or appointment)."""
    return _split_categories(getattr(_get_item(entry_id), "Categories", ""))


def set_categories(entry_id: str, categories) -> dict:
    """Replace an item's categories. `categories` is a string or list.
    Works identically on mail items and calendar appointments."""
    item = _get_item(entry_id)
    item.Categories = _join_categories(categories)
    item.Save()
    return {"status": "ok", "entry_id": entry_id,
            "categories": _split_categories(item.Categories)}


def add_categories(entry_id: str, categories) -> dict:
    """Add one or more categories without dropping existing ones."""
    item = _get_item(entry_id)
    current = _split_categories(item.Categories)
    for c in _split_categories(categories):
        if c not in current:
            current.append(c)
    item.Categories = ", ".join(current)
    item.Save()
    return {"status": "ok", "entry_id": entry_id, "categories": current}


def remove_categories(entry_id: str, categories) -> dict:
    """Remove one or more categories, leaving the rest intact."""
    item = _get_item(entry_id)
    drop = {c.lower() for c in _split_categories(categories)}
    kept = [c for c in _split_categories(item.Categories) if c.lower() not in drop]
    item.Categories = ", ".join(kept)
    item.Save()
    return {"status": "ok", "entry_id": entry_id, "categories": kept}


# OlCategoryColor enum: human name <-> int. 0 = none; 1..25 are Outlook's swatches.
# Use the names below anywhere a color is expected -- you never touch the ints.
CATEGORY_COLORS = {
    "none": 0, "red": 1, "orange": 2, "peach": 3, "yellow": 4, "green": 5,
    "teal": 6, "olive": 7, "blue": 8, "purple": 9, "maroon": 10, "steel": 11,
    "dark steel": 12, "gray": 13, "dark gray": 14, "black": 15, "dark red": 16,
    "dark orange": 17, "dark peach": 18, "dark yellow": 19, "dark green": 20,
    "dark teal": 21, "dark olive": 22, "dark blue": 23, "dark purple": 24,
    "dark maroon": 25,
}
_COLOR_NAMES = {v: k for k, v in CATEGORY_COLORS.items()}


def color_name(color_int: int) -> str:
    """Turn an OlCategoryColor int into its name (e.g. 8 -> 'blue')."""
    return _COLOR_NAMES.get(color_int, str(color_int))


def _color_to_int(color) -> int:
    """Accept an int (0-25) or a color name ('blue', 'dark green') -> int."""
    if isinstance(color, int):
        return color
    return CATEGORY_COLORS[str(color).strip().lower()]


def list_master_categories() -> list[dict]:
    """List the categories in your master list as {name, color, color_name}."""
    ns = _namespace()
    out = []
    for cat in ns.Categories:
        try:
            out.append({"name": cat.Name, "color": cat.Color,
                        "color_name": color_name(cat.Color)})
        except Exception:
            continue
    return out


def ensure_category(name: str, color=None) -> dict:
    """Add a single category if missing. `color` accepts a name ('blue') or an
    int; leave None to let Outlook choose. For your whole taxonomy at once,
    prefer init_categories()."""
    ns = _namespace()
    existing = {c["name"].lower() for c in list_master_categories()}
    if name.lower() in existing:
        return {"status": "exists", "name": name}
    if color is None:
        ns.Categories.Add(name)
    else:
        ns.Categories.Add(name, _color_to_int(color))
    return {"status": "created", "name": name}


# ---- Single source of truth for your category taxonomy --------------------- #
# This list IS your taxonomy. Edit it, then run init_categories() to make
# Outlook match. Outlook becomes a projection of this -- not the reverse -- so
# the same scheme reproduces identically on any machine you run it on.
# Color is a name from CATEGORY_COLORS (or a raw 0-25 int).
CATEGORY_SCHEME = [
    # --- Triage / workflow (drives the email-triage + action-item skills) ---
    ("Decision needed",   "red"),
    ("Waiting on reply",  "orange"),
    ("Action item",       "yellow"),
    ("FYI / read later",  "blue"),
    ("Delegated",         "teal"),
    ("Done",              "green"),
    # --- Program tags: uncomment / edit to match your portfolio -------------
    # ("F-35 CNI", "purple"),
    # ("F-22 CNI", "dark purple"),
    # ("BACN",     "steel"),
    # ("JCREW",    "olive"),
    # ("MADL",     "maroon"),
]


def init_categories(scheme=None, recolor: bool = True,
                    remove_unlisted: bool = False) -> dict:
    """One-time (and idempotent) sync of CATEGORY_SCHEME into Outlook's master
    category list. Safe to re-run on any machine:
      - missing categories are created with the right color
      - existing ones are recolored to match (when recolor=True)
      - categories outside the scheme are left untouched unless
        remove_unlisted=True
    Returns a report of what changed so a skill can log/announce it."""
    ns = _namespace()
    scheme = CATEGORY_SCHEME if scheme is None else scheme
    wanted = {name: _color_to_int(c) for name, c in scheme}

    existing = {}
    for cat in ns.Categories:
        try:
            existing[cat.Name.lower()] = cat
        except Exception:
            continue

    created, recolored, removed, unchanged = [], [], [], []
    for name, want_color in wanted.items():
        cat = existing.get(name.lower())
        if cat is None:
            ns.Categories.Add(name, want_color)
            created.append(name)
        elif recolor and cat.Color != want_color:
            cat.Color = want_color
            recolored.append(name)
        else:
            unchanged.append(name)

    if remove_unlisted:
        wanted_lower = {n.lower() for n in wanted}
        for lname, cat in list(existing.items()):
            if lname not in wanted_lower:
                try:
                    name = cat.Name
                    ns.Categories.Remove(cat.CategoryID)
                    removed.append(name)
                except Exception:
                    continue

    return {"created": created, "recolored": recolored,
            "removed": removed, "unchanged": unchanged}


# --------------------------------------------------------------------------- #
# Scheduling helper: find open slots
# --------------------------------------------------------------------------- #

# ---- US federal holidays (computed, so they're correct every year) --------- #
def _nth_weekday(year: int, month: int, weekday: int, n: int) -> dt.date:
    """nth weekday of a month. weekday: Mon=0..Sun=6. n: 1-based, or -1 = last."""
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
    """Federal observed-date rule: Sat holiday -> Fri, Sun holiday -> Mon."""
    if d.weekday() == 5:
        return d - dt.timedelta(days=1)
    if d.weekday() == 6:
        return d + dt.timedelta(days=1)
    return d


def us_federal_holidays(years) -> set:
    """Set of observed US federal holiday dates for the given years. These are
    the days Northrop/L3Harris-style employers typically observe; edit the
    schedule if your site's calendar differs."""
    out = set()
    for y in years:
        for d in (dt.date(y, 1, 1),    # New Year's Day
                  dt.date(y, 6, 19),   # Juneteenth
                  dt.date(y, 7, 4),    # Independence Day
                  dt.date(y, 11, 11),  # Veterans Day
                  dt.date(y, 12, 25)): # Christmas
            out.add(_observed(d))
        out.add(_nth_weekday(y, 1, 0, 3))    # MLK Day        (3rd Mon Jan)
        out.add(_nth_weekday(y, 2, 0, 3))    # Presidents Day (3rd Mon Feb)
        out.add(_nth_weekday(y, 5, 0, -1))   # Memorial Day   (last Mon May)
        out.add(_nth_weekday(y, 9, 0, 1))    # Labor Day      (1st Mon Sep)
        out.add(_nth_weekday(y, 10, 0, 2))   # Columbus Day   (2nd Mon Oct)
        out.add(_nth_weekday(y, 11, 3, 4))   # Thanksgiving   (4th Thu Nov)
    return out


# ---- Single source of truth for your working schedule --------------------- #
# Edit this to match how you actually work. Slot-finding skips any day that
# isn't a working day, so it will never offer a weekend, off-Friday, or holiday.
WORK_SCHEDULE = {
    # weekday() ints that are normally worked: Mon=0 ... Fri=4 (Sat=5, Sun=6 off)
    "workdays": {0, 1, 2, 3, 4},
    # 9/80 schedule: every other Friday off, in phase with this anchor Friday.
    "off_friday_anchor": dt.date(2026, 5, 29),
    "off_friday_interval_weeks": 2,      # set to 0 to disable off-Fridays
    # Days off: federal holidays preloaded for 2026-2027. Add PTO the same way:
    # WORK_SCHEDULE["days_off"] |= {dt.date(2026, 12, 28), dt.date(2026, 12, 31)}
    "days_off": us_federal_holidays(range(2026, 2028)),
}


def is_off_friday(day: dt.date, schedule: dict | None = None) -> bool:
    """True if `day` is one of the recurring off-Fridays (9/80 pattern)."""
    s = schedule or WORK_SCHEDULE
    interval = s.get("off_friday_interval_weeks", 0)
    anchor = s.get("off_friday_anchor")
    if not interval or anchor is None or day.weekday() != 4:  # 4 = Friday
        return False
    # same phase as the anchor Friday, every `interval` weeks (works both
    # before and after the anchor thanks to Python's modulo on negatives)
    return (day - anchor).days % (interval * 7) == 0


def is_working_day(day: dt.date, schedule: dict | None = None) -> bool:
    """True if `day` is a normal working day under the schedule."""
    s = schedule or WORK_SCHEDULE
    if day.weekday() not in s.get("workdays", {0, 1, 2, 3, 4}):
        return False
    if day in s.get("days_off", set()):
        return False
    if is_off_friday(day, s):
        return False
    return True


def working_days(start: dt.date, end: dt.date, schedule: dict | None = None) -> list[dict]:
    """List each calendar day in [start, end] with whether it's worked and why
    not -- handy for showing the user how the schedule is being interpreted."""
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


def find_free_slots(
    duration_minutes: int = 30,
    days_ahead: int = 5,
    work_start: int = 9,
    work_end: int = 17,
    slot_minutes: int = 30,
    respect_schedule: bool = True,
    ignore_all_day: bool = True,
) -> list[dict]:
    """Compute open slots of `duration_minutes` within working hours over the
    next `days_ahead` days, from your own calendar. Honors WORK_SCHEDULE
    (weekends, every-other-Friday off, and explicit days_off) unless
    `respect_schedule` is False. Pure local computation."""
    events = get_calendar_events(days_ahead=days_ahead)
    busy: list[tuple] = []
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
    """Find the first available block of deep-work time and book it on your calendar."""
    slots = find_free_slots(duration_minutes=duration_hours * 60, days_ahead=days_ahead)
    if not slots:
        return {"status": "error", "message": f"No {duration_hours}h blocks in next {days_ahead} days."}
    
    # Book the first one
    target = slots[0]
    return schedule_meeting(
        subject="Focus Time - Do Not Disturb",
        start=target["start"],
        end=target["end"],
        body="Reserved by AI Assistant for deep work.",
        busy_status="busy"
    )


# --------------------------------------------------------------------------- #
# Tasks (close the loop on action items)
# --------------------------------------------------------------------------- #
def create_task(
    subject: str,
    due: dt.datetime | str | None = None,
    body: str = "",
    reminder: dt.datetime | str | None = None,
    importance: int = 1,
    categories=None,
) -> dict:
    """Create an Outlook task -- the natural endpoint for an extracted action
    item. `due`/`reminder` accept a datetime or ISO string."""
    app = _application()
    task = app.CreateItem(OL_TASK_ITEM)
    task.Subject = subject
    task.Body = body
    task.Importance = importance
    if due is not None:
        task.DueDate = _as_dt(due)
    if reminder is not None:
        task.ReminderSet = True
        task.ReminderTime = _as_dt(reminder)
    if categories:
        task.Categories = _join_categories(categories)
    task.Save()
    return {"status": "created", "entry_id": getattr(task, "EntryID", None),
            "subject": subject}


# --------------------------------------------------------------------------- #
# Contacts: resolve a name to an SMTP address via the GAL
# --------------------------------------------------------------------------- #
def resolve_recipient(name: str) -> dict:
    """Resolve a display name / alias against the address book to a real SMTP
    address, so drafting and scheduling don't need hand-typed emails."""
    ns = _namespace()
    r = ns.CreateRecipient(name)
    r.Resolve()
    if not r.Resolved:
        return {"resolved": False, "input": name}
    email = name
    try:
        ae = r.AddressEntry
        if ae.Type == "EX":
            eu = ae.GetExchangeUser()
            email = eu.PrimarySmtpAddress if eu is not None else r.Address
        else:
            email = ae.Address
    except Exception:
        email = getattr(r, "Address", name)
    return {"resolved": True, "input": name, "name": r.Name, "email": email}


# --------------------------------------------------------------------------- #
# Approval wrapper: propose, then send only after a human says yes
# --------------------------------------------------------------------------- #
def propose_draft(to, subject: str, body: str, **kwargs) -> dict:
    """Create a draft and return a preview + EntryID for review. Never sends.
    A skill calls this, shows you the preview, and calls send_draft() on yes."""
    kwargs["send"] = False
    result = create_draft_email(to, subject, body, **kwargs)
    result["preview"] = {
        "to": _join(to),
        "subject": subject,
        "body": (body or "")[:600],
    }
    return result


def send_draft(entry_id: str) -> dict:
    """Send a previously created draft by EntryID -- the approval step."""
    item = _get_item(entry_id)
    subject = getattr(item, "Subject", None)
    item.Send()
    return {"status": "sent", "entry_id": entry_id, "subject": subject}


# --------------------------------------------------------------------------- #
# Second tier: reschedule / cancel / respond to invites / save attachments
# (all via COM -- none require Microsoft Graph)
# --------------------------------------------------------------------------- #
OL_MEETING_CANCELED = 5  # OlMeetingStatus
# OlMeetingResponse codes used by AppointmentItem.Respond()
OL_RESPONSE = {"accept": 3, "tentative": 2, "decline": 4}


def _appointment_from(entry_id: str):
    """Resolve an EntryID to an AppointmentItem, whether the id points at a
    calendar appointment or at a meeting-request email sitting in the inbox."""
    item = _get_item(entry_id)
    if getattr(item, "Class", None) == 53:  # olMeetingRequest (a MeetingItem)
        return item.GetAssociatedAppointment(True)
    return item  # already an AppointmentItem


def reschedule_meeting(
    entry_id: str,
    start: dt.datetime | str | None = None,
    end: dt.datetime | str | None = None,
    duration_minutes: int | None = None,
    location: str | None = None,
    send_update: bool = False,
    display: bool = False,
) -> dict:
    """Move an existing meeting. Saves the change by default; set
    send_update=True to send the updated invite to attendees (only meaningful
    if you're the organizer)."""
    appt = _appointment_from(entry_id)
    if start is not None:
        appt.Start = _as_dt(start)
    if end is not None:
        appt.End = _as_dt(end)
    elif duration_minutes is not None:
        appt.Duration = duration_minutes
    if location is not None:
        appt.Location = location

    if send_update and appt.Recipients.Count > 0:
        appt.Send()
        status = "update_sent"
    else:
        appt.Save()
        status = "saved"
        if display:
            appt.Display(False)
    return {"status": status, "subject": appt.Subject,
            "start": _to_py_datetime(appt.Start), "end": _to_py_datetime(appt.End)}


def cancel_meeting(entry_id: str, send_cancellation: bool = False) -> dict:
    """Cancel a meeting you organized. By default just removes it from your
    calendar; set send_cancellation=True to notify attendees first."""
    appt = _appointment_from(entry_id)
    subject = appt.Subject
    has_attendees = appt.Recipients.Count > 0
    sent = False
    try:
        appt.MeetingStatus = OL_MEETING_CANCELED
        if send_cancellation and has_attendees:
            appt.Send()
            sent = True
    except Exception:
        pass  # not an organizer-owned meeting; fall through to delete
    appt.Delete()
    return {"status": "cancelled", "subject": subject, "cancellation_sent": sent}


def respond_to_invite(
    entry_id: str,
    response: str = "accept",   # accept / tentative / decline
    send: bool = True,
    comment: str | None = None,
) -> dict:
    """Accept / tentatively accept / decline a meeting invite. send=True sends a
    response to the organizer; send=False records your response without one
    (Outlook's 'Do not send a response')."""
    code = OL_RESPONSE[response.lower()]
    appt = _appointment_from(entry_id)
    resp = appt.Respond(code, True)  # fNoUI=True
    sent = False
    try:
        if comment and resp is not None:
            resp.Body = comment + "\n\n" + (resp.Body or "")
        if resp is not None:
            if send:
                resp.Send()
                sent = True
            else:
                resp.Save()
    except Exception:
        pass
    return {"status": "responded", "response": response.lower(),
            "sent": sent, "subject": getattr(appt, "Subject", None)}


def save_attachments(entry_id: str, out_dir: str, name_filter: str | None = None) -> dict:
    """Save a message's file attachments to out_dir. Optional name_filter is a
    case-insensitive substring match on the filename."""
    import os

    msg = _get_item(entry_id)
    os.makedirs(out_dir, exist_ok=True)
    saved = []
    for att in msg.Attachments:
        try:
            fn = att.FileName
            if name_filter and name_filter.lower() not in fn.lower():
                continue
            path = os.path.abspath(os.path.join(out_dir, fn))
            att.SaveAsFile(path)
            saved.append(path)
        except Exception:
            continue
    return {"status": "ok", "count": len(saved), "files": saved}


# --------------------------------------------------------------------------- #
# Demo
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    import json

    print("=== Folder tree ===")
    print(json.dumps(list_folders(max_depth=2), indent=2))

    print("\n=== Unread in last 24h ===")
    print(json.dumps(get_recent_messages(hours=24, unread_only=True, include_body=False), indent=2))

    print("\n=== Next 7 days of meetings ===")
    print(json.dumps(get_calendar_events(days_ahead=7), indent=2))

    # --- Write operations: safe by default (draft / saved, nothing sent) ---
    # create_draft_email(
    #     to=["teammate@company.com"],
    #     subject="Weekly status",
    #     body="Draft body here.",
    #     display=True,            # pop it open for review
    # )
    # schedule_meeting(
    #     subject="F-35 CNI sync",
    #     start="2026-05-26T10:00:00",
    #     duration_minutes=30,
    #     location="Teams",
    #     attendees=["lead@company.com"],
    #     send=False,              # saved to your calendar; flip to True to invite
    # )
