#!/usr/bin/env python3
"""
outlook_cli.py
--------------
A thin, JSON-emitting command line over outlook_helpers.py. This is the stable
contract a skill calls: the agent runs a subcommand, reads JSON from stdout, and
reasons over it. All deterministic I/O lives here; all judgment lives in the
skill/model. Nothing transmits except `send` (and `reply`/`meeting` with --send).

Examples:
    python outlook_cli.py recent --hours 168 --unread
    python outlook_cli.py events --days-ahead 7
    python outlook_cli.py free --duration 30 --days-ahead 5
    python outlook_cli.py draft --to a@x.com b@y.com --subject "Hi" --body-file -
    python outlook_cli.py send <entry_id>

Output: JSON on stdout. On error: {"error": "..."} on stdout and exit code 1.
"""

import argparse
import json
import os
import sys

# ---------------------------------------------------------------------------
# Mock injection: set OUTLOOK_MOCK=1 to use the in-memory stub instead of
# the real COM bridge.  Enables full test coverage without Outlook running.
# ---------------------------------------------------------------------------
if os.environ.get("OUTLOOK_MOCK"):
    import outlook_mock as ol  # type: ignore
else:
    import outlook_helpers as ol  # type: ignore

import outlook_contacts as oc  # pure Python — always available


def _emit(obj):
    json.dump(obj, sys.stdout, indent=2, default=str, ensure_ascii=False)
    sys.stdout.write("\n")


def _read_body(args):
    """Resolve message/appointment body from --body or --body-file ('-' = stdin)."""
    if getattr(args, "body_file", None):
        if args.body_file == "-":
            return sys.stdin.read()
        with open(args.body_file, "r", encoding="utf-8") as fh:
            return fh.read()
    return getattr(args, "body", "") or ""


def _expand_group(explicit: list | None, group_name: str | None) -> list:
    """
    Merge --to/--attendees list with --to-group/--attendees-group expansion.
    Returns a deduplicated combined list, raising ValueError if the group is
    specified but resolves to nothing (likely a typo).
    """
    result = list(explicit or [])
    if group_name:
        group_emails = oc.group_emails(group_name)
        if not group_emails:
            raise ValueError(
                f"Group {group_name!r} not found or empty in contacts.json. "
                f"Run: python outlook_cli.py group list"
            )
        for e in group_emails:
            if e not in result:
                result.append(e)
    return result


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="outlook_cli", description="Outlook over JSON.")
    sub = p.add_subparsers(dest="cmd", required=True)

    # ---- Read: mail -------------------------------------------------------- #
    s = sub.add_parser("recent", help="Recent Inbox messages")
    s.add_argument("--hours", type=int, default=24)
    s.add_argument("--unread", action="store_true")
    s.add_argument("--max", type=int, default=50)
    s.add_argument("--no-body", action="store_true")

    s = sub.add_parser("folder", help="Messages from a named folder")
    s.add_argument("path")
    s.add_argument("--hours", type=int, default=None)
    s.add_argument("--unread", action="store_true")
    s.add_argument("--max", type=int, default=50)
    s.add_argument("--base", default="inbox", choices=["inbox", "root"])

    s = sub.add_parser("folders", help="List the folder tree")
    s.add_argument("--max-depth", type=int, default=3)

    s = sub.add_parser("search", help="Search Inbox subjects")
    s.add_argument("query")
    s.add_argument("--max", type=int, default=25)

    sub.add_parser("radar", help="List messages flagged for follow-up")

    s = sub.add_parser("message", help="One message in full")
    s.add_argument("entry_id")

    s = sub.add_parser("thread", help="Full conversation thread")
    s.add_argument("entry_id")
    s.add_argument("--max", type=int, default=50)

    s = sub.add_parser("summarize-thread", help="Flattened thread data for LLM summarization")
    s.add_argument("entry_id")
    s.add_argument("--max", type=int, default=25)

    # ---- Read: calendar ---------------------------------------------------- #
    s = sub.add_parser("events", help="Calendar events in a window")
    s.add_argument("--days-ahead", type=int, default=7)
    s.add_argument("--days-back", type=int, default=0)

    s = sub.add_parser("pulse", help="Check upcoming meetings for issues (e.g. no agenda)")
    s.add_argument("--days-ahead", type=int, default=7,
                   help="How far ahead to scan (default 7 — full week)")

    s = sub.add_parser("analytics", help="Time spent in meetings by category")
    s.add_argument("--days-back", type=int, default=7)

    s = sub.add_parser("free", help="Find open slots")
    s.add_argument("--duration", type=int, default=30)
    s.add_argument("--days-ahead", type=int, default=5)
    s.add_argument("--work-start", type=int, default=9)
    s.add_argument("--work-end", type=int, default=17)
    s.add_argument("--slot", type=int, default=30)
    s.add_argument("--all-days", action="store_true",
                   help="ignore the work schedule (include weekends/off-Fridays)")

    s = sub.add_parser("workdays", help="Show which upcoming days are worked vs off")
    s.add_argument("--days-ahead", type=int, default=14)

    s = sub.add_parser("focus", help="Auto-book a block of focus time")
    s.add_argument("--hours", type=int, default=2)
    s.add_argument("--days-ahead", type=int, default=3)

    # ---- Write: drafts / send (the approval loop) -------------------------- #
    s = sub.add_parser("draft", help="Create a draft (never sends), return preview + entry_id")
    s.add_argument("--to", nargs="*", default=None)
    s.add_argument("--to-group", default=None,
                   help="Send to all members of a contact group (can combine with --to)")
    s.add_argument("--subject", required=True)
    s.add_argument("--body", default="")
    s.add_argument("--body-file", default=None, help="path, or '-' for stdin")
    s.add_argument("--cc", nargs="*", default=None)
    s.add_argument("--html", action="store_true")
    s.add_argument("--display", action="store_true")

    s = sub.add_parser("send", help="Send a previously created draft by entry_id (the 'yes')")
    s.add_argument("entry_id")

    s = sub.add_parser("reply", help="Draft a thread-aware reply")
    s.add_argument("entry_id")
    s.add_argument("--body", default="")
    s.add_argument("--body-file", default=None)
    s.add_argument("--all", action="store_true", help="reply-all")
    s.add_argument("--send", action="store_true")
    s.add_argument("--display", action="store_true")

    s = sub.add_parser("forward", help="Forward a message to new recipients (draft by default)")
    s.add_argument("entry_id")
    s.add_argument("--to", nargs="*", default=None)
    s.add_argument("--to-group", default=None,
                   help="Forward to all members of a contact group")
    s.add_argument("--body", default="")
    s.add_argument("--body-file", default=None)
    s.add_argument("--html", action="store_true")
    s.add_argument("--send", action="store_true")
    s.add_argument("--display", action="store_true")

    # ---- Write: calendar / tasks ------------------------------------------ #
    s = sub.add_parser("meeting", help="Create appointment / meeting request")
    s.add_argument("--subject", required=True)
    s.add_argument("--start", required=True, help="ISO datetime")
    s.add_argument("--end", default=None, help="ISO datetime")
    s.add_argument("--duration", type=int, default=30)
    s.add_argument("--location", default="")
    s.add_argument("--body", default="")
    s.add_argument("--attendees", nargs="*", default=None)
    s.add_argument("--attendees-group", default=None,
                   help="Add all members of a contact group as required attendees")
    s.add_argument("--optional", nargs="*", default=None)
    s.add_argument("--send", action="store_true", help="actually send invites")

    s = sub.add_parser("task", help="Create an Outlook task from an action item")
    s.add_argument("--subject", required=True)
    s.add_argument("--due", default=None, help="ISO datetime")
    s.add_argument("--body", default="")
    s.add_argument("--reminder", default=None, help="ISO datetime")
    s.add_argument("--categories", nargs="*", default=None)

    # ---- Triage / organize ------------------------------------------------- #
    s = sub.add_parser("mark", help="Mark read/unread, flag, and/or move a message")
    s.add_argument("entry_id")
    s.add_argument("--read", dest="read", action="store_true")
    s.add_argument("--unread", dest="unread", action="store_true")
    s.add_argument("--flag", action="store_true")
    s.add_argument("--unflag", action="store_true")
    s.add_argument("--move", default=None, help="destination folder path")

    s = sub.add_parser("categorize", help="Set/add/remove categories on mail or a meeting")
    s.add_argument("entry_id")
    s.add_argument("--set", nargs="*", default=None)
    s.add_argument("--add", nargs="*", default=None)
    s.add_argument("--remove", nargs="*", default=None)

    sub.add_parser("categories", help="List master category list (with color names)")

    s = sub.add_parser("init-categories", help="Sync CATEGORY_SCHEME into Outlook (idempotent)")
    s.add_argument("--remove-unlisted", action="store_true")

    s = sub.add_parser("resolve", help="Resolve a name to an SMTP address via the GAL")
    s.add_argument("name")

    # ---- Second tier: reschedule / cancel / respond / attachments ---------- #
    s = sub.add_parser("reschedule", help="Move an existing meeting")
    s.add_argument("entry_id")
    s.add_argument("--start", default=None, help="ISO datetime")
    s.add_argument("--end", default=None, help="ISO datetime")
    s.add_argument("--duration", type=int, default=None)
    s.add_argument("--location", default=None)
    s.add_argument("--send-update", action="store_true", help="notify attendees")

    s = sub.add_parser("cancel", help="Cancel a meeting you organized")
    s.add_argument("entry_id")
    s.add_argument("--send-cancellation", action="store_true", help="notify attendees")

    s = sub.add_parser("respond", help="Respond to a meeting invite")
    s.add_argument("entry_id")
    s.add_argument("--response", choices=["accept", "tentative", "decline"], default="accept")
    s.add_argument("--no-send", action="store_true", help="record without sending a response")
    s.add_argument("--comment", default=None)

    s = sub.add_parser("save-attachments", help="Save a message's attachments to a folder")
    s.add_argument("entry_id")
    s.add_argument("out_dir")
    s.add_argument("--filter", default=None, help="filename substring filter")

    # ---- Contact groups ---------------------------------------------------- #
    s = sub.add_parser("group", help="Contact group management (list/show/emails/add/remove/who/tier)")
    s.add_argument("action",
                   choices=["list", "show", "emails", "add", "remove", "who", "tier"],
                   help="Action: list | show <group> | emails <group> | "
                        "add <group> <email> | remove <group> <email> | who <email> | tier <email>")
    s.add_argument("group_or_email", nargs="?", default=None,
                   help="Group name (show/emails/add/remove) or email (who/tier)")
    s.add_argument("email", nargs="?", default=None,
                   help="Email address (for add / remove)")
    s.add_argument("--name", default=None, help="Display name when adding a contact")

    s = sub.add_parser("enrich", help="Add relationship tier + priority signals to recent mail")
    s.add_argument("--hours", type=int, default=168)
    s.add_argument("--unread", action="store_true")

    return p


def dispatch(args) -> object:
    c = args.cmd
    if c == "recent":
        return ol.get_recent_messages(hours=args.hours, unread_only=args.unread,
                                      max_items=args.max, include_body=not args.no_body)
    if c == "folder":
        return ol.get_messages_from_folder(args.path, hours=args.hours,
                                           unread_only=args.unread, max_items=args.max,
                                           base=args.base)
    if c == "folders":
        return ol.list_folders(max_depth=args.max_depth)
    if c == "search":
        return ol.search_messages(args.query, max_items=args.max)
    if c == "radar":
        return ol.get_flagged_items()
    if c == "message":
        return ol.get_message(args.entry_id)
    if c == "thread":
        return ol.get_conversation(args.entry_id, max_items=args.max)
    if c == "summarize-thread":
        return ol.get_conversation_summary_data(args.entry_id, max_items=args.max)
    if c == "events":
        return ol.get_calendar_events(days_ahead=args.days_ahead, days_back=args.days_back)
    if c == "pulse":
        return ol.get_meeting_health(days_ahead=args.days_ahead)
    if c == "analytics":
        return ol.get_time_spent_analytics(days_back=args.days_back)
    if c == "free":
        return ol.find_free_slots(duration_minutes=args.duration, days_ahead=args.days_ahead,
                                  work_start=args.work_start, work_end=args.work_end,
                                  slot_minutes=args.slot, respect_schedule=not args.all_days)
    if c == "workdays":
        import datetime as _dt
        today = _dt.date.today()
        return ol.working_days(today, today + _dt.timedelta(days=args.days_ahead))
    if c == "focus":
        return ol.auto_block_focus_time(duration_hours=args.hours, days_ahead=args.days_ahead)
    if c == "draft":
        to = _expand_group(args.to, getattr(args, "to_group", None))
        if not to:
            return {"error": "No recipients: provide --to and/or --to-group"}
        return ol.propose_draft(to=to, subject=args.subject, body=_read_body(args),
                                cc=args.cc, html=args.html, display=args.display)
    if c == "send":
        return ol.send_draft(args.entry_id)
    if c == "reply":
        return ol.reply_to_message(args.entry_id, body=_read_body(args), reply_all=args.all,
                                   send=args.send, display=args.display)
    if c == "forward":
        to = _expand_group(args.to, getattr(args, "to_group", None))
        if not to:
            return {"error": "No recipients: provide --to and/or --to-group"}
        return ol.forward_message(args.entry_id, to=to, body=_read_body(args),
                                  html=args.html, send=args.send, display=args.display)
    if c == "meeting":
        attendees = _expand_group(args.attendees, getattr(args, "attendees_group", None))
        return ol.schedule_meeting(subject=args.subject, start=args.start, end=args.end,
                                   duration_minutes=args.duration, location=args.location,
                                   body=args.body, attendees=attendees or None,
                                   optional_attendees=args.optional, send=args.send)
    if c == "task":
        return ol.create_task(subject=args.subject, due=args.due, body=args.body,
                              reminder=args.reminder, categories=args.categories)
    if c == "mark":
        results = {}
        if args.read:
            results["read"] = ol.mark_read(args.entry_id, read=True)
        if args.unread:
            results["read"] = ol.mark_read(args.entry_id, read=False)
        if args.flag:
            results["flag"] = ol.flag_message(args.entry_id, flag=True)
        if args.unflag:
            results["flag"] = ol.flag_message(args.entry_id, flag=False)
        if args.move:
            results["move"] = ol.move_message(args.entry_id, args.move)
        return results or {"status": "no-op", "hint": "pass --read/--unread/--flag/--unflag/--move"}
    if c == "categorize":
        if args.set is not None:
            return ol.set_categories(args.entry_id, args.set)
        if args.add is not None:
            return ol.add_categories(args.entry_id, args.add)
        if args.remove is not None:
            return ol.remove_categories(args.entry_id, args.remove)
        return ol.get_categories(args.entry_id)
    if c == "categories":
        return ol.list_master_categories()
    if c == "init-categories":
        return ol.init_categories(remove_unlisted=args.remove_unlisted)
    if c == "resolve":
        # Check contacts.json first (groups + aliases), fall back to GAL
        contacts_result = oc.resolve(args.name)
        if contacts_result["type"] != "unknown":
            return contacts_result
        return ol.resolve_recipient(args.name)
    if c == "reschedule":
        return ol.reschedule_meeting(args.entry_id, start=args.start, end=args.end,
                                     duration_minutes=args.duration, location=args.location,
                                     send_update=args.send_update)
    if c == "cancel":
        return ol.cancel_meeting(args.entry_id, send_cancellation=args.send_cancellation)
    if c == "respond":
        return ol.respond_to_invite(args.entry_id, response=args.response,
                                    send=not args.no_send, comment=args.comment)
    if c == "save-attachments":
        return ol.save_attachments(args.entry_id, args.out_dir, name_filter=args.filter)
    if c == "group":
        action = args.action
        g = args.group_or_email
        e = args.email
        if action == "list":
            return oc.group_list()
        if action == "show":
            return oc.group_show(g)
        if action == "emails":
            return oc.group_emails(g)
        if action == "add":
            return oc.group_add(g, e, name=args.name)
        if action == "remove":
            return oc.group_remove(g, e)
        if action in ("who", "tier"):
            result = oc.who_is(g)   # g holds the email for who/tier
            if action == "tier":
                result["tier"] = oc.priority_tier(g)
            return result
    if c == "enrich":
        msgs = ol.get_recent_messages(hours=args.hours, unread_only=args.unread)
        return oc.enrich_messages(msgs)
    raise ValueError(f"unknown command: {c}")


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    try:
        _emit(dispatch(args))
        return 0
    except Exception as exc:  # surface a clean JSON error for the agent
        _emit({"error": f"{type(exc).__name__}: {exc}"})
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
