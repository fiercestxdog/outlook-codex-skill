---
name: weekly-brief
description: Generate a Monday weekly brief from Outlook by pulling the week's calendar, unread mail, open follow-ups, and free slots, then synthesizing them into priorities, decisions owed, meeting prep, and waiting-on items — and optionally drafting a status update held for approval. Use this whenever the user asks for a weekly brief, Monday brief, weekly status, "what's on my plate this week," week-ahead meeting prep, or wants their Outlook inbox and calendar summarized into priorities and action items, even if they don't say the word "brief."
---

# Weekly Brief

Turn the user's Outlook mail and calendar into a concise, decision-oriented brief.
All Outlook access goes through the CLI in `scripts/`; this file is the judgment
layer that decides what matters and how to present it.

## Prerequisites (read once)

- Runs on the user's **Windows machine with Outlook installed and running** —
  COM talks to the local, already-authenticated client. It will NOT work in a
  cloud sandbox. If `python scripts/outlook_cli.py recent` errors with a COM
  failure, say so plainly rather than inventing data.
- One-time setup: `pip install pywin32`, then `python scripts/outlook_cli.py init-categories`.
- **Work schedule is encoded** in `WORK_SCHEDULE` (in `outlook_helpers.py`): Mon–Fri,
  every-other-Friday off (9/80 RDO, anchored 2026-05-29), and US federal holidays
  for 2026–2027. Slot-finding already respects all of this; run
  `python scripts/outlook_cli.py workdays --days-ahead 14` to see worked vs off days.
- **Data handling:** message bodies stay local unless you transmit them. Do not
  paste full bodies into any external system. Summarize; cite by subject/sender.
- **Personal Style:** Always consult `STYLE_GUIDE.md` before drafting replies
  or new emails. Adhere strictly to the "Writing DNA" and "Banned Words" rules.

## How to call the data layer

Every command prints JSON to stdout. Run from the `scripts/` directory:

```
python outlook_cli.py recent --hours 168 --unread        # unread, last 7 days
python outlook_cli.py events --days-ahead 7              # week-ahead meetings
python outlook_cli.py free --duration 30 --days-ahead 5  # open 30-min slots
python outlook_cli.py thread <entry_id>                  # full thread for context
python outlook_cli.py categories                         # the label taxonomy
python outlook_cli.py radar                              # flagged follow-ups
python outlook_cli.py analytics --days-back 7            # time spent in meetings
python outlook_cli.py pulse                              # meeting health check
python outlook_cli.py focus --hours 2                    # auto-book focus time
python outlook_cli.py summarize-thread <entry_id>        # flat thread data for LLM
```

On any error the JSON is `{"error": "..."}` — surface it; do not fabricate.

## 🛠️ CLI Reference & Examples

### Data Gathering (Read)
- **Weekly Context**: `python outlook_cli.py recent --hours 168` (Last 7 days of inbox).
- **Unread Only**: `python outlook_cli.py recent --hours 168 --unread` (Unread only — faster for triage).
- **Upcoming Schedule**: `python outlook_cli.py events --days-ahead 7` (Next 7 days of meetings). ⚠️ Events include `entry_id` — save it for reschedule/cancel/respond.
- **Availability**: `python outlook_cli.py free --duration 45 --days-ahead 3` (Find 45-min slots in next 3 days).
- **Thread for LLM Reasoning**: `python outlook_cli.py summarize-thread <entry_id>` (Flat text dump for summarization — prefer over `thread` when you'll reason over the content).
- **Thread with Full Structure**: `python outlook_cli.py thread <entry_id>` (Structured JSON, sorted oldest-first — use when you need metadata like sender/date per message).
- **Search**: `python outlook_cli.py search "Architecture Review"` (Case-insensitive subject search).
- **Flagged Follow-ups**: `python outlook_cli.py radar` (Messages flagged for follow-up).
- **Schedule Check**: `python outlook_cli.py workdays --days-ahead 14` (Working vs off days — RDOs, holidays). ⚠️ Returns N+1 entries (today through today+N inclusive).
- **Meeting Health**: `python outlook_cli.py pulse` (Meetings with thin/missing agendas in the next 7 days).
- **Time Audit**: `python outlook_cli.py analytics --days-back 7` (Hours in meetings by category).
- **Subfolder Messages**: `python outlook_cli.py folder "Projects/CNI"` (Messages from a named subfolder).
- **Folder Tree**: `python outlook_cli.py folders` (All mail folders with unread counts).

### Triage & Coordination (Write)
- **Drafting**: `python outlook_cli.py draft --to "john.doe@example.com" --subject "Weekly Sync" --body-file -` (Pipe a body into a new draft).
- **Replying**: `python outlook_cli.py reply <entry_id> --body "I will have this to you by EOD."` (Quick thread-aware reply).
- **Forwarding**: `python outlook_cli.py forward <entry_id> --to "colleague@co.com" --body "FYI —"` (Forward as draft; add `--send` only after user approval).
- **Tasks**: `python outlook_cli.py task --subject "Submit Expense Report" --due "2026-05-30" --categories "Action item"` (Convert a mail to a task).
- **Organization**: `python outlook_cli.py categorize <entry_id> --add "Decision needed"` (Apply a triage label).
- **Rescheduling**: `python outlook_cli.py reschedule <entry_id> --start "2026-05-26T14:00:00"` (Move a meeting; add `--send-update` to notify attendees).
- **Cancel**: `python outlook_cli.py cancel <entry_id>` (Remove a meeting; add `--send-cancellation` to notify attendees).
- **Respond to Invite**: `python outlook_cli.py respond <entry_id> --response accept` (Accept / tentative / decline; add `--no-send` to record silently).
- **Save Attachments**: `python outlook_cli.py save-attachments <entry_id> ./downloads` (Save message attachments locally).
- **Schedule Meeting**: `python outlook_cli.py meeting --subject "..." --start "2026-06-01T10:00:00" --duration 30 --attendees a@co.com b@co.com` (Creates on your calendar; add `--send` to send invites).

---

## 💡 Use Cases (Prompt Shots)

### 1. The "Monday Morning Brief"
**User:** "What's my week look like? Give me a brief."
**Agent Strategy:**
1. Run `workdays --days-ahead 7` to check for RDOs/Holidays.
2. Run `events --days-ahead 7` to see meeting load.
3. Run `recent --hours 72 --unread` to catch up on the weekend's traffic.
4. Synthesize into "Top Priorities", "Decisions Owed", and "Meeting Prep".

### 2. The "Meeting Context Resolver"
**User:** "I have a meeting with Sarah about 'Project Phoenix' in 20 minutes. What's the latest status?"
**Agent Strategy:**
1. Run `search "Project Phoenix"` to find related threads.
2. Run `summarize-thread <entry_id>` on the most recent 2-3 results (use `summarize-thread`, not `thread` — it returns a single flat text string optimized for reasoning rather than structured JSON).
3. Summarize: what was the last decision? What was Sarah's last request? What's unresolved?
4. Check `events --days-ahead 1` for any prep materials attached to the invite.

### 3. The "Conflict Negotiator"
**User:** "I can't make the 2pm sync today. Find another time with Bob tomorrow and draft a note."
**Agent Strategy:**
1. Run `events` for tomorrow to see user availability.
2. Run `free --duration 30 --days-ahead 2` to identify gaps.
3. Run `resolve "Bob"` to get his email address.
4. Run `draft --to <bob_email> --subject "Rescheduling 2pm Sync" --body "Hi Bob, I can't make 2pm today. Would 10am or 3pm tomorrow work for you?"`
5. Show the user the draft and the proposed times for approval.

### 4. The "Inbox Triage"
**User:** "Triage my unread emails from the last 24 hours."
**Agent Strategy:**
1. Run `recent --hours 24 --unread`.
2. For each message:
   - If it's a request for a decision -> `categorize --add "Decision needed"`.
   - If it's a task for the user -> `task --subject "..."` + `categorize --add "Action item"`.
   - If it's just info -> `categorize --add "FYI / read later"`.
3. Report a summary of labels applied and any high-priority items found.

### 5. The "Digital Twin" Ghostwriter
**User:** "Draft a reply to Sarah's last email about the budget."
**Agent Strategy:**
1. Run `search "budget"` to find Sarah's email.
2. Run `summarize-thread <entry_id>` to get full context.
3. Read `STYLE_GUIDE.md` to understand the user's voice.
4. Compose a draft using `reply <entry_id> --body "..."` following the style rules (concise, no AI-isms).
5. Show the user the preview and ask for explicit approval.
6. Only on yes: `send <entry_id>`.

### 6. The "Meeting Scheduler"
**User:** "Set up a 45-minute review with Tom and Priya for Tuesday at 2pm."
**Agent Strategy:**
1. Run `workdays --days-ahead 7` to confirm Tuesday is a working day (not RDO/holiday).
2. Run `events --days-ahead 7` to check for conflicts at 2pm Tuesday; if blocked, note it.
3. Run `resolve "Tom"` and `resolve "Priya"` — contacts.json aliases + GAL are both checked.
4. Compose: `meeting --subject "Q2 Review" --start "2026-05-26T14:00:00" --duration 45 --attendees <tom_email> <priya_email> --body "<agenda>"`.
5. Show the proposed invite details and ask: "Want me to send the invites?"
6. Only on yes: re-run with `--send`.

**Group shortcut variant:**
- "Team sync" → `meeting --subject "Team Sync" --start "..." --attendees-group direct_reports --send`
- "Status update to boss" → `draft --subject "Weekly Status" --to-group supervisors --body-file -`

> **Note:** `meeting` without `--send` saves to your calendar only. Invites are NOT sent until `--send` is passed and the user has approved.

### 7. The "Time Audit"
**User:** "How much time am I spending in meetings this week? I feel like I have no focus time."
**Agent Strategy:**
1. Run `analytics --days-back 7` for meeting load over the last week.
2. Run `events --days-ahead 7` to see what's coming.
3. Run `pulse --days-ahead 7` to flag meetings with thin agendas (candidates for cancellation).
4. Summarize: total meeting hours, meetings-per-day average, lightest day, heaviest day.
5. Offer to run `focus --hours 2 --days-ahead 5` to auto-book a focus block if the user wants.

## Workflow

### 1. Gather (run these, then reason over the JSON)
- `enrich --hours 168` instead of bare `recent` — this adds `tier` + `priority_signals` to every message so synthesis is relationship-aware. Falls back gracefully if `contacts.json` is empty.
- `events --days-ahead 7` for the week's meetings.
- `free --duration 30 --days-ahead 5` for schedulable openings.
- `workdays --days-ahead 14` so the brief can call out a non-working day in the window (an RDO/off-Friday or a holiday) that affects deadlines or availability.
- For any thread that needs a decision or reply, pull `summarize-thread <entry_id>` before reasoning — never summarize a reply chain from the latest message alone.

### 2. Synthesize
Cluster items into the sections below. Be selective: a brief is triage, not a dump.
Tag each item's confidence when sources disagree — VERIFIED (seen in 2+ places),
SINGLE-SOURCE, or INFERRED — and keep INFERRED out of the "decisions" section.

**Relationship-aware priority rules** (applied after `enrich`):

| Signal in `priority_signals` | How to handle |
|---|---|
| `always_surface` (supervisor) | Appears in brief regardless of Outlook importance; unread → "Decisions you owe" |
| `review_flag` (supervisor / program_lead requesting sign-off) | Always goes to "Decisions you owe" with deadline if stated |
| `blocker_flag` (direct report says they're stuck) | Goes to "Waiting on others" with flag — your team is blocked on YOU |
| `stale_Nh` (noteworthy sender, unread > 24 h) | Flag with age in "Waiting on others" or "FYIs" |
| `tier = unknown`, no signals | Standard Outlook importance + recency rules apply |

**Group dispatch shortcuts** — use when the user says:
- "Send to my team" → `draft --to-group direct_reports`
- "Loop in my boss" → `draft --to-group supervisors` or `forward <id> --to-group supervisors`
- "Schedule a team sync" → `meeting --attendees-group direct_reports`
- "Who is blocking me?" → look for `blocker_flag` in `enrich` output

### 3. Output format
Produce the brief in this order. Lead with what needs the user's judgment.

> **Weekly Brief — <week of DATE>**
>
> **Top priorities (3–5).** The few things that actually move this week.
> **Decisions you owe.** Each: what, to whom, by when, and the source (subject/sender).
> **Meetings & prep.** Per meeting: what it's for, what to prepare, open free slots if prep time is needed.
> **Waiting on others.** Open loops where someone owes the user — flag the stale ones.
> **FYIs.** Low-effort awareness items, one line each.

If a non-working day (RDO/off-Friday or holiday) falls in the week or just after,
note it where it matters — e.g. "Heads up: Fri is your RDO, so the review compresses
into Thu." Don't manufacture urgency; only flag it when it actually affects timing.

Keep it scannable. No raw EntryIDs in the prose (keep them aside for actions).

### 4. Optional: draft a status update or replies (approval-gated)
If the user wants a status update or a reply sent:
- Compose the text, then create it as a **draft only**:
  `python outlook_cli.py draft --to <addr> --subject "<s>" --body-file -`
  (pipe the body via stdin), or `reply <entry_id> --body-file -`.
- Show the user the returned `preview` and the `entry_id`. Ask for explicit approval.
- Only after the user says yes: `python outlook_cli.py send <entry_id>`.
- **Never** pass `--send` on `draft`/`reply`/`meeting` on the user's behalf. Sending
  is a separate, explicit human-approved step.

### 5. Optional: organize while you triage
When the user asks, apply the taxonomy (see `categories`):
- `categorize <entry_id> --add "Decision needed"` / `"Waiting on reply"` / `"Action item"`.
- Turn a committed action item into a tracked task: `task --subject "<x>" --due <ISO>`.
- Move or flag: `mark <entry_id> --flag` or `--move "Projects/<name>"`.

## Guardrails
- Default to read-only. Writing a draft is fine; sending, moving, and deleting are not
  done without the user asking.
- If a command returns `{"error": ...}` or empty results, report that honestly and
  continue with what you do have — a partial brief beats a confident fabrication.
- Don't invent attendees, dates, or commitments not present in the JSON.
