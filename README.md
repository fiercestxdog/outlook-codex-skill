# Outlook Codex Skill

A local-first, JSON-emitting CLI that lets a Claude Code skill (or any LLM) drive
**Classic Outlook** via COM automation — no Azure registration, no OAuth, no IT ticket.
Powers weekly briefs, inbox triage, meeting scheduling, and contact-aware priority scoring.

> **Current version:** v2.0.0 — see [CHANGELOG.md](CHANGELOG.md)

---

## ⚠️ Requirements

| Requirement | Detail |
|---|---|
| **OS** | Windows only (COM automation is Windows-specific) |
| **Outlook** | **Classic Outlook only** — New Outlook (Win11 default) does NOT support COM. Toggle: Outlook → Settings → turn off *"Try the new Outlook"* |
| **Python** | 3.8+ |
| **pywin32** | `pip install pywin32>=306` |

> **M365 Security Popup (post-Oct 2024):** Outlook may prompt *"A program is trying to
> send an email on your behalf"* for draft/send operations. Workarounds:
> 1. Add an antivirus exception for `python.exe` (simplest)
> 2. Domain GPO: `HKLM\...\Office\16.0\Outlook\Security\ObjectModelGuard=0`
> 3. Move send/draft to Microsoft Graph for fully unattended use

---

## 🚀 Setup

### 1. Download

```powershell
git clone https://github.com/fiercestxdog/outlook-codex-skill.git
cd outlook-codex-skill
```

Or download the zip from the [v2.0.0 release](https://github.com/fiercestxdog/outlook-codex-skill/releases/tag/v2.0.0) and extract it.

### 2. Install dependency

```powershell
pip install pywin32>=306
```

### 3. Verify Classic Outlook is running

```powershell
python -c "import outlook_helpers; print(outlook_helpers.__version__)"
```

- Prints `2.0.0` → Classic Outlook detected, COM bridge is working ✅
- Raises `RuntimeError` → New Outlook or Outlook not running — follow the message instructions

### 4. Initialize categories

Run once to create the triage label set in Outlook's master category list:

```powershell
python outlook_cli.py init-categories
```

### 5. Verify the connection

```powershell
python outlook_cli.py workdays --days-ahead 7    # shows worked vs RDO/holiday days
python outlook_cli.py recent --hours 1           # last hour of inbox
python outlook_cli.py events --days-ahead 3      # next 3 days of calendar
```

All commands emit JSON to stdout. On error: `{"error": "..."}`.

### 6. Configure contact groups (optional but recommended)

Copy the template and fill in real addresses:

```powershell
copy contacts.json contacts.json.bak   # keep the template
```

Edit `contacts.json` — add your direct reports, supervisor, peers, and program leads.
The skill uses this for priority scoring (supervisor emails always surface first) and
group dispatch (`--to-group direct_reports`, `--attendees-group supervisors`).

> **⚠️ Privacy:** `contacts.json` with real employee data is a secrets file.
> Do not commit or sync it to any public or shared repository.

---

## 🛠️ CLI Reference

All commands: `python outlook_cli.py <command> [options]`  
Output: JSON on stdout. Pipe to `jq` or let the skill parse it.

### Read

| Command | What it returns | Key options |
|---|---|---|
| `recent` | Inbox messages | `--hours 168` `--unread` `--max 50` |
| `enrich` | Inbox + contact tier + priority signals | `--hours 168` (preferred over `recent` for briefs) |
| `events` | Calendar appointments | `--days-ahead 7` |
| `free` | Open time slots | `--duration 30` `--days-ahead 5` |
| `thread` | Full conversation (structured JSON) | `<entry_id>` |
| `summarize-thread` | Flat text for LLM reasoning | `<entry_id>` |
| `search` | Subject search | `"query string"` |
| `radar` | Flagged follow-ups | — |
| `analytics` | Meeting hours by category | `--days-back 7` |
| `pulse` | Meetings with thin/missing agendas | `--days-ahead 7` |
| `workdays` | Worked vs RDO/holiday calendar | `--days-ahead 14` |
| `folder` | Messages from a subfolder | `"Projects/CNI"` |
| `folders` | Full folder tree with unread counts | — |
| `categories` | Master category list | — |

### Write (all draft by default — nothing sends without `--send` + user approval)

| Command | What it does | Key options |
|---|---|---|
| `draft` | Create a draft email | `--to` `--to-group` `--subject` `--body` `--body-file -` |
| `reply` | Reply to a thread | `<entry_id>` `--body` `--reply-all` |
| `forward` | Forward a message | `<entry_id>` `--to` `--to-group` `--body` |
| `send` | Send a saved draft | `<entry_id>` |
| `meeting` | Create a calendar event / send invites | `--subject` `--start` `--duration` `--attendees` `--attendees-group` `--send` |
| `reschedule` | Move an existing meeting | `<entry_id>` `--start` `--duration` `--send-update` |
| `cancel` | Cancel a meeting | `<entry_id>` `--send-cancellation` |
| `respond` | Accept/tentative/decline an invite | `<entry_id>` `--response accept\|tentative\|decline` |
| `task` | Create an Outlook task | `--subject` `--due` `--categories` |
| `categorize` | Add/remove categories on any item | `<entry_id>` `--add` `--remove` |
| `mark` | Flag, move, or mark read | `<entry_id>` `--flag` `--read` `--move "folder"` |
| `focus` | Auto-book a focus block | `--hours 2` `--days-ahead 3` |
| `save-attachments` | Save file attachments locally | `<entry_id>` `./downloads` |
| `init-categories` | Sync category taxonomy to Outlook | — |

### Contacts & Groups

| Command | What it does | Example |
|---|---|---|
| `resolve` | Name → SMTP (contacts.json first, then GAL) | `resolve "Alice"` |
| `group list` | Show all groups | `group list` |
| `group show` | Members of one group | `group show direct_reports` |
| `group emails` | Email addresses for a group | `group emails supervisors` |
| `group add` | Add someone to a group | `group add direct_reports alice@co.com` |
| `group remove` | Remove from a group | `group remove peers dave@co.com` |
| `group who` | Which groups an email belongs to | `group who alice@co.com` |
| `group tier` | Priority tier for an email | `group tier manager@co.com` |

### Group dispatch shortcuts

```powershell
python outlook_cli.py draft --to-group direct_reports --subject "Team update" --body "..."
python outlook_cli.py meeting --subject "Team sync" --start "2026-06-01T10:00" --attendees-group direct_reports --send
python outlook_cli.py forward <entry_id> --to-group supervisors --body "FYI —"
```

Natural language in `contacts.json` maps "my team" → direct_reports, "boss" → supervisors, etc.

---

## 🧬 Architecture

```
SKILL.md                ← judgment layer (what the LLM does with the data)
    └── outlook_cli.py  ← stable JSON CLI contract (subprocess calls)
            ├── outlook_helpers.py   ← COM automation (mail, calendar, tasks)
            ├── outlook_mock.py      ← in-memory stub (OUTLOOK_MOCK=1)
            └── outlook_contacts.py ← contact groups + priority scoring (pure Python)
                    └── contacts.json
```

**`OUTLOOK_MOCK=1`** — set this env var to use the in-memory stub instead of live Outlook.
Used by the test suite; also useful for local development on a machine without Outlook.

**Priority tiers** (applied by `enrich`): supervisor > program_lead > direct_report > peer > unknown  
**Priority signals**: `always_surface`, `blocker_flag`, `review_flag`, `stale_Nh`

**Work schedule** — encoded in `WORK_SCHEDULE` in `outlook_helpers.py`:
- Mon–Fri, 9/80 (every-other-Friday off, anchor `2026-05-29`)
- US federal holidays 2026–2027 preloaded
- Edit `WORK_SCHEDULE["days_off"]` to add PTO

**IncludeRecurrences ordering** — `get_calendar_events()` enforces the mandatory
Sort → IncludeRecurrences → Restrict sequence required by Outlook COM. Do not reorder.

---

## 🧪 Testing

Run the full test suite without Outlook:

```powershell
pip install pytest
$env:OUTLOOK_MOCK = "1"
pytest tests/ -v
```

- `tests/test_cli.py` — 105 tests covering all 30+ subcommands (subprocess-based)
- `tests/test_contacts.py` — 54 tests for contact groups, priority scoring, group dispatch
- `tests/fixtures/contacts.json` — isolated fixture with 5 contacts, 4 groups

Stateful round-trips (draft → send) use in-process module reload for state isolation.

---

## 🔒 Security & Privacy

- **Local only** — no cloud APIs, no OAuth, no data leaves your machine unless you call `send`
- **Approval-gated writes** — `draft`, `reply`, `meeting`, `forward` all save as drafts; `send` is a separate explicit step
- **`contacts.json`** — treat as a secrets file once populated with real employee data; never push to a shared or public repo
- **Locale** — `_DATE_FMT` assumes US Windows locale. Run `python -c "import outlook_helpers; print(outlook_helpers._detect_date_fmt())"` to verify on non-US machines

---

## 📁 Files

| File | Purpose |
|---|---|
| `outlook_helpers.py` | COM automation library (v2) |
| `outlook_cli.py` | JSON CLI wrapper |
| `outlook_mock.py` | In-memory test stub |
| `outlook_contacts.py` | Contact groups + priority scoring |
| `contacts.json` | Group/contact config (template — fill in locally) |
| `SKILL.md` | Claude Code skill definition + prompt shots |
| `STYLE_GUIDE.md` | Writing voice rules for drafted emails |
| `CHANGELOG.md` | Version history |
| `tests/` | pytest suite (159 tests) |

---

## 📦 Version

`v2.0.0` — see [CHANGELOG.md](CHANGELOG.md) for full history.  
Rollback: `git checkout v1.0.0 -- outlook_helpers.py`
