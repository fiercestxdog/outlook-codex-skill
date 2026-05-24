# Codex Skills: Weekly Brief & Outlook Automation

A robust, local-first automation layer for Microsoft Outlook. This project provides a JSON-emitting CLI that allows an LLM (the "skill") to reason over your calendar, mail, and tasks to generate high-quality weekly briefs and perform complex triage.

## 🚀 Quick Start

### 1. Prerequisites
- **OS**: Windows (with Outlook installed and running).
- **Python**: 3.8+
- **Dependency**: `pip install pywin32`

### 2. Installation & Setup
Clone the repository and install the dependencies:
```powershell
pip install pywin32
# Ensure you are in the project directory
python outlook_cli.py init-categories
```

### 3. Verify Connection
Run a few commands to ensure the COM bridge is working:
```powershell
# Show your upcoming work schedule (honors RDOs/Holidays)
python outlook_cli.py workdays --days-ahead 14

# List the last hour of emails
python outlook_cli.py recent --hours 1
```

---

## 🛠️ CLI Commands & Examples

The CLI emits JSON to `stdout`. Every command is designed to be parsed by an LLM or script.

### Reading Data
| Command | Description | Example |
| :--- | :--- | :--- |
| `recent` | Last N hours of mail | `python outlook_cli.py recent --hours 48 --unread` |
| `events` | Calendar window | `python outlook_cli.py events --days-ahead 7` |
| `free` | Open time slots | `python outlook_cli.py free --duration 30` |
| `thread` | Full email context | `python outlook_cli.py thread <EntryID>` |
| `search` | Subject search | `python outlook_cli.py search "Budget"` |

### Taking Action
| Command | Description | Example |
| :--- | :--- | :--- |
| `draft` | Create a draft | `python outlook_cli.py draft --to bob@x.com --subject "Review" --body "Please check."` |
| `reply` | Reply to thread | `python outlook_cli.py reply <EntryID> --body "Got it."` |
| `task` | Create a task | `python outlook_cli.py task --subject "Fix Bug" --due 2026-06-01` |
| `mark` | Flag/Move/Read | `python outlook_cli.py mark <EntryID> --flag --read` |
| `categorize`| Set Categories | `python outlook_cli.py categorize <EntryID> --add "Action item"` |
| `radar` | Follow-up Radar | `python outlook_cli.py radar` |
| `focus` | Auto-book Focus | `python outlook_cli.py focus --hours 2` |
| `analytics` | Time Analytics | `python outlook_cli.py analytics` |
| `pulse` | Meeting Health | `python outlook_cli.py pulse` |

---

## 🧬 Personal Style & Ghostwriting
The assistant uses a `STYLE_GUIDE.md` to ensure all drafts and replies match your personal voice. It avoids "AI-isms" and follows your specific "Writing DNA."

---

## 🧠 Architecture

- **`outlook_helpers.py`**: The engine. Uses `pywin32` to talk to Outlook via COM. Handles timezones, work schedules, and JSON normalization.
- **`outlook_cli.py`**: The interface. A thin wrapper that maps CLI arguments to helper functions and emits JSON.
- **`SKILL.md`**: The brain. Contains the system prompts and "guardrails" for the LLM to act as a senior executive assistant.

## 📅 Work Schedule & Holidays
The system is aware of:
- **9/80 Schedules**: Every other Friday off (configured via `off_friday_anchor` in `outlook_helpers.py`).
- **US Federal Holidays**: Preloaded for 2026-2027.
- **Custom PTO**: Can be added to the `WORK_SCHEDULE` dictionary in `outlook_helpers.py`.

---

## 🔒 Security & Privacy
- **Local Only**: All data stays on your machine. No cloud APIs or third-party OAuth flows.
- **Read-Heavy**: The LLM is instructed to treat destructive actions (delete, send) as approval-gated.
- **Minimal Surface**: Only the necessary fields (Subject, Sender, Date, etc.) are pulled by default to keep the context window efficient.
