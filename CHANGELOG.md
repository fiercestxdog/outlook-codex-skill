# Changelog

All notable changes to this project will be documented here.

Format: [Keep a Changelog](https://keepachangelog.com/en/1.0.0/)  
Versioning: [Semantic Versioning](https://semver.org/spec/v2.0.0.html)

---

## [2.0.0] — 2026-05-25

### Added
- `_check_classic_outlook()` — raises `RuntimeError` with actionable fix instructions
  if New Outlook (Win11 web-based default, COM-incompatible) is detected. Checks that
  `app.Version` major is 14/15/16 (Classic Outlook range). Called from `_application()`
  so it gates every helper function automatically.
- `_detect_date_fmt()` — derives the correct Outlook `Restrict()` date format from the
  Windows system locale via `GetLocaleInfoW`. Maps Windows picture tokens (`MM`, `yyyy`,
  `tt`, etc.) to `strftime` directives. Provides a one-liner diagnostic for non-US
  Windows debugging. Falls back to US format on any failure.
- `__version__ = "2.0.0"` added to `outlook_helpers.py`, `outlook_cli.py`,
  `outlook_contacts.py`.
- `CHANGELOG.md` (this file).

### Changed
- **`CoInitialize` / `CoUninitialize`** — moved from per-call inside `_application()`
  to module-level init with `atexit.register(CoUninitialize)`. Eliminates COM apartment
  leak when the module is embedded in a long-running process (web server, task queue,
  background service). Worker threads still need their own paired calls.
- **`get_calendar_events()` docstring** — `IncludeRecurrences` ordering is now
  documented as a mandatory, immutable three-step sequence (Sort → IncludeRecurrences →
  Restrict). Inline step labels added. Explains the silent failure modes (dropped events,
  infinite loops) that result from reordering.
- **Module docstring** — marked `(v2)`, added Classic Outlook EOL warning, M365 security
  popup regression note (post-Oct 2024) with three workarounds, `pywin32>=306` minimum
  version, and locale diagnostic reference.

### Security
- Documented M365 programmatic-access popup regression (Oct 2024): old
  `ObjectModelGuard` registry bypass no longer suppresses it. Three workarounds
  documented in module docstring.

### Known Issues / Limitations
- `_DATE_FMT` is still hardcoded to US locale (`%m/%d/%Y %I:%M %p`). Use
  `_detect_date_fmt()` to get the correct value for non-US Windows; automatic
  locale detection at import time is not yet applied.
- Classic Outlook COM support has a sunset dependency — Microsoft is migrating all
  users to New Outlook; send/draft operations will eventually need to move to
  Microsoft Graph (OAuth).

---

## [1.0.0] — 2026-05-24

### Added
- `outlook_helpers.py` — full COM automation library: mail (read, search, thread,
  draft, reply, forward, flag, move, categorize), calendar (events, free slots,
  meeting health, analytics, focus-block booking), tasks, contacts (GAL resolve),
  reschedule/cancel/respond to invites, save attachments.
- `outlook_cli.py` — JSON-emitting argparse CLI over `outlook_helpers.py`. 30+
  subcommands. `OUTLOOK_MOCK=1` env var swaps in the in-memory test stub.
- `outlook_mock.py` — in-memory COM stub: 8 fixture emails (EMAIL001–EMAIL008),
  6 calendar events (EVT001–EVT006), 2 tasks. Full `outlook_helpers.py` API parity.
  `reset_state()` for in-process test isolation.
- `outlook_contacts.py` — pure-Python contact group management (zero COM dependency).
  Groups: `direct_reports`, `supervisors`, `peers`, `program_leads`. Priority tier
  scoring: supervisor > program_lead > direct_report > peer > unknown.
  `enrich_messages()` adds `tier` + `priority_signals` (`always_surface`,
  `blocker_flag`, `review_flag`, `stale_Nh`) and sorts supervisor-first.
  Natural-language group resolution ("my team" → direct_reports, "boss" → supervisors).
- `contacts.json` — template (empty groups). **Populate locally; never push real data.**
- `SKILL.md` — skill definition with full CLI reference (30+ commands, `entry_id`
  notes), 7 prompt-shot use cases (Monday Brief, Meeting Context, Conflict Negotiator,
  Inbox Triage, Digital Twin Ghostwriter, Meeting Scheduler, Time Audit), relationship-
  aware priority rules table, group dispatch shortcuts, workflow with `enrich`.
- `STYLE_GUIDE.md` — writing voice, banned words, tone rules for drafted emails.
- `PROJECT.md` — project overview and setup guide.
- `tests/conftest.py` — sets `OUTLOOK_MOCK=1` and adds `Codex_Skills/` to `sys.path`.
- `tests/test_cli.py` — 105 tests covering all 30+ CLI subcommands. Subprocess-based
  (auto-resets state) + in-process round-trip for stateful flows (draft → send).
- `tests/test_contacts.py` — 54 tests: load, group CRUD, `who_is`, `priority_tier`,
  `enrich_messages`, `resolve`, CLI group commands, `--to-group`, `--attendees-group`.
- `tests/fixtures/contacts.json` — 5 contacts, 4 groups for test isolation.
- `.gitignore` — excludes `__pycache__/`, `.pytest_cache/`, `.env`, `*.py[cod]`.

---

[2.0.0]: https://github.com/fiercestxdog/outlook-codex-skill/compare/v1.0.0...v2.0.0
[1.0.0]: https://github.com/fiercestxdog/outlook-codex-skill/releases/tag/v1.0.0
