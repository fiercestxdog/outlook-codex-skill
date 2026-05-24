"""
conftest.py — pytest configuration for outlook_cli tests.

All tests use subprocess so state resets automatically (each CLI call is a
fresh process). No Outlook / pywin32 installation required.

Run from Codex_Skills/:
    pytest tests/ -v
    pytest tests/ -v -k "TestMeeting"      # one class
    pytest tests/ -v --tb=short            # compact tracebacks
"""
import os
import sys
from pathlib import Path

# Ensure the Codex_Skills dir is on sys.path so direct-import tests work too
SKILLS_DIR = Path(__file__).parent.parent
sys.path.insert(0, str(SKILLS_DIR))

# Activate the mock for any test that imports outlook_helpers directly
os.environ.setdefault("OUTLOOK_MOCK", "1")
