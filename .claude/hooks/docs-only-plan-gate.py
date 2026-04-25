#!/usr/bin/env python3
"""
docs-only-plan-gate — PreToolUse Bash gate + PostToolUse Bash stamp updater.

PreToolUse Bash:
  Blocks `git mv .*docs/plans/.*archived/` unless one holds:
   - a pytest-pass stamp is fresh (mtime < STALE_S seconds old), or
   - the command contains <!-- docs-only-plan-gate-override: <reason ≥10> -->.

PostToolUse Bash:
  When the command invoked `pytest tests/...` and exit code is 0, write/refresh
  the stamp so subsequent archive moves pass without re-running pytest.

Stamp file: ~/.claude/sessions/<md5(CLAUDE_PROJECT_DIR)[:8]>/pytest-pass-stamp

Why this gate exists:
  `tests/test_config_sheet.py::test_config_named_range_targets_match_spec_doc`
  parses the named-range table in docs/spec.md. A docs-only plan whose only
  change is restructuring that table breaks pytest, but the user — seeing
  "diff is docs/**-only" — may skip the pytest step. The gate enforces it.

Source: docs/plans/archived/2026-04-25-prose-rules-to-hooks-audit.md (W1.2).
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

STALE_S = 1800  # 30 minutes
OVERRIDE_RE = re.compile(
    r"<!--\s*docs-only-plan-gate-override:\s*(.{10,}?)\s*-->"
)
# `git mv` of a plan into docs/plans/archived/ — order-tolerant; both source
# and dest can appear in either token order. Match if both `docs/plans/` and
# `archived/` appear within a single git mv invocation.
GIT_MV_RE = re.compile(
    r"\bgit\s+mv\b[^|;&]*\bdocs/plans/[^|;&]*\barchived/[^|;&]*"
)
# Detect a real `pytest tests/...` invocation (any pytest binary, any args).
# Excludes `pytest --help` and similar non-test invocations.
# `\bpytest\b` matches both bare `pytest` and `.venv/bin/pytest` (path-style
# launchers); the `\btests/` requirement after rules out `pytest --help`.
PYTEST_RE = re.compile(r"\bpytest\b[^|;&]*\btests/")
LOG_REL = "docs/analysis/docs-only-plan-overrides.log"


def project_dir() -> Path:
    return Path(os.environ.get("CLAUDE_PROJECT_DIR", "."))


def stamp_path() -> Path:
    pdir = str(project_dir().resolve())
    project_hash = hashlib.md5(pdir.encode()).hexdigest()[:8]
    return Path.home() / ".claude" / "sessions" / project_hash / "pytest-pass-stamp"


def stamp_fresh() -> bool:
    p = stamp_path()
    try:
        age = time.time() - p.stat().st_mtime
    except OSError:
        return False
    return age <= STALE_S


def write_stamp() -> None:
    p = stamp_path()
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(
            datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ") + "\n",
            encoding="utf-8",
        )
    except OSError:
        pass


def append_log(cmd_summary: str, reason: str) -> None:
    log = project_dir() / LOG_REL
    try:
        log.parent.mkdir(parents=True, exist_ok=True)
        ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        with log.open("a", encoding="utf-8") as fh:
            fh.write(f'{ts}  {cmd_summary}  reason="{reason}"\n')
    except OSError:
        pass


def handle_pre_tool_use(payload: dict) -> int:
    cmd = payload.get("tool_input", {}).get("command", "")
    if not cmd or not GIT_MV_RE.search(cmd):
        return 0

    if stamp_fresh():
        return 0

    m = OVERRIDE_RE.search(cmd)
    if m:
        append_log("git mv (override)", m.group(1).strip())
        return 0

    age_msg = "absent" if not stamp_path().exists() else "stale (>30 min)"
    sys.stderr.write(
        "docs-only-plan-gate: archiving a plan, but pytest-pass stamp is "
        f"{age_msg}.\n"
        "Per .claude/rules/plan-hygiene.md: docs-only plan archives must run "
        "pytest first (test_config_sheet.py parses docs/spec.md tables; a "
        "docs-only restructure can break it without touching code).\n"
        "Unblock: run `pytest tests/ -q` (the PostToolUse hook will refresh "
        "the stamp on success), or include "
        "<!-- docs-only-plan-gate-override: <reason ≥10 chars> --> in the "
        "command.\n"
    )
    return 2


def post_tool_use_was_passing_pytest(payload: dict) -> bool:
    """True iff the just-finished Bash invocation was `pytest tests/...` exit 0."""
    cmd = payload.get("tool_input", {}).get("command", "")
    if not cmd or not PYTEST_RE.search(cmd):
        return False
    resp = payload.get("tool_response", {}) or {}
    if not isinstance(resp, dict):
        return False
    # Prefer explicit exit_code when present.
    exit_code = resp.get("exit_code")
    if isinstance(exit_code, int):
        return exit_code == 0
    # Fallback: interrupted means not a clean pass.
    if resp.get("interrupted"):
        return False
    # Last resort: stdout contains a pytest summary line with "passed" and no
    # "failed"/"error" markers. Conservative — false negatives over false
    # positives (false positive would skip a needed pytest run).
    stdout = resp.get("stdout", "") or ""
    if not stdout:
        return False
    if re.search(r"\b\d+\s+failed\b|\berror\b", stdout, re.IGNORECASE):
        return False
    return bool(re.search(r"\b\d+\s+passed\b", stdout))


def handle_post_tool_use(payload: dict) -> int:
    if post_tool_use_was_passing_pytest(payload):
        write_stamp()
    return 0


def main() -> int:
    try:
        payload = json.load(sys.stdin)
    except (json.JSONDecodeError, OSError):
        return 0
    if payload.get("tool_name") != "Bash":
        return 0
    event = payload.get("hook_event_name", "")
    if event == "PreToolUse":
        return handle_pre_tool_use(payload)
    if event == "PostToolUse":
        return handle_post_tool_use(payload)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        sys.exit(0)
