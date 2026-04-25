#!/usr/bin/env python3
"""
external-process-test-gate — PreToolUse Write/Edit on tests/**/*.py.

Blocks new/modified test code that drives external systems (docker, kubectl,
headless browsers, HTTP services) outside the manual tier. Forces such tests
into tests/manual/ so they don't run on every `pytest tests/ -q`.

Override sentinel:  <!-- external-process-test-gate-override: <reason ≥10> -->

Path scope:
  Fires on tests/**/*.py EXCLUDING:
   - tests/manual/**          (the carve-out: external-process tests live here)
   - tests/__init__.py        (package marker, never test bodies)
   - tests/conftest.py        (session-fixture file; project-specific recalc
                               subprocess use is established and scoped to
                               LibreOffice, not external services)

Detected patterns (scope: heavyweight external systems, per
~/.claude/rules/testing.md "External-process tests"):
  1. import of `docker` / `playwright` / `selenium`
  2. `requests.<get|post|put|delete|patch|head|options>(` HTTP call
  3. `subprocess.<run|Popen|check_call|check_output|call|getoutput>(` whose
     first arg starts with "docker" / "kubectl" / "playwright" / "selenium"
  4. `os.system("docker...")` / `os.system("kubectl...")` etc.

Why narrower than the W1.4 audit's `subprocess.run` blanket:
  This project legitimately runs `subprocess.run([...recalc.py...])` in many
  tests (LibreOffice formula recalc — the project's own toolchain, not an
  external service). A blanket subprocess.run match would false-positive on
  every recalc test edit. The narrowed match enforces the source rule's
  actual scope (docker / kubectl / headless browser) without that noise.

Source: docs/plans/2026-04-25-prose-rules-to-hooks-audit.md (W1.4).
"""
from __future__ import annotations

import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

OVERRIDE_RE = re.compile(
    r"<!--\s*external-process-test-gate-override:\s*(.{10,}?)\s*-->"
)

EXEMPT_FILES = {"tests/__init__.py", "tests/conftest.py"}
EXEMPT_PREFIX = "tests/manual/"

PATTERNS: list[tuple[str, re.Pattern[str]]] = [
    (
        "docker/playwright/selenium import",
        re.compile(
            r"^\s*(?:from|import)\s+(?:docker|playwright|selenium)\b",
            re.MULTILINE,
        ),
    ),
    (
        "requests HTTP call",
        re.compile(
            r"\brequests\.(?:get|post|put|delete|patch|head|options)\s*\("
        ),
    ),
    (
        "subprocess to external CLI",
        re.compile(
            r"\bsubprocess\.(?:run|Popen|check_call|check_output|call|getoutput)"
            r"\s*\(\s*\[?\s*['\"]?(?:docker|kubectl|playwright|selenium)\b"
        ),
    ),
    (
        "os.system to external CLI",
        re.compile(
            r"\bos\.system\s*\(\s*['\"](?:docker|kubectl|playwright|selenium)\b"
        ),
    ),
]

LOG_REL = "docs/analysis/external-process-test-overrides.log"


def project_dir() -> Path:
    return Path(os.environ.get("CLAUDE_PROJECT_DIR", "."))


def relpath_or_none(file_path: str) -> str | None:
    try:
        return str(Path(file_path).resolve().relative_to(project_dir().resolve()))
    except (ValueError, OSError):
        return None


def is_gated_path(rel: str) -> bool:
    if not rel.startswith("tests/") or not rel.endswith(".py"):
        return False
    if rel in EXEMPT_FILES:
        return False
    if rel.startswith(EXEMPT_PREFIX):
        return False
    return True


def find_matches(text: str) -> list[str]:
    hits: list[str] = []
    for label, pat in PATTERNS:
        if pat.search(text):
            hits.append(label)
    return hits


def append_log(rel: str, hits: list[str], reason: str, tool: str) -> None:
    log = project_dir() / LOG_REL
    try:
        log.parent.mkdir(parents=True, exist_ok=True)
        ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        with log.open("a", encoding="utf-8") as fh:
            fh.write(
                f'{ts}  {rel}  hits=[{",".join(hits)}]  '
                f'reason="{reason}"  tool={tool}\n'
            )
    except OSError:
        pass


def handle(payload: dict, tool: str) -> int:
    ti = payload.get("tool_input", {}) or {}
    file_path = ti.get("file_path", "")
    if not file_path:
        return 0
    rel = relpath_or_none(file_path)
    if rel is None or not is_gated_path(rel):
        return 0

    if tool == "Edit":
        added = ti.get("new_string", "") or ""
    else:  # Write
        added = ti.get("content", "") or ""
    if not added:
        return 0

    hits = find_matches(added)
    if not hits:
        return 0

    m = OVERRIDE_RE.search(added)
    if m:
        append_log(rel, hits, m.group(1).strip(), tool)
        return 0

    sys.stderr.write(
        f"external-process-test-gate: {rel} adds external-process patterns "
        f"({', '.join(hits)}) outside the manual tier.\n"
        "Per ~/.claude/rules/testing.md: external-process tests (docker, "
        "kubectl, headless browser, live HTTP) belong in tests/manual/ — "
        "they run on demand, not in `pytest tests/ -q`.\n"
        "Unblock: move the test under tests/manual/, or add "
        "<!-- external-process-test-gate-override: <reason ≥10 chars> --> "
        "in the new content.\n"
    )
    return 2


def main() -> int:
    try:
        payload = json.load(sys.stdin)
    except (json.JSONDecodeError, OSError):
        return 0
    tool = payload.get("tool_name", "")
    if tool not in ("Write", "Edit"):
        return 0
    return handle(payload, tool)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        sys.exit(0)
