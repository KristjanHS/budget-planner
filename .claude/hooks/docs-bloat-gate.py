#!/usr/bin/env python3
"""
docs-bloat-gate — PreToolUse hook on Write/Edit/Bash.

Blocks Write/Edit to gated docs/rule paths (and Bash commands targeting them)
unless the change passes one of: net delta <= 0, '-> see ' pointer in added
text, new '## ' heading introduced, or override sentinel present.

Override sentinel:  <!-- docs-bloat-gate-override: <reason >=10 chars> -->

Design: docs/plans/2026-04-25-doc-bloat-root-causes-brainstorm.md
"""
from __future__ import annotations

import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

GATED_DOCS = {"docs/spec.md", "docs/runbook.md", "docs/refactor-incidents.md"}
GATED_RULES_PREFIX = ".claude/rules/"
OVERRIDE_RE = re.compile(r"<!--\s*docs-bloat-gate-override:\s*(.{10,}?)\s*-->")
POINTER_RE = re.compile(r"→\s*see\s+\S+|->\s*see\s+\S+")
# Bash write-action patterns that actually MUTATE a gated path. We don't
# block on bare references to the path (e.g. `git log docs/spec.md`,
# heredoc commit messages mentioning the path) — only on commands that
# would write to it: redirection (>, >>), `tee`, or `sed -i`.
_GATED_PATH_RE = (
    r"(?:docs/(?:spec|runbook|refactor-incidents)\.md|\.claude/rules/\S+\.md)"
)
BASH_WRITE_PATTERNS = [
    re.compile(r">>?\s*['\"]?" + _GATED_PATH_RE),
    re.compile(r"\btee\b[^|;&]*?\s" + _GATED_PATH_RE),
    re.compile(r"\bsed\b[^|;&]*?(?:-i|--in-place)\b[^|;&]*?\s" + _GATED_PATH_RE),
]
AUDIT_KEYWORDS_RE = re.compile(r"\b(audit|analysis|report|research)\b", re.IGNORECASE)
LOG_REL = "docs/analysis/bloat-overrides.log"


def project_dir() -> Path:
    return Path(os.environ.get("CLAUDE_PROJECT_DIR", "."))


def relpath_or_none(file_path: str) -> str | None:
    try:
        return str(Path(file_path).resolve().relative_to(project_dir().resolve()))
    except (ValueError, OSError):
        return None


def is_gated_path(rel: str) -> bool:
    if rel in GATED_DOCS:
        return True
    if rel.startswith(GATED_RULES_PREFIX) and rel.endswith(".md"):
        return True
    if rel == "CLAUDE.md" and os.environ.get("DOCS_BLOAT_GATE_CLAUDE_MD") == "1":
        return True
    return False


def line_count(s: str) -> int:
    if not s:
        return 0
    return s.count("\n") + (0 if s.endswith("\n") else 1)


def l2_headings(text: str) -> set[str]:
    """Return the set of level-2 heading texts in `text`."""
    return {ln[3:].rstrip() for ln in text.splitlines() if ln.startswith("## ")}


def lines_outside_new_sections(added: str, new_headings: set[str]) -> int:
    """Count lines in `added` not under a brand-new L2 heading.

    Design §4.1: when an Edit introduces an L2 heading not present in the
    on-disk file, lines from that heading until the next L2 heading (or
    EOF) are exempt from net-delta accounting in the introducing edit.
    """
    counted = 0
    in_exempt = False
    for line in added.splitlines():
        if line.startswith("## "):
            head_text = line[3:].rstrip()
            if head_text in new_headings:
                in_exempt = True
                continue
            in_exempt = False
        if not in_exempt:
            counted += 1
    return counted


def append_log(rel: str, delta: int, reason: str, tool: str) -> None:
    log = project_dir() / LOG_REL
    try:
        log.parent.mkdir(parents=True, exist_ok=True)
        ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        with log.open("a", encoding="utf-8") as fh:
            fh.write(f'{ts}  {rel}  {delta:+d} lines  reason="{reason}"  tool={tool}\n')
    except OSError:
        pass


def gate(rel: str, added: str, removed: str, current: str, tool: str) -> int:
    """Return 0 (pass) or 2 (block) per the four §2 pass conditions."""
    delta = line_count(added) - line_count(removed)

    if delta <= 0:
        return 0
    if POINTER_RE.search(added):
        return 0

    # §4.1 strict: lines under a brand-new L2 heading are exempt from
    # net-delta accounting in this introducing edit. Only block if the
    # *non-exempt* delta is still positive.
    new_h = l2_headings(added) - l2_headings(current)
    if new_h:
        carved_delta = lines_outside_new_sections(added, new_h) - line_count(removed)
        if carved_delta <= 0:
            return 0

    m = OVERRIDE_RE.search(added)
    if m:
        append_log(rel, delta, m.group(1).strip(), tool)
        return 0

    sys.stderr.write(
        f"docs-bloat-gate: net +{delta} lines to {rel} without pointer / "
        f"removal / new section / override. Add a '-> see X §Y' pointer, pair "
        f"with a removal, or add <!-- docs-bloat-gate-override: <reason> -->.\n"
    )
    return 2


def handle_write_edit(payload: dict, tool: str) -> int:
    ti = payload.get("tool_input", {})
    file_path = ti.get("file_path", "")
    if not file_path:
        return 0
    rel = relpath_or_none(file_path)
    if rel is None:
        return 0

    # W1.1 — block new audit-style files at docs/ root (depth=1, not in GATED_DOCS).
    # Existing reference docs and subdirs (analysis/, plans/, human/, research/) bypass.
    if (
        tool == "Write"
        and rel.startswith("docs/")
        and rel.count("/") == 1
        and rel.endswith(".md")
        and rel not in GATED_DOCS
    ):
        content = ti.get("content", "")
        if AUDIT_KEYWORDS_RE.search(content) and not OVERRIDE_RE.search(content):
            sys.stderr.write(
                "docs-bloat-gate: new docs/*.md root file with "
                "audit/analysis/report/research keywords. Move to "
                "docs/analysis/ (gitignored), use a docs/ subdir, or add "
                "<!-- docs-bloat-gate-override: <reason> -->.\n"
            )
            return 2

    if not is_gated_path(rel):
        return 0

    if tool == "Edit":
        added = ti.get("new_string", "")
        removed = ti.get("old_string", "")
    else:  # Write
        added = ti.get("content", "")
        try:
            removed = (project_dir() / rel).read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError):
            removed = ""

    try:
        current = (project_dir() / rel).read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError):
        current = ""

    return gate(rel, added, removed, current, tool)


def handle_bash(payload: dict) -> int:
    cmd = payload.get("tool_input", {}).get("command", "")
    if not cmd:
        return 0
    matches = [m for pat in BASH_WRITE_PATTERNS for m in pat.findall(cmd)]
    if not matches:
        return 0
    m = OVERRIDE_RE.search(cmd)
    if m:
        append_log("|".join(matches), 0, m.group(1).strip(), "Bash")
        return 0
    sys.stderr.write(
        "docs-bloat-gate: Bash command writes to a gated doc/rule path "
        f"({', '.join(sorted(set(matches)))}). Use Write/Edit (gate "
        "computes delta), or include "
        "<!-- docs-bloat-gate-override: <reason> --> in the command.\n"
    )
    return 2


def main() -> int:
    try:
        payload = json.load(sys.stdin)
    except (json.JSONDecodeError, OSError):
        return 0
    tool = payload.get("tool_name", "")
    if tool in ("Write", "Edit"):
        return handle_write_edit(payload, tool)
    if tool == "Bash":
        return handle_bash(payload)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        sys.exit(0)
