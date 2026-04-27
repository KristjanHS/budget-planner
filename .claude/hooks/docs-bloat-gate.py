#!/usr/bin/env python3
"""
docs-bloat-gate v2 — PreToolUse hook on Write/Edit/Bash.

Three signals (any blocks):
  S1: char-delta exceeds tier cap (rule<50ln=150, doc=800, spec=2000)  bypassable
  S2: AI-slop stoplist phrase present in net-added text                UNBYPASSABLE
  S3: lexical density < 0.45 on >100-char addition                     UNBYPASSABLE

S1 bypasses: brand-new L2 heading (<=30 lines / 1500 chars exempt) or
override sentinel (hard cap N=1 per session, reason >=30 chars, no
anti-stopwords).

Override sentinel: <!-- docs-bloat-gate-override: <reason >=30 chars> -->

Design: docs/plans/archived/2026-04-26-docs-bloat-gate-redesign.md
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

# --- Gated paths ---------------------------------------------------------
GATED_DOCS = {"docs/spec.md", "docs/runbook.md", "docs/refactor-incidents.md"}
GATED_RULES_PREFIX = ".claude/rules/"

# --- Bash write-pattern detection ----------------------------------------
_GATED_PATH_RE = (
    r"(?:docs/(?:spec|runbook|refactor-incidents)\.md|\.claude/rules/\S+\.md)"
)
BASH_WRITE_PATTERNS = [
    re.compile(r">>?\s*['\"]?" + _GATED_PATH_RE),
    re.compile(r"\btee\b[^|;&]*?\s" + _GATED_PATH_RE),
    re.compile(r"\bsed\b[^|;&]*?(?:-i|--in-place)\b[^|;&]*?\s" + _GATED_PATH_RE),
]
AUDIT_KEYWORDS_RE = re.compile(r"\b(audit|analysis|report|research)\b", re.IGNORECASE)

# --- Sentinel + log ------------------------------------------------------
OVERRIDE_RE = re.compile(r"<!--\s*docs-bloat-gate-override:\s*(.+?)\s*-->", re.DOTALL)
ANTI_STOPWORDS = {"update", "rewrite", "improvement", "cleanup", "fix"}
LOG_REL = "docs/analysis/bloat-overrides.log"
SENTINEL_CAP = 1
STATE_DIR = Path.home() / ".claude" / "state"
SENTINEL_REASON_MIN = 30

# --- Tier thresholds -----------------------------------------------------
TIER_CAPS = {"rule": 150, "doc": 800, "spec": 2000}

# --- L2 heading exemption ------------------------------------------------
NEW_SECTION_LINE_EXEMPT = 30
NEW_SECTION_CHAR_EXEMPT = 1500

# --- S3 thresholds -------------------------------------------------------
DENSITY_THRESHOLD = 0.45
DENSITY_DELTA_GUARD = 100

# --- S2 stoplist (~30 patterns) ------------------------------------------
SLOP_WORDS = {
    # transition adverbs
    "furthermore", "moreover", "additionally", "consequently",
    "subsequently", "accordingly", "nevertheless", "notwithstanding",
    # hedge adverbs
    "fundamentally", "essentially", "ultimately",
    # AI tells (single-word)
    "comprehensive", "robust", "seamless", "leverage", "utilize",
}
SLOP_PHRASES = (
    # hedges (multi-word)
    "it's worth noting", "it should be noted", "it is important to",
    "that said", "having said that", "in essence", "at its core",
    "while this may vary", "in the context of",
    # AI tells (multi-word)
    "delve into", "navigate the", "landscape of", "realm of", "tapestry of",
)

# ~150 English function words for S3 lexical-density stoplist.
FUNCTION_WORDS = {
    # articles
    "a", "an", "the",
    # pronouns
    "i", "me", "my", "mine", "we", "us", "our", "ours",
    "you", "your", "yours", "he", "him", "his",
    "she", "her", "hers", "it", "its",
    "they", "them", "their", "theirs",
    "this", "that", "these", "those",
    "who", "whom", "whose", "which", "what",
    # conjunctions / subordinators
    "and", "or", "but", "nor", "so", "yet", "for",
    "if", "then", "than", "while", "although", "though",
    "because", "since", "unless", "until", "when",
    "where", "whether", "how", "why", "as",
    # prepositions
    "in", "on", "at", "to", "of", "by", "with",
    "from", "into", "onto", "upon", "about", "above",
    "across", "after", "against", "along", "among",
    "around", "before", "behind", "below", "beneath",
    "beside", "between", "beyond", "during", "except",
    "inside", "near", "outside", "over", "through",
    "throughout", "toward", "towards", "under", "underneath",
    "up", "down", "off", "out",
    # auxiliaries / be / have / do / modals
    "be", "am", "is", "are", "was", "were", "been", "being",
    "have", "has", "had", "having",
    "do", "does", "did", "doing",
    "will", "would", "shall", "should", "can", "could",
    "may", "might", "must", "ought",
    # determiners / quantifiers
    "no", "not", "all", "any", "some", "none", "both",
    "each", "every", "much", "many", "more", "most",
    "less", "least", "few", "several", "such", "same",
    "other", "another", "own",
    # discourse / fillers
    "only", "just", "very", "too", "also", "even", "again",
    "here", "there", "now", "well", "like", "yes", "ok",
    # contraction stubs (after \w+ tokenization)
    "s", "t", "re", "ve", "ll", "d", "m",
}


# --- Helpers -------------------------------------------------------------
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


def file_tier(current_lines: int) -> str:
    if current_lines < 50:
        return "rule"
    if current_lines < 500:
        return "doc"
    return "spec"


def _safe_sid(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]", "_", s)[:64] or "default"


def session_id() -> str:
    sid = os.environ.get("CLAUDE_SESSION_ID")
    if sid:
        return _safe_sid(sid)
    transcript = os.environ.get("CLAUDE_TRANSCRIPT_PATH")
    if transcript:
        return hashlib.md5(transcript.encode()).hexdigest()[:12]
    seed = f"{project_dir()}|{datetime.now(timezone.utc).strftime('%Y%m%d%H')}"
    return hashlib.md5(seed.encode()).hexdigest()[:12]


def sentinel_counter_path() -> Path:
    return STATE_DIR / f"docs-bloat-gate-sentinels-{session_id()}.txt"


def read_sentinel_count() -> int:
    p = sentinel_counter_path()
    try:
        return int(p.read_text().strip() or "0")
    except (OSError, ValueError):
        return 0


def increment_sentinel_count() -> int:
    p = sentinel_counter_path()
    new = read_sentinel_count() + 1
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(str(new))
    except OSError:
        pass
    return new


# --- L2-heading carve-out ------------------------------------------------
def l2_headings(text: str) -> set[str]:
    return {ln[3:].rstrip() for ln in text.splitlines() if ln.startswith("## ")}


def carved_char_delta(added: str, removed: str, current: str) -> int:
    """Char-delta with brand-new L2 sections exempted (<=30 lines / 1500 chars).

    Lines from a brand-new ``## `` heading until the next L2 heading (or EOF)
    are exempt up to whichever limit hits first.
    """
    delta = len(added) - len(removed)
    new_h = l2_headings(added) - l2_headings(current)
    if not new_h:
        return delta

    in_exempt = False
    exempt_lines = 0
    exempt_chars = 0
    for line in added.splitlines(keepends=True):
        if line.startswith("## "):
            head_text = line[3:].rstrip("\r\n")
            if head_text in new_h:
                in_exempt = True
                exempt_lines = 0
                exempt_chars = 0
            else:
                in_exempt = False
        if in_exempt:
            if (
                exempt_lines < NEW_SECTION_LINE_EXEMPT
                and exempt_chars + len(line) <= NEW_SECTION_CHAR_EXEMPT
            ):
                exempt_lines += 1
                exempt_chars += len(line)
                delta -= len(line)
            else:
                in_exempt = False
    return delta


# --- S2: slop detection --------------------------------------------------
_SLOP_WORD_RE = re.compile(
    r"\b(" + "|".join(re.escape(w) for w in sorted(SLOP_WORDS)) + r")\b",
    re.IGNORECASE,
)


def slop_hit_counts(text: str) -> Counter[str]:
    """Return a Counter of slop-pattern occurrences in `text`.

    Counts (not just presence) so that net-added detection in
    `gate_text_edit` catches "added one more occurrence" cases — a
    set-difference would miss the second instance when the first already
    appears in `removed`.
    """
    counts: Counter[str] = Counter()
    if not text:
        return counts
    for m in _SLOP_WORD_RE.finditer(text):
        counts[m.group(1).lower()] += 1
    lower = text.lower()
    for ph in SLOP_PHRASES:
        n, start = 0, 0
        while (idx := lower.find(ph, start)) != -1:
            n += 1
            start = idx + len(ph)
        if n:
            counts[ph] += n
    return counts


# --- S3: lexical density -------------------------------------------------
def lexical_density(added: str) -> tuple[float, int]:
    tokens = re.findall(r"\w+", added.lower())
    if not tokens:
        return 1.0, 0
    content = sum(1 for t in tokens if t not in FUNCTION_WORDS)
    return content / len(tokens), len(tokens)


# --- Override log --------------------------------------------------------
def now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def append_log(entry: dict) -> None:
    log = project_dir() / LOG_REL
    try:
        log.parent.mkdir(parents=True, exist_ok=True)
        with log.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(entry, separators=(",", ":")) + "\n")
    except OSError:
        pass


def validate_reason(reason: str) -> str | None:
    r = reason.strip()
    if len(r) < SENTINEL_REASON_MIN:
        return f"reason too short ({len(r)} < {SENTINEL_REASON_MIN} chars)"
    lower = r.lower()
    for sw in ANTI_STOPWORDS:
        if re.search(rf"\b{sw}\b", lower):
            return f"anti-stopword in reason: {sw!r}"
    return None


# --- Gate ---------------------------------------------------------------
def gate_text_edit(rel: str, added: str, removed: str, current: str, tool: str) -> int:
    """Return 0 (pass) or 2 (block) per design §4 signal composition."""
    char_delta = len(added) - len(removed)

    # S2: slop list (UNBYPASSABLE) - compute net-added phrases only.
    added_counts = slop_hit_counts(added)
    removed_counts = slop_hit_counts(removed)
    new_hits = sorted(
        p for p, n in added_counts.items() if n > removed_counts.get(p, 0)
    )
    if new_hits:
        sys.stderr.write(
            f"docs-bloat-gate: AI-slop phrases detected: {', '.join(new_hits)}. "
            "UNBYPASSABLE.\nRemove these phrases - they signal LLM-generated "
            "bloat regardless of edit size.\n"
        )
        return 2

    # S3: lexical density (UNBYPASSABLE), gated on char_delta > 100.
    if char_delta > DENSITY_DELTA_GUARD:
        density, _ = lexical_density(added)
        if density < DENSITY_THRESHOLD:
            sys.stderr.write(
                f"docs-bloat-gate: lexical density {density:.2f} below "
                f"{DENSITY_THRESHOLD} on {char_delta}-char addition. "
                "Function-word ratio too high - prose is filler-dense. "
                "UNBYPASSABLE.\nTighten by removing hedges, articles, and "
                "transition phrases.\n"
            )
            return 2

    # S1: tier-cap (bypassable).
    tier = file_tier(line_count(current))
    cap = TIER_CAPS[tier]

    if char_delta <= cap:
        return 0

    # L2-heading carve-out.
    new_h = l2_headings(added) - l2_headings(current)
    if new_h:
        carved = carved_char_delta(added, removed, current)
        if carved <= cap:
            append_log({
                "ts": now_iso(),
                "file": rel,
                "signal": "S1",
                "bypass_type": "l2_heading",
                "char_delta": char_delta,
                "carved_delta": carved,
                "headings": sorted(new_h),
                "session_id": session_id(),
                "tool": tool,
            })
            return 0

    # Sentinel bypass (cap-gated).
    used = read_sentinel_count()
    m = OVERRIDE_RE.search(added)
    if m:
        if used >= SENTINEL_CAP:
            sys.stderr.write(
                f"docs-bloat-gate: session sentinel cap exhausted "
                f"({used}/{SENTINEL_CAP} used).\n"
                "All gated-doc edits hard-block until next session.\n"
            )
            return 2
        reason = m.group(1).strip()
        err = validate_reason(reason)
        if err:
            sys.stderr.write(
                f"docs-bloat-gate: override sentinel rejected ({err}). "
                f"Reason must be >={SENTINEL_REASON_MIN} chars and contain no "
                f"anti-stopwords ({', '.join(sorted(ANTI_STOPWORDS))}).\n"
            )
            return 2
        new_count = increment_sentinel_count()
        append_log({
            "ts": now_iso(),
            "file": rel,
            "signal": "S1",
            "bypass_type": "sentinel",
            "reason": reason,
            "char_delta": char_delta,
            "session_id": session_id(),
            "tool": tool,
        })
        sys.stderr.write(
            f"docs-bloat-gate: sentinel burned ({new_count}/{SENTINEL_CAP} "
            "this session). Edit allowed.\n"
        )
        return 0

    sys.stderr.write(
        f"docs-bloat-gate: char_delta={char_delta} exceeds cap={cap} "
        f"({tier} tier, file={rel}).\nOptions:\n"
        "  1. Revise edit to fit within budget\n"
        "  2. Add L2 heading if introducing genuinely new section "
        f"(<={NEW_SECTION_LINE_EXEMPT} lines / "
        f"{NEW_SECTION_CHAR_EXEMPT} chars exempt)\n"
        "  3. Override: <!-- docs-bloat-gate-override: "
        f"<reason >={SENTINEL_REASON_MIN} chars, no stopwords> -->\n"
        f"     WARNING: {SENTINEL_CAP} sentinel per session. After use, all "
        "further edits\n"
        "     to gated docs hard-block until next session.\n"
        f"Sentinels used this session: {used}/{SENTINEL_CAP}.\n"
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

    # W1.1 - block new audit-style files at docs/ root (depth=1, not gated).
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
                f"<!-- docs-bloat-gate-override: <reason >={SENTINEL_REASON_MIN} "
                "chars> -->.\n"
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

    return gate_text_edit(rel, added, removed, current, tool)


def handle_bash(payload: dict) -> int:
    cmd = payload.get("tool_input", {}).get("command", "")
    if not cmd:
        return 0
    matches = [m for pat in BASH_WRITE_PATTERNS for m in pat.findall(cmd)]
    if not matches:
        return 0

    used = read_sentinel_count()
    m = OVERRIDE_RE.search(cmd)
    if m:
        if used >= SENTINEL_CAP:
            sys.stderr.write(
                f"docs-bloat-gate: session sentinel cap exhausted "
                f"({used}/{SENTINEL_CAP} used). Bash write blocked.\n"
            )
            return 2
        reason = m.group(1).strip()
        err = validate_reason(reason)
        if err:
            sys.stderr.write(
                f"docs-bloat-gate: override sentinel rejected ({err}).\n"
            )
            return 2
        new_count = increment_sentinel_count()
        append_log({
            "ts": now_iso(),
            "file": "|".join(sorted(set(matches))),
            "signal": "S1",
            "bypass_type": "sentinel",
            "reason": reason,
            "char_delta": 0,
            "session_id": session_id(),
            "tool": "Bash",
        })
        sys.stderr.write(
            f"docs-bloat-gate: sentinel burned ({new_count}/{SENTINEL_CAP} "
            "this session). Bash write allowed.\n"
        )
        return 0

    sys.stderr.write(
        "docs-bloat-gate: Bash command writes to a gated doc/rule path "
        f"({', '.join(sorted(set(matches)))}). Use Write/Edit (gate "
        "computes delta + density), or include "
        f"<!-- docs-bloat-gate-override: <reason >={SENTINEL_REASON_MIN} "
        "chars> --> in the command. "
        f"Sentinels used: {used}/{SENTINEL_CAP}.\n"
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
