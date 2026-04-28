"""Tests for `.claude/hooks/docs-bloat-gate.py` v2.

Coverage matches design plan §10:
  - S1 caps (rule/doc/spec); at-cap passes, over-cap blocks
  - S1 bypasses: L2 heading (line + char carve-out), sentinel
  - S2 slop list (each phrase blocks; sentinel/L2 do not bypass)
  - S3 lexical density (block at >100-char delta; <=100-char delta passes)
  - Sentinel cap N=1: 1st use passes + logs, 2nd hard-blocks, cross-session reset
  - Reason validation: <30 chars fails, anti-stopword fails, valid passes
  - Path scope (gated-docs list + .claude/rules/*.md, depth-1 audit-keyword block)
  - Override log JSONL fields
  - Bash write-pattern detection + sentinel cap

Hook is invoked as a subprocess so each test gets a fresh module import
(re-evaluating Path.home() and module-level state).
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[2]
HOOK_PATH = REPO_ROOT / ".claude" / "hooks" / "docs-bloat-gate.py"


def run_hook(
    tmp_path: Path,
    payload: dict,
    *,
    session_id: str = "test-session",
    extra_env: dict[str, str] | None = None,
):
    home = tmp_path / "home"
    home.mkdir(exist_ok=True)
    env = {
        "CLAUDE_PROJECT_DIR": str(tmp_path),
        "HOME": str(home),
        "CLAUDE_SESSION_ID": session_id,
        "PATH": os.environ.get("PATH", ""),
    }
    if extra_env:
        env.update(extra_env)
    proc = subprocess.run(
        [sys.executable, str(HOOK_PATH)],
        input=json.dumps(payload),
        capture_output=True,
        text=True,
        env=env,
        timeout=10,
    )
    return proc.returncode, proc.stderr, proc.stdout


def make_file(tmp_path: Path, rel: str, *, lines: int, line_text: str = "x") -> Path:
    p = tmp_path / rel
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text("\n".join([line_text] * lines) + "\n")
    return p


def edit_payload(file_path: Path, *, old: str, new: str) -> dict:
    return {
        "tool_name": "Edit",
        "tool_input": {
            "file_path": str(file_path),
            "old_string": old,
            "new_string": new,
        },
    }


def write_payload(file_path: Path, content: str) -> dict:
    return {
        "tool_name": "Write",
        "tool_input": {"file_path": str(file_path), "content": content},
    }


# ── S1: tier caps ────────────────────────────────────────────────────────


class TestS1RuleTier:
    """rule tier (current <50 lines) — cap=150."""

    def test_under_cap_passes(self, tmp_path):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        rc, _, _ = run_hook(tmp_path, edit_payload(f, old="x", new="x" + "a" * 100))
        assert rc == 0

    def test_at_cap_passes(self, tmp_path):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        rc, _, _ = run_hook(tmp_path, edit_payload(f, old="", new="a" * 150))
        assert rc == 0

    def test_over_cap_blocks(self, tmp_path):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        rc, err, _ = run_hook(tmp_path, edit_payload(f, old="", new="a" * 151))
        assert rc == 2
        assert "cap=150" in err
        assert "rule tier" in err


class TestS1DocTier:
    """doc tier (50 <= current < 500 lines) — cap=800."""

    def test_under_cap_passes(self, tmp_path):
        f = make_file(tmp_path, "docs/spec.md", lines=200)
        rc, _, _ = run_hook(tmp_path, edit_payload(f, old="", new="a" * 700))
        assert rc == 0

    def test_at_cap_passes(self, tmp_path):
        f = make_file(tmp_path, "docs/spec.md", lines=200)
        rc, _, _ = run_hook(tmp_path, edit_payload(f, old="", new="a" * 800))
        assert rc == 0

    def test_over_cap_blocks(self, tmp_path):
        f = make_file(tmp_path, "docs/spec.md", lines=200)
        rc, err, _ = run_hook(tmp_path, edit_payload(f, old="", new="a" * 801))
        assert rc == 2
        assert "cap=800" in err
        assert "doc tier" in err


class TestS1SpecTier:
    """spec tier (current >= 500 lines) — cap=2000."""

    def test_under_cap_passes(self, tmp_path):
        f = make_file(tmp_path, "docs/spec.md", lines=600)
        rc, _, _ = run_hook(tmp_path, edit_payload(f, old="", new="a" * 1500))
        assert rc == 0

    def test_over_cap_blocks(self, tmp_path):
        f = make_file(tmp_path, "docs/spec.md", lines=600)
        rc, err, _ = run_hook(tmp_path, edit_payload(f, old="", new="a" * 2100))
        assert rc == 2
        assert "cap=2000" in err
        assert "spec tier" in err


# ── S1: L2-heading bypass ───────────────────────────────────────────────


class TestS1L2HeadingBypass:
    def test_new_section_under_line_budget_passes(self, tmp_path):
        """Carve-out is load-bearing: raw char_delta exceeds the rule cap (150)
        but fits inside the 30-line / 1500-char carve-out, so the edit passes
        only because the carve-out subtracts the new-heading body.
        """
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        # 10 body lines × 20 chars + heading = ~225 chars raw → > 150 cap.
        # 11 lines / ~225 chars all fit under 30-line / 1500-char exemption.
        body = "\n".join(["a" * 20] * 10)
        new = f"## New Section\n{body}\n"
        assert len(new) > 150, "test premise: raw delta must exceed rule cap"
        rc, _, _ = run_hook(tmp_path, edit_payload(f, old="", new=new))
        assert rc == 0

    def test_new_section_over_line_budget_blocks(self, tmp_path):
        """Carve-out exempts heading + 29 body lines (30-line cap). Lines
        beyond that count toward S1; with 50 short body lines, the residue
        comfortably exceeds the rule cap (150)."""
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        body = "\n".join(["a" * 20] * 50)
        new = f"## New Section\n{body}\n"
        rc, err, _ = run_hook(tmp_path, edit_payload(f, old="", new=new))
        assert rc == 2
        assert "cap=150" in err

    def test_existing_heading_no_carve_out(self, tmp_path):
        """Heading already in current file → no exemption; large add blocks."""
        f = make_file(tmp_path, ".claude/rules/a.md", lines=5)
        f.write_text("## Existing\nxx\n")
        new = "## Existing\n" + "a" * 200 + "\n"
        rc, err, _ = run_hook(tmp_path, edit_payload(f, old="## Existing\nxx\n", new=new))
        assert rc == 2
        assert "cap=150" in err


# ── S2: slop list ────────────────────────────────────────────────────────


SLOP_WORDS_SAMPLE = [
    "furthermore", "moreover", "additionally", "consequently",
    "nevertheless", "fundamentally", "essentially", "ultimately",
    "comprehensive", "robust", "seamless", "leverage", "utilize",
]

SLOP_PHRASES_SAMPLE = [
    "it's worth noting",
    "it should be noted",
    "having said that",
    "in essence",
    "delve into",
    "tapestry of",
    "in the context of",
]


class TestS2Slop:
    @pytest.mark.parametrize("word", SLOP_WORDS_SAMPLE)
    def test_each_slop_word_blocks(self, tmp_path, word):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        rc, err, _ = run_hook(
            tmp_path, edit_payload(f, old="x", new=f"x {word} value")
        )
        assert rc == 2, f"expected block on slop word {word!r}"
        assert "AI-slop phrases detected" in err
        assert "UNBYPASSABLE" in err
        assert word in err

    @pytest.mark.parametrize("phrase", SLOP_PHRASES_SAMPLE)
    def test_each_slop_phrase_blocks(self, tmp_path, phrase):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        rc, err, _ = run_hook(
            tmp_path, edit_payload(f, old="x", new=f"x {phrase} foo")
        )
        assert rc == 2, f"expected block on slop phrase {phrase!r}"
        assert "AI-slop phrases detected" in err

    def test_pre_existing_slop_in_old_string_not_flagged(self, tmp_path):
        """Slop word present in BOTH old and new (paragraph rewrite) → not net-added."""
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        rc, _, _ = run_hook(
            tmp_path,
            edit_payload(
                f,
                old="ultimately we want X",
                new="ultimately we want Y",
            ),
        )
        assert rc == 0

    def test_added_extra_occurrence_of_existing_slop_blocks(self, tmp_path):
        """Counter-based net detection: old has 1× 'leverage', new has 2×.
        A set-difference would miss this; the count comparison must catch it.
        """
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        rc, err, _ = run_hook(
            tmp_path,
            edit_payload(
                f,
                old="we leverage the system",
                new="we leverage the system and leverage the API",
            ),
        )
        assert rc == 2
        assert "leverage" in err

    def test_sentinel_does_not_bypass_s2(self, tmp_path):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        new = (
            "<!-- docs-bloat-gate-override: legitimate paragraph addition with "
            "specific rationale -->\nfurthermore some content"
        )
        rc, err, _ = run_hook(tmp_path, edit_payload(f, old="", new=new))
        assert rc == 2
        assert "AI-slop phrases" in err

    def test_l2_heading_does_not_bypass_s2(self, tmp_path):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        new = "## New Section\nfurthermore this delves into the realm of bloat"
        rc, err, _ = run_hook(tmp_path, edit_payload(f, old="", new=new))
        assert rc == 2
        assert "AI-slop phrases" in err

    def test_word_boundary_no_substring_match(self, tmp_path):
        """'robustness' should not match 'robust' — \\b enforces word boundary."""
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        rc, _, _ = run_hook(
            tmp_path, edit_payload(f, old="x", new="x robustness measures")
        )
        assert rc == 0


# ── S3: lexical density ─────────────────────────────────────────────────


# All-function-words string → density 0. ~17 chars per repeat.
LOW_DENSITY_UNIT = "the of and to in "


class TestS3Density:
    def test_low_density_over_delta_guard_blocks(self, tmp_path):
        """200+ chars of function words → density<0.45 → block, UNBYPASSABLE."""
        f = make_file(tmp_path, "docs/spec.md", lines=200)
        new = LOW_DENSITY_UNIT * 20  # ~340 chars
        rc, err, _ = run_hook(tmp_path, edit_payload(f, old="", new=new))
        assert rc == 2
        assert "lexical density" in err
        assert "UNBYPASSABLE" in err

    def test_low_density_under_delta_guard_passes(self, tmp_path):
        """50-char low-density add passes (under 100-char guard)."""
        f = make_file(tmp_path, "docs/spec.md", lines=200)
        # 51 chars of pure function words.
        new = "the of and to in the of and to in the of and to in xx"
        assert len(new) > 50 and len(new) <= 100
        rc, _, _ = run_hook(tmp_path, edit_payload(f, old="", new=new))
        assert rc == 0

    def test_high_density_passes(self, tmp_path):
        f = make_file(tmp_path, "docs/spec.md", lines=200)
        new = (
            "regex constants compiled module load token stoplist "
            "cardinality mtime hash session bypass sentinel cap "
            "tier threshold density carve heading exemption ratio."
        ) * 2
        rc, _, _ = run_hook(tmp_path, edit_payload(f, old="", new=new))
        assert rc == 0

    def test_sentinel_does_not_bypass_s3(self, tmp_path):
        f = make_file(tmp_path, "docs/spec.md", lines=200)
        new = (
            "<!-- docs-bloat-gate-override: legitimate paragraph addition with "
            "specific rationale -->\n" + LOW_DENSITY_UNIT * 20
        )
        rc, err, _ = run_hook(tmp_path, edit_payload(f, old="", new=new))
        assert rc == 2
        assert "lexical density" in err

    def test_l2_heading_does_not_bypass_s3(self, tmp_path):
        f = make_file(tmp_path, "docs/spec.md", lines=200)
        new = "## New\n" + LOW_DENSITY_UNIT * 20
        rc, err, _ = run_hook(tmp_path, edit_payload(f, old="", new=new))
        assert rc == 2
        assert "lexical density" in err


# ── Sentinel cap (N=1) ──────────────────────────────────────────────────


VALID_REASON = "specific rationale for this load-bearing prose addition"


class TestSentinelCap:
    def test_first_use_passes_and_logs(self, tmp_path):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        new = (
            "<!-- docs-bloat-gate-override: " + VALID_REASON + " -->\n"
            + "x" * 200
        )
        rc, err, _ = run_hook(
            tmp_path, edit_payload(f, old="", new=new), session_id="sid-cap-1"
        )
        assert rc == 0, err
        assert "sentinel burned" in err
        log = (tmp_path / "docs/analysis/bloat-overrides.log").read_text()
        entry = json.loads(log.strip().splitlines()[-1])
        assert entry["bypass_type"] == "sentinel"
        assert entry["signal"] == "S1"
        assert entry["reason"] == VALID_REASON
        assert entry["char_delta"] >= 200
        assert entry["session_id"]
        assert entry["tool"] == "Edit"

    def test_second_use_same_session_blocks(self, tmp_path):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        new1 = "<!-- docs-bloat-gate-override: " + VALID_REASON + " -->\n" + "x" * 200
        new2 = (
            "<!-- docs-bloat-gate-override: another genuinely "
            "load-bearing prose justification text -->\n" + "y" * 200
        )
        rc1, _, _ = run_hook(tmp_path, edit_payload(f, old="", new=new1), session_id="sid-cap-2")
        assert rc1 == 0
        rc2, err2, _ = run_hook(tmp_path, edit_payload(f, old="", new=new2), session_id="sid-cap-2")
        assert rc2 == 2
        assert "sentinel cap exhausted" in err2

    def test_cross_session_counter_resets(self, tmp_path):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        new = "<!-- docs-bloat-gate-override: " + VALID_REASON + " -->\n" + "x" * 200
        rc1, _, _ = run_hook(tmp_path, edit_payload(f, old="", new=new), session_id="sid-A")
        assert rc1 == 0
        rc2, _, _ = run_hook(tmp_path, edit_payload(f, old="", new=new), session_id="sid-B")
        assert rc2 == 0  # new session_id → new counter file → cap not exhausted

    def test_block_message_shows_counter_before_burn(self, tmp_path):
        """Pre-burn blocks (e.g. no sentinel) should still show counter status."""
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        rc, err, _ = run_hook(tmp_path, edit_payload(f, old="", new="a" * 200))
        assert rc == 2
        assert "Sentinels used this session: 0/1" in err


# ── Sentinel reason validation ─────────────────────────────────────────


class TestSentinelReasonValidation:
    def test_29_char_reason_rejected(self, tmp_path):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        reason = "x" * 29
        new = f"<!-- docs-bloat-gate-override: {reason} -->\n" + "a" * 200
        rc, err, _ = run_hook(
            tmp_path, edit_payload(f, old="", new=new), session_id="sid-reason-1"
        )
        assert rc == 2
        assert "reason too short" in err

    def test_30_char_reason_accepted(self, tmp_path):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        reason = "x" * 30
        new = f"<!-- docs-bloat-gate-override: {reason} -->\n" + "a" * 200
        rc, _, _ = run_hook(
            tmp_path, edit_payload(f, old="", new=new), session_id="sid-reason-2"
        )
        assert rc == 0

    @pytest.mark.parametrize(
        "stop", ["update", "rewrite", "improvement", "cleanup", "fix"]
    )
    def test_anti_stopword_in_reason_rejected(self, tmp_path, stop):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        reason = (
            f"specific {stop} for this load-bearing prose addition justified"
        )
        assert len(reason) >= 30
        new = f"<!-- docs-bloat-gate-override: {reason} -->\n" + "a" * 200
        rc, err, _ = run_hook(
            tmp_path,
            edit_payload(f, old="", new=new),
            session_id=f"sid-stop-{stop}",
        )
        assert rc == 2
        assert "anti-stopword" in err
        assert stop in err


# ── Path scope / regression coverage ────────────────────────────────────


class TestPathScope:
    @pytest.mark.parametrize(
        "rel",
        [
            "docs/spec.md",
            ".claude/rules/foo.md",
            ".claude/rules/sub/bar.md",
        ],
    )
    def test_gated_paths_enforce(self, tmp_path, rel):
        f = make_file(tmp_path, rel, lines=10)
        rc, _, _ = run_hook(tmp_path, edit_payload(f, old="", new="a" * 250))
        assert rc == 2

    @pytest.mark.parametrize(
        "rel",
        ["docs/plans/x.md", "docs/analysis/y.md", "docs/human/z.md", "README.md"],
    )
    def test_non_gated_paths_pass_through(self, tmp_path, rel):
        f = make_file(tmp_path, rel, lines=10)
        # Even huge addition; non-gated paths bypass S1/S2/S3 entirely.
        rc, _, _ = run_hook(
            tmp_path,
            edit_payload(f, old="", new="furthermore " * 200),
        )
        assert rc == 0

    def test_claude_md_gated_only_with_env(self, tmp_path):
        f = make_file(tmp_path, "CLAUDE.md", lines=10)
        # Without env: not gated.
        rc, _, _ = run_hook(tmp_path, edit_payload(f, old="", new="a" * 250))
        assert rc == 0
        # With env: gated.
        rc, _, _ = run_hook(
            tmp_path,
            edit_payload(f, old="", new="a" * 250),
            extra_env={"DOCS_BLOAT_GATE_CLAUDE_MD": "1"},
        )
        assert rc == 2


class TestAuditKeywordBlock:
    def test_new_audit_file_at_docs_root_blocks(self, tmp_path):
        target = tmp_path / "docs" / "weekly-report.md"
        target.parent.mkdir(parents=True, exist_ok=True)
        rc, err, _ = run_hook(
            tmp_path,
            write_payload(target, "# Weekly report\nThis is a research summary."),
        )
        assert rc == 2
        assert "audit/analysis/report/research" in err

    def test_audit_file_in_subdir_passes(self, tmp_path):
        target = tmp_path / "docs" / "analysis" / "weekly-report.md"
        target.parent.mkdir(parents=True, exist_ok=True)
        rc, _, _ = run_hook(
            tmp_path,
            write_payload(target, "# Weekly report\nThis is a research summary."),
        )
        assert rc == 0

    def test_audit_file_with_override_passes(self, tmp_path):
        target = tmp_path / "docs" / "weekly-report.md"
        target.parent.mkdir(parents=True, exist_ok=True)
        body = (
            "<!-- docs-bloat-gate-override: " + VALID_REASON + " -->\n"
            "# Weekly report\nThis is a research summary."
        )
        rc, _, _ = run_hook(tmp_path, write_payload(target, body))
        assert rc == 0


# ── Bash gate ───────────────────────────────────────────────────────────


def bash_payload(cmd: str) -> dict:
    return {"tool_name": "Bash", "tool_input": {"command": cmd}}


class TestBashGate:
    def test_redirect_to_gated_path_blocks(self, tmp_path):
        rc, err, _ = run_hook(tmp_path, bash_payload("echo hi >> docs/spec.md"))
        assert rc == 2
        assert "Bash command writes" in err

    def test_tee_to_gated_path_blocks(self, tmp_path):
        rc, _, _ = run_hook(
            tmp_path, bash_payload("echo hi | tee -a .claude/rules/foo.md")
        )
        assert rc == 2

    def test_sed_inplace_to_gated_path_blocks(self, tmp_path):
        rc, _, _ = run_hook(
            tmp_path, bash_payload("sed -i 's/x/y/' docs/spec.md")
        )
        assert rc == 2

    def test_read_only_command_passes(self, tmp_path):
        rc, _, _ = run_hook(tmp_path, bash_payload("cat docs/spec.md"))
        assert rc == 0

    def test_heredoc_commit_message_passes(self, tmp_path):
        cmd = (
            "git commit -m \"$(cat <<'EOF'\nupdates docs/spec.md\nEOF\n)\""
        )
        rc, _, _ = run_hook(tmp_path, bash_payload(cmd))
        assert rc == 0

    def test_sentinel_with_valid_reason_unblocks(self, tmp_path):
        cmd = (
            "echo hi >> docs/spec.md  "
            "# <!-- docs-bloat-gate-override: " + VALID_REASON + " -->"
        )
        rc, err, _ = run_hook(
            tmp_path, bash_payload(cmd), session_id="sid-bash-1"
        )
        assert rc == 0, err
        assert "sentinel burned" in err

    def test_sentinel_cap_applies_to_bash(self, tmp_path):
        cmd1 = (
            "echo a >> docs/spec.md  "
            "# <!-- docs-bloat-gate-override: " + VALID_REASON + " -->"
        )
        cmd2 = (
            "echo b >> .claude/rules/foo.md  "
            "# <!-- docs-bloat-gate-override: another genuinely load-bearing "
            "rationale text here -->"
        )
        rc1, _, _ = run_hook(tmp_path, bash_payload(cmd1), session_id="sid-bash-2")
        assert rc1 == 0
        rc2, err2, _ = run_hook(tmp_path, bash_payload(cmd2), session_id="sid-bash-2")
        assert rc2 == 2
        assert "sentinel cap exhausted" in err2


# ── Override log format ─────────────────────────────────────────────────


class TestOverrideLog:
    def test_l2_bypass_writes_log_entry(self, tmp_path):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        # Must exceed cap=150 BEFORE carve-out, fit AFTER carve-out.
        # 25 lines of 10 chars = 275 chars (>150 raw, all exempt).
        body = "\n".join(["a" * 10] * 25)
        new = f"## New Section\n{body}\n"
        rc, _, _ = run_hook(
            tmp_path, edit_payload(f, old="", new=new), session_id="sid-log-l2"
        )
        assert rc == 0
        log = (tmp_path / "docs/analysis/bloat-overrides.log").read_text()
        entry = json.loads(log.strip().splitlines()[-1])
        assert entry["bypass_type"] == "l2_heading"
        assert entry["signal"] == "S1"
        assert "carved_delta" in entry
        assert entry["headings"] == ["New Section"]

    def test_log_is_appended_jsonl(self, tmp_path):
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        new1 = f"<!-- docs-bloat-gate-override: {VALID_REASON} -->\n" + "x" * 200
        run_hook(tmp_path, edit_payload(f, old="", new=new1), session_id="sid-log-A")
        run_hook(tmp_path, edit_payload(f, old="", new=new1), session_id="sid-log-B")
        lines = (tmp_path / "docs/analysis/bloat-overrides.log").read_text().splitlines()
        assert len(lines) == 2
        for ln in lines:
            entry = json.loads(ln)
            assert entry["bypass_type"] == "sentinel"


# ── v1 regression: pointer bypass removed ──────────────────────────────


class TestV1RegressionPointerBypass:
    def test_arrow_pointer_no_longer_bypasses_s1(self, tmp_path):
        """v1 allowed '→ see X' as a bypass. v2 design §5 removes it."""
        f = make_file(tmp_path, ".claude/rules/a.md", lines=10)
        new = "→ see docs/spec.md §3 for details — " + "a" * 200
        rc, err, _ = run_hook(tmp_path, edit_payload(f, old="", new=new))
        assert rc == 2
        assert "cap=150" in err
