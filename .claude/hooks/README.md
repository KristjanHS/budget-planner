# .claude/hooks

Hooks adapted from [pedrohcgs/claude-code-my-workflow](https://github.com/pedrohcgs/claude-code-my-workflow) (2026-04-15) plus project-specific gates.

## Active hooks (7)

| File | Event | Purpose |
|---|---|---|
| `pre-compact.py` | `PreCompact` | Before compaction: saves active plan + recent decisions to `~/.claude/sessions/<hash>/pre-compact-state.json`; appends a timestamped note to the newest `docs/session_logs/*.md` if any exists |
| `post-compact-restore.py` | `SessionStart` (matcher: `compact\|resume`) | After compaction/resume: reads and deletes the saved state; echoes plan, log, and recovery actions |
| `context-monitor.py` | `PostToolUse` (matcher: `Bash\|Task`) | Progressive context-usage warnings at 40/55/65% (capture reusable insights), 80% (info), 90% (critical). Heuristic: tool-call count ÷ 150 |
| `verify-reminder.py` | `PostToolUse` (matcher: `Write\|Edit`) | On `.py` edits outside skip-dirs, nudges to run `pytest tests/ -q`. Throttled to 60s per file |
| `docs-bloat-gate.py` | `PreToolUse` (matcher: `Write\|Edit\|Bash`) | Blocks net-positive Write/Edit to `docs/spec.md`, `docs/runbook.md`, `docs/refactor-incidents.md`, `.claude/rules/*.md`, plus Bash `>`/`>>`/`tee`/`sed -i` writes targeting those paths. Pass conditions: net delta ≤ 0, `→ see X` pointer, brand-new `## ` heading (lines under it exempt), or override sentinel. Design: `docs/plans/2026-04-25-doc-bloat-root-causes-brainstorm.md` |
| `docs-only-plan-gate.py` | `PreToolUse` Bash + `PostToolUse` Bash | **Pre:** blocks `git mv .*docs/plans/.*archived/` unless a pytest-pass stamp is fresh (≤30 min) or override sentinel present. **Post:** stamps `~/.claude/sessions/<hash>/pytest-pass-stamp` when a `pytest tests/...` invocation exits 0. Source: `docs/plans/2026-04-25-prose-rules-to-hooks-audit.md` (W1.2) |
| `external-process-test-gate.py` | `PreToolUse` (matcher: `Write\|Edit`) | Blocks Write/Edit to `tests/**/*.py` (excluding `tests/manual/`, `tests/conftest.py`, `tests/__init__.py`) when content matches docker / kubectl / playwright / selenium imports or subprocess calls, or `requests.<get\|post\|...>(`. Source: `docs/plans/2026-04-25-prose-rules-to-hooks-audit.md` (W1.4) |

All hooks fail open (exit 0 on any exception) so a hook bug never blocks Claude.

## Troubleshooting `docs-bloat-gate.py`

**Symptom: legitimate edit blocked.**
Read the stderr message — it reports the net-delta and the four bypass options. Pick whichever fits:
- Add a `→ see <path> § <section>` pointer in the added text (regex: `→\s*see\s+\S+` or `->\s*see\s+\S+`).
- Pair the addition with a removal so net delta ≤ 0.
- Wrap the new content under a brand-new `## ` heading (an L2 heading not already in the file). Lines under that heading are exempt; lines outside still count.
- Add `<!-- docs-bloat-gate-override: <reason ≥10 chars> -->` in the added content. The reason is appended to `docs/analysis/bloat-overrides.log` (gitignored).

**Symptom: Bash command blocked.**
Stderr lists the matched gated path(s). The matcher only fires on true write-actions: `>`/`>>` redirection, `tee`, or `sed -i`. Read-only commands (`cat`, `grep`, `git log`) and commit messages mentioning gated paths must pass — if they don't, that's a bug, file it.

**Test the hook in isolation:**
```sh
echo '{"tool_name":"Edit","tool_input":{"file_path":"'$PWD'/docs/spec.md","old_string":"x","new_string":"x\nA\nB"}}' \
  | python3 .claude/hooks/docs-bloat-gate.py
echo "exit code: $?"   # 2 = blocked, 0 = allowed; stderr has the reason
```

**Inspect override log:** `cat docs/analysis/bloat-overrides.log` — one line per override, format `<ISO-ts>  <path>  <±N> lines  reason="…"  tool=<Write|Edit|Bash>`. ≥3 entries since last `/retro` is a signal to revisit gate calibration.

**Disable temporarily:**
- For one Edit: add the override sentinel.
- For all of CLAUDE.md (opt-in only): unset `DOCS_BLOAT_GATE_CLAUDE_MD` (default off).
- For the whole hook: comment out the `PreToolUse` block in `.claude/settings.json` (and tell the user; restart for it to take effect).

**Hook isn't firing on a path you expected:**
Path scope is exact-list (`GATED_DOCS`) plus `.claude/rules/*.md` glob. Subdirs of `docs/` (`docs/plans/`, `docs/analysis/`, `docs/human/`, `docs/research/`) are intentionally out of scope per design §1.

**Latency check:** `time for i in $(seq 1 10); do echo '{"tool_name":"Edit","tool_input":{"file_path":"'$PWD'/docs/spec.md","old_string":"a","new_string":"a"}}' | python3 .claude/hooks/docs-bloat-gate.py >/dev/null 2>&1; done` — target <50ms per invocation.

## Troubleshooting `docs-only-plan-gate.py`

**Symptom: `git mv` to `archived/` blocked.**
Stderr says the pytest-pass stamp is absent or stale. The fastest unblock:

1. Run `pytest tests/ -q`. On a clean pass, the PostToolUse branch of this same hook writes `~/.claude/sessions/<hash>/pytest-pass-stamp`.
2. Retry the `git mv` within 30 minutes.

If pytest is genuinely irrelevant to the change (e.g. moving a plan whose archive lost its content even before this commit), include `<!-- docs-only-plan-gate-override: <reason ≥10 chars> -->` in the bash command. The reason is appended to `docs/analysis/docs-only-plan-overrides.log` (gitignored).

**Symptom: stamp not refreshing after green pytest.**
The PostToolUse branch fires only when:
- `tool_name == "Bash"` AND
- the command matches `\bpytest\b[^|;&]*\btests/` (i.e. `pytest tests/...` — bare `pytest` with no `tests/` does NOT stamp), AND
- `tool_response.exit_code == 0` (or the stdout-fallback heuristic sees `N passed` with no `failed`/`error`).

Verify: `ls -l ~/.claude/sessions/$(python3 -c "import hashlib; print(hashlib.md5(b'$PWD'.encode()).hexdigest()[:8])")/pytest-pass-stamp`.

**Test the hook in isolation:**
```sh
# block path (no stamp)
echo '{"hook_event_name":"PreToolUse","tool_name":"Bash","tool_input":{"command":"git mv docs/plans/X.md docs/plans/archived/X.md"}}' \
  | python3 .claude/hooks/docs-only-plan-gate.py
# stamp via PostToolUse
echo '{"hook_event_name":"PostToolUse","tool_name":"Bash","tool_input":{"command":"pytest tests/ -q"},"tool_response":{"exit_code":0}}' \
  | python3 .claude/hooks/docs-only-plan-gate.py
```

**Inspect override log:** `cat docs/analysis/docs-only-plan-overrides.log`. ≥3 entries since last `/retro` is a signal to revisit gate calibration.

## Troubleshooting `external-process-test-gate.py`

**Symptom: legitimate test edit blocked.**
Stderr lists the matched patterns (e.g. `subprocess to external CLI`, `requests HTTP call`). Pick whichever fits:
- **Move the test under `tests/manual/`** (the carve-out — see `tests/manual/README.md`). Run with `pytest tests/manual/ -q` on demand.
- **Add the override sentinel** in the new content: `<!-- external-process-test-gate-override: <reason ≥10 chars> -->`. The reason is appended to `docs/analysis/external-process-test-overrides.log` (gitignored).

**Symptom: hook fires on legitimate `subprocess.run([...recalc.py...])`.**
The hook's subprocess pattern only matches first-arg strings starting with `docker`/`kubectl`/`playwright`/`selenium`. If you're seeing a false-positive on `recalc.py` or another project tool, file it — the regex is intentionally narrow to avoid that case (see hook docstring "Why narrower than the W1.4 audit's `subprocess.run` blanket").

**Test the hook in isolation:**
```sh
# block path
echo '{"hook_event_name":"PreToolUse","tool_name":"Write","tool_input":{"file_path":"'$PWD'/tests/test_x.py","content":"import requests\nrequests.get(\"http://x\")"}}' \
  | python3 .claude/hooks/external-process-test-gate.py
# allow path (under tests/manual/)
echo '{"hook_event_name":"PreToolUse","tool_name":"Write","tool_input":{"file_path":"'$PWD'/tests/manual/test_x.py","content":"import requests\nrequests.get(\"http://x\")"}}' \
  | python3 .claude/hooks/external-process-test-gate.py
```

**Path scope (exempt):** `tests/__init__.py`, `tests/conftest.py`, anything under `tests/manual/**`. Files outside `tests/` are silently ignored.

**Inspect override log:** `cat docs/analysis/external-process-test-overrides.log`.

## Regression-test scaffold

`.claude/hooks/_manual_test_stage2.sh` runs 23 block/allow cases plus latency checks for the two Stage-2 hooks. Run from repo root: `bash .claude/hooks/_manual_test_stage2.sh`. Exit code = number of failures.

## Dropped (2, intentional)

- `notify.sh` — desktop notifier. WSL2 has no `notify-send` and no shared DBus session, so it would no-op anyway.
- `log-reminder.py` — Stop-hook that **blocks** Claude from stopping until a `docs/session_logs/*.md` is written. Too opinionated for this project.

## Key adaptations from upstream

- **Paths**: `quality_reports/{plans,session_logs}/` → `docs/{plans,session_logs}/` (plus `plansDirectory` in `settings.json`).
- **verify-reminder extensions**: `.tex/.qmd/.R` (LaTeX/Quarto/R) → `.py` with `pytest tests/ -q` as the verify action.
- **verify-reminder skip-dir matching**: upstream used absolute-path substring (`"/.claude/" in path`). This project is edited inside `.claude/worktrees/<name>/`, so the substring matched every file and every reminder was silently suppressed. Rewrote to match `SKIP_DIRS` against project-relative `Path.parts` using `CLAUDE_PROJECT_DIR`.
- **context-monitor**: removed the `/learn` skill reference (not installed here); generic "capture to memory/CLAUDE.md/docs/" nudge instead.

## Wired in

- `.claude/settings.json` — `hooks` block for the 4 events above, plus `"plansDirectory": "docs/plans"`.
- `docs/session_logs/` — empty dir (`.gitkeep`). Hooks handle it being empty gracefully; create log files there to activate decision-extraction in `pre-compact`.

## Cache location

All hooks write transient state to `~/.claude/sessions/<md5(CLAUDE_PROJECT_DIR)[:8]>/`:
`pre-compact-state.json`, `context-monitor-cache.json`, `verify-reminder-cache.json`. Safe to delete.
