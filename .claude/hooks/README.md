# .claude/hooks

Hooks adapted from [pedrohcgs/claude-code-my-workflow](https://github.com/pedrohcgs/claude-code-my-workflow) (2026-04-15) plus project-specific gates.

## Active hooks (6)

| File | Event | Purpose |
|---|---|---|
| `pre-compact.py` | `PreCompact` | Before compaction: saves active plan + recent decisions to `~/.claude/sessions/<hash>/pre-compact-state.json`; appends a timestamped note to the newest `docs/session_logs/*.md` if any exists |
| `post-compact-restore.py` | `SessionStart` (matcher: `compact\|resume`) | After compaction/resume: reads and deletes the saved state; echoes plan, log, and recovery actions |
| `context-monitor.py` | `PostToolUse` (matcher: `Bash\|Task`) | Progressive context-usage warnings at 40/55/65% (capture reusable insights), 80% (info), 90% (critical). Heuristic: tool-call count ÷ 150 |
| `verify-reminder.py` | `PostToolUse` (matcher: `Write\|Edit`) | On `.py` edits outside skip-dirs, nudges to run `pytest tests/ -q`. Throttled to 60s per file |
| `docs-bloat-gate.py` | `PreToolUse` (matcher: `Write\|Edit\|Bash`) | Three signals on Write/Edit to `docs/spec.md`, `docs/runbook.md`, `docs/refactor-incidents.md`, `.claude/rules/*.md` (any blocks): **S1** char-delta vs tier cap (rule<50 lines=150 chars, doc=800, spec=2000) — bypassable; **S2** ~30-pattern slop stoplist (transition adverbs, hedges, AI tells) — UNBYPASSABLE; **S3** lexical density <0.45 on >100-char addition — UNBYPASSABLE. S1 bypasses: brand-new `## ` heading (≤30 lines / 1500 chars exempt on the introducing edit) or override sentinel (hard cap **1/session**, reason ≥30 chars, no anti-stopwords). Bash gate blocks `>`/`>>`/`tee`/`sed -i` to gated paths; same sentinel cap applies. Design: `docs/plans/archived/2026-04-26-docs-bloat-gate-redesign.md` |
| `docs-only-plan-gate.py` | `PreToolUse` Bash + `PostToolUse` Bash | **Pre:** blocks `git mv .*docs/plans/.*archived/` unless a pytest-pass stamp is fresh (≤30 min) or override sentinel present. **Post:** stamps `~/.claude/sessions/<hash>/pytest-pass-stamp` when a `pytest tests/...` invocation exits 0. Source: `docs/plans/archived/2026-04-25-prose-rules-to-hooks-audit.md` (W1.2) |

`external-process-test-gate.py` was promoted to global (`~/.claude/hooks/`, registered in global `settings.json`) on 2026-04-25 since its source rule (`~/.claude/rules/testing.md`) is also global. Override log moved to `~/.claude/logs/external-process-test-overrides.log`. Behavior unchanged for this project.

All hooks fail open (exit 0 on any exception) so a hook bug never blocks Claude.

## Troubleshooting `docs-bloat-gate.py`

**Symptom: edit blocked. Which signal fired?** Read the stderr prefix:

- **`AI-slop phrases detected: …`** (S2, UNBYPASSABLE). Remove the listed phrase(s) from the added text. Sentinel and L2-heading bypasses do NOT apply. Stoplist covers transition adverbs (`furthermore`, `moreover`, …), hedges (`it's worth noting`, `having said that`, `in essence`, …), and AI tells (`comprehensive`, `robust`, `leverage`, `delve into`, …). The check subtracts hits already in `old_string`, so paragraph rewrites that preserve the phrase are not flagged.
- **`lexical density <D> below 0.45 …`** (S3, UNBYPASSABLE). Function-word ratio is too high — tighten by removing hedges, articles, and transition phrases. Only fires on additions >100 chars.
- **`char_delta=N exceeds cap=C (<tier> tier, …)`** (S1, bypassable). Three options:
  1. Revise the edit to fit the tier cap (rule<50 lines=150, doc=800, spec=2000 chars).
  2. Wrap new content under a brand-new `## ` heading (not already in the file). The introducing edit gets a 30-line / 1500-char carve-out, whichever hits first; lines past that count toward S1 normally.
  3. Add `<!-- docs-bloat-gate-override: <reason ≥30 chars> -->`. **Hard cap: 1 sentinel per session.** The reason must contain no anti-stopwords (`update`, `rewrite`, `improvement`, `cleanup`, `fix`).

**Symptom: sentinel cap exhausted.** Stderr shows `K/1 used`. All gated-doc edits hard-block until the next session — by design (per design §6). Either land the change differently or end the session.

**Symptom: Bash command blocked.** Stderr lists the matched gated path(s). The matcher fires only on true write-actions: `>`/`>>`, `tee`, or `sed -i`. Read-only commands (`cat`, `grep`, `git log`) and heredoc commit messages mentioning gated paths must pass — if they don't, that's a bug, file it. Bash blocks honor the same 1/session sentinel cap.

**Test the hook in isolation:**
```sh
echo '{"tool_name":"Edit","tool_input":{"file_path":"'$PWD'/docs/spec.md","old_string":"x","new_string":"x\nA\nB"}}' \
  | python3 .claude/hooks/docs-bloat-gate.py
echo "exit code: $?"   # 2 = blocked, 0 = allowed; stderr has the reason
```

**Inspect override log:** `jq -Rc 'try fromjson' docs/analysis/bloat-overrides.log` — JSONL, one entry per bypass use. (`-R 'try fromjson'` skips any pre-v2 plain-text lines that may co-exist in the file.) Fields: `ts`, `file`, `signal` (always `S1` in v2), `bypass_type` (`l2_heading` or `sentinel`), `char_delta`, optional `reason`/`carved_delta`/`headings`, `session_id`, `tool`. ≥3 entries since last `/retro` is a signal to revisit calibration.

**Sentinel counter file:** `~/.claude/state/docs-bloat-gate-sentinels-<session_id>.txt` (single integer). Session ID resolves from `CLAUDE_SESSION_ID` env → hash of `CLAUDE_TRANSCRIPT_PATH` → hourly-rolling token. Counter auto-resets when the session ID changes. Safe to delete to reset within a session for debugging.

**Disable temporarily:**
- For one S1 edit: use the override sentinel (subject to the 1/session cap).
- For S2/S3: not bypassable. Revise the prose.
- For CLAUDE.md gating (opt-in only): unset `DOCS_BLOAT_GATE_CLAUDE_MD` (default off).
- For the whole hook: comment out the `PreToolUse` block in `.claude/settings.json` (tell the user; restart for it to take effect).

**Hook isn't firing on a path you expected:**
Path scope is exact-list (`GATED_DOCS`) plus any `.md` under `.claude/rules/` (recursive — subdirs are gated). Subdirs of `docs/` (`docs/plans/`, `docs/analysis/`, `docs/human/`, `docs/research/`) are intentionally out of scope. The audit-keyword block is a separate check that fires ONLY on a brand-new file written directly under `docs/` (depth=1, e.g. `docs/weekly-report.md`) when its content matches `audit|analysis|report|research`; plans under `docs/plans/`, analyses under `docs/analysis/`, and any other subdir are unaffected.

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
