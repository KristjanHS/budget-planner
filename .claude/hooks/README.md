# .claude/hooks

Hooks adapted from [pedrohcgs/claude-code-my-workflow](https://github.com/pedrohcgs/claude-code-my-workflow) (2026-04-15).

## Adopted (4)

| File | Event | Purpose |
|---|---|---|
| `pre-compact.py` | `PreCompact` | Before compaction: saves active plan + recent decisions to `~/.claude/sessions/<hash>/pre-compact-state.json`; appends a timestamped note to the newest `docs/session_logs/*.md` if any exists |
| `post-compact-restore.py` | `SessionStart` (matcher: `compact\|resume`) | After compaction/resume: reads and deletes the saved state; echoes plan, log, and recovery actions |
| `context-monitor.py` | `PostToolUse` (matcher: `Bash\|Task`) | Progressive context-usage warnings at 40/55/65% (capture reusable insights), 80% (info), 90% (critical). Heuristic: tool-call count ÷ 150 |
| `verify-reminder.py` | `PostToolUse` (matcher: `Write\|Edit`) | On `.py` edits outside skip-dirs, nudges to run `pytest tests/ -q`. Throttled to 60s per file |

All hooks fail open (exit 0 on any exception) so a hook bug never blocks Claude.

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
