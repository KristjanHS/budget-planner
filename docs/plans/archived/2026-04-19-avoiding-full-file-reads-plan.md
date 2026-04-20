# Plan: Avoiding Full-File Reads

**Date**: 2026-04-19
**Status**: Draft — awaiting approval
**Scope**: Claude Code tool-usage habits for this project and globally. Complements (does not replace) the prior `2026-04-12-context-optimizations.md` plan, which addressed auto-loaded docs.

## Context

A full Read of a 1000-line file is ~8K tokens; a 1500-line file ~12K (Read adds ~70% overhead via line-number formatting). Several files in this project exceed that:

| File | Lines | Full-read cost (approx.) |
|------|-------|--------------------------|
| `budget_planner/reader.py` | 1494 | ~12K tokens |
| `budget_planner/timelog_plan.py` | 1292 | ~10K tokens |
| `tests/test_timelog_plan.py` | 1218 | ~10K tokens |
| `budget_planner/capacity.py` | 1066 | ~9K tokens |
| `budget_planner/config_sheet.py` | 1031 | ~8K tokens |
| `budget_planner/workbook.py` | 935 | ~7K tokens |
| `docs/plans/archived/2026-04-12-rebuild-shape-drift-plan.md` | 1225 | ~10K tokens |
| `docs/plans/archived/2026-04-13-hours-timelog-v1-plan.md` | 861 | ~7K tokens |

Two naïve Reads of `reader.py` + `timelog_plan.py` is ~22K tokens — ~11% of a 200K window in a single turn, and a silent autocompact trigger at opus/effort=high. The current `reading-large-files.md` rule only gates on `.lock/.csv/.min.*`, so `.py` and `.md` files are uncovered.

**Research sources** (Anthropic docs + Claude Code guide agent, 2026-04-19):
- Best Practices for Claude Code — managing context, subagent delegation
- How Claude Code works — subagent isolation
- Tools reference — Read offset/limit, Grep/Glob mechanics
- Global CLAUDE.md Efficiency section — existing 800-line threshold

## Best Practices Distilled

1. **Grep-first for known symbols.** If you know the function/class/string you want, `Grep` returns just the matching lines (±context). A full Read is wasteful.
2. **`offset`/`limit` for structural sampling.** Reading "the top of `reader.py` to orient" → `Read(offset=0, limit=80)`, not a full read.
3. **Targeted range reads.** Once Grep tells you a symbol is at line 874, `Read(offset=860, limit=80)` gives the surrounding block.
4. **Subagent dispatch for ≥3-round exploration.** `Agent(subagent_type="Explore")` keeps intermediate grep/read output out of the parent context — only the summary returns.
5. **Don't re-Read after Edit.** Edit validates `old_string` matched; success means the change landed. Skip verification reads.
6. **Thresholds by line count** (derived from docs + prior practice):
   - <300 lines: full Read is fine.
   - 300–800 lines: prefer Grep for specific symbols; full Read acceptable if whole-file structure matters.
   - ≥800 lines: `offset`/`limit` or Grep. Full Read needs justification.
7. **Declare intent before reading.** One-line "reading X to find Y" forces the Grep-vs-Read decision before spending tokens.

## Proposed Changes

### Change 1 — Extend `~/.claude/rules/reading-large-files.md` to cover `.py` / `.md`

**Why**: Current rule only triggers on lock/csv/minified. It should also gate on large source / doc files in projects where they exist, with a clear threshold.

**Edit** `~/.claude/rules/reading-large-files.md`:
- Add `"**/*.py"` and `"**/*.md"` to `paths:` frontmatter.
- Update body to distinguish the two categories:
  - Unconditional (lock/csv/min): never full-read without offset/limit.
  - Conditional (py/md): apply the ≥800-line threshold. Under 800, normal Read is fine.
- Add the "declare intent before reading" one-liner.

**Tradeoff**: This rule will load on every `.py` and `.md` file touch (~15 lines, ~50 tokens). The prior optimization plan explicitly narrowed path-gated rules; adding back a broad gate is a reversal. Justification: the rule's cost (~50 tokens / session) is ~100× smaller than one avoided full-read of `reader.py` (~12K tokens).

**Alternative considered**: Add the guidance directly to global CLAUDE.md Efficiency section (already loaded every session, no new gate). Rejected because CLAUDE.md is already token-dense and the guidance is operational (applies when *reading* a file), which is what path-gates exist for.

### Change 2 — Add large-file directory to project CLAUDE.md

**Why**: The file table above is project-specific truth. Loading it every session primes Claude to reach for Grep when a file is known-large.

**Edit** `~/projects/budget-planner/CLAUDE.md`:
Add a subsection under "Critical Rules" or a new section just after:

```markdown
## Large Files (prefer Grep / `offset`+`limit`)

Files ≥800 lines — a full Read is ~8K+ tokens. Use Grep for specific symbols; Read with `offset`/`limit` for ranges.

| File | Lines |
|------|-------|
| `budget_planner/reader.py` | 1494 |
| `budget_planner/timelog_plan.py` | 1292 |
| `tests/test_timelog_plan.py` | 1218 |
| `budget_planner/capacity.py` | 1066 |
| `budget_planner/config_sheet.py` | 1031 |
| `budget_planner/workbook.py` | 935 |
| `tests/test_config_sheet.py` | 838 |
| `tests/test_reader_new.py` | 824 |
| `tests/test_trends.py` | 817 |
| `tests/test_rebuild_e2e.py` | 804 |

Line counts drift; re-check with `wc -l` if a number feels wrong. Archived plans in `docs/plans/archived/` are not listed — treat all of them as large by default.
```

**Cost**: ~200 tokens added to always-loaded project CLAUDE.md. Acceptable against the ~12K saved per avoided full-read.

### Change 3 — Codify the "declare intent" habit in global CLAUDE.md Efficiency

**Why**: The habit forces a Grep-vs-Read decision *before* the expensive call. Currently implicit.

**Edit** `~/.claude/CLAUDE.md` Efficiency section, add one bullet after the existing "For files >800 lines…" line:

```markdown
- **State the goal before the read**: one sentence — "reading X to find Y" — before Read/Grep on a large file. If the goal is "find the FOO function", Grep beats Read. If it's "understand the module's top-level structure", `Read(offset=0, limit=80)` beats a full read. The declaration is the check.
```

**Cost**: ~70 tokens in always-loaded global CLAUDE.md.

### Change 4 — No change to subagent guidance (already strong)

Global CLAUDE.md already covers Explore/Plan subagents for ≥3-round exploration. No edit needed — verifying the guidance is followed is a behavior check, not a doc change.

### Change 5 — No change to Read-after-Edit guidance (already strong)

Global CLAUDE.md already says "Don't re-read files already in context. Especially after Edit". No edit needed.

## What's Explicitly NOT in Scope

- **Refactoring large files to be smaller.** That's a code-architecture concern, not a tool-usage concern. `reader.py` at 1494 lines may or may not warrant splitting — out of scope for this plan.
- **New MCP tools for file reading.** The existing Read/Grep/Glob cover every case in the distilled practices.
- **Automated token-budget tracking.** Research turned up no documented tooling; would be bespoke work.

## Verification (after applying)

1. Start a fresh session, touch `budget_planner/reader.py` → confirm `reading-large-files.md` loads (check system-reminder).
2. Ask a question like "how does `_read_plan_sheet` parse the WP rows?" — confirm Claude reaches for Grep, not a full Read of `reader.py`.
3. Start a session and ask about `docs/plans/archived/2026-04-12-rebuild-shape-drift-plan.md` — confirm Claude uses `offset`/`limit` rather than full read.
4. Inspect an Edit-followed-by-Read sequence in the transcript — confirm no redundant Reads after successful Edits.

## Files to Change

- `~/.claude/rules/reading-large-files.md` — extend frontmatter + body (Change 1)
- `~/projects/budget-planner/CLAUDE.md` — add "Large Files" section (Change 2)
- `~/.claude/CLAUDE.md` — add one bullet to Efficiency (Change 3)

## Combined Impact

| # | Change | Tokens cost/session | Tokens saved per avoided full-read |
|---|--------|---------------------|-------------------------------------|
| 1 | Extend rule to py/md | ~50 (on py/md touch) | 7–12K per avoidance |
| 2 | Project CLAUDE.md large-file table | ~200 (always) | 7–12K per avoidance |
| 3 | Global CLAUDE.md "declare intent" | ~70 (always) | 7–12K per avoidance |

Net break-even after ~1 avoided full-read per session. In practice, a typical editing session touches 2–4 large files — so 2–4 avoidances = 14–48K tokens saved, vs ~320 tokens cost.

## Open Questions

- Change 2's table will drift as files grow/shrink. Should we periodically re-check with `wc -l` (manual) or add a script to `scripts/` that regenerates the table? Defer decision; manual re-check is low-effort.
- The global CLAUDE.md dotfiles-propagation dance (commit in devcontainer → pull on host → re-stow) applies if Change 3 is made from a devcontainer. This session is on WSL host (`cwd = ~/projects/...`), so direct edit works — but flag for future sessions.
