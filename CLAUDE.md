# CLAUDE.md — Project Instructions for Claude Code

## What This Is

An openpyxl-based Excel workbook generator for EU project FTE budget planning. The user will iterate on this with you.

<!-- L1=this file (always), L2=docs/ (on demand via trigger table), L3=.claude/rules/ (path-gated) -->

## Reference Index (encounter a problem? check here)

| Trigger | Read | Key content |
|---------|------|-------------|
| Programme cost model, funding rates, 25% indirect | `docs/grant-domain.md` | Cost categories, PM→EUR chain, reallocation rules |
| Sheet builders, stateless pattern | `docs/internals.md` § Sheet Builders | Builder contract, dependency chain; don't add state |
| Reader breaks, rebuild loses data | `docs/internals.md` § Reader and Rebuild | Formula chain resolution, fragile parsing |
| Column offsets, month positioning, styles, row/column shifts, new sheet/column | `budget_planner/layout.py` + `budget_planner/styles.py` | `layout.py` is the single source for row/column positions, validation-column offsets, freeze panes, alert thresholds. Tab-color + styling conventions: `docs/internals.md` § Styling conventions. |
| Cross-sheet invariant, trust boundary, derived-metric rule | `docs/design-invariants.md` | Two-cadence actuals, single-month-cell rule; update before merge |
| C4 architecture, workflows, formula graph | `docs/architecture.md` | Mermaid diagrams, build order, cross-sheet dependencies |
| Cloud/container env, LO missing, confidential-workbook debugging | `docs/ops-quickref.md` | LO install, failure-mode table, `make_reproducer` pointer. Full runbooks in `docs/human/` (intentionally `.claudeignore`'d — user-facing runbooks, not auto-loaded). |
| Code review, /impag, worktrees, MCP tools for workbook inspection | `docs/dev-workflow-review.md` (§ MCP Tools for openpyxl-read / mcp-excel) | Grep-cascade checks, planning protocol, worktree sync, MCP usage |
| Add WP, add team member, add actuals, compare scenarios | `docs/common-tasks.md` | Step-by-step SOPs for common data tasks |
| Timelog Plan, solver, 1.0-FTE logging allocation, WP fill order, merged grid, resolve toggle, per-year EUR cols, monthly total, yearly total, yearly budget, vacation month, Pool reduction | `budget_planner/timelog_plan.py` + `budget_planner/timelog_solver.py` + `docs/plans/archived/2026-04-19-timelog-plan-enhancements-design.md` (Phase 3–4 totals + vacation) + `docs/plans/archived/2026-04-19-timelog-plan-v2-design.md` (v2) + `docs/plans/archived/2026-04-18-timelog-plan-design.md` (v1) | Separate allocation surface from Plan; two-pass greedy solver; **merged editable grid** — `timelog_resolve_flag` (TRUE = re-solve + wipe edits; FALSE = preserve) controls rebuild; vacation cols yellow-locked; Monthly total € row under grid; Yearly total € + Yearly budget € rows under PER-WP SUMMARY |
| Rebuild drops rows, stale Config refs | `docs/plans/archived/2026-04-12-rebuild-shape-drift-design.md` | Readers use named ranges; tolerate internal blanks |
| WP schedule single range, migration from legacy workbook, multi-segment error | `docs/plans/archived/2026-04-19-wp-gap-removal-design.md` + `docs/design-invariants.md` § WP schedule — single active range | Start/End/Type merged into WP BUDGET cols H/I/J; reader hard-fails "Legacy Config layout detected" / "Duplicate WP id" with migration recipe |
| Run commands, generate workbook, compare scenarios, dev setup | `docs/commands.md` | Init, rebuild, snapshot/restore, pytest, recalc |
| "record this", "remember this", organizing docs | `~/.claude/references/recording-principles.md` | L1/L2/L3 criteria, what belongs where |

## Critical Rules

1. **Always use Excel formulas, never hardcode calculated values in Python.** The spreadsheet must remain dynamic.
2. **After every generation: `python scripts/recalc.py output/budget_plan.xlsx` then `pytest tests/ -v`.** Skipping recalc makes `data_only=True` reads return None — tests pass silently on broken output. Verify all pass before presenting to user.
3. **Font color convention is sacred**: Blue = input, Black = formula, Green = cross-sheet ref. Never mix these.
4. **Inactive month cells must be greyed out and empty.** Auditors flag time logged outside WP active periods.
5. **Named ranges `project_start` and `fte_rate`** must always exist and point to Config!$B$2 and Config!$B$4.
6. **New architectural invariants** (cross-sheet constraint, trust boundary, derived-metric rule) must be recorded in `docs/design-invariants.md` before merge. Rationale + scope carve-out in `.claude/rules/dev-workflow.md`.

## Large Files (prefer Grep / `offset`+`limit`)

Files ≥800 lines — a full Read is ~8K+ tokens. Use Grep for specific symbols; Read with `offset`/`limit` for ranges. Find them on demand: `wc -l budget_planner/*.py tests/*.py | sort -rn | head`. Archived plans in `docs/plans/archived/` are not listed — treat all of them as large by default.

## Rules Index (path-gated — auto-loaded when touching matching files)

| Rule file | Triggers on | Key content |
|-----------|-------------|-------------|
| `.claude/rules/dev-workflow.md` | `**/*.py` | Build/test, linting, formatting, invariants (core cross-cutting) |
| `.claude/rules/testing-project.md` | `tests/**`, `**/test_*.py` | Project testing rules (no hardcoded seeds, round-trip, docs-only stages) |
| `.claude/rules/plan-hygiene.md` | `docs/**`, `CLAUDE.md`, `.claude/rules/**` | Plan archival, doc hygiene, bucket migrations, single-file convention |
| `.claude/rules/rule-authoring.md` | `.claude/rules/**` | Copy shipped-fix form verbatim when adding rule examples |
| `.claude/rules/openpyxl.md` | builder `.py` files (not `styles.py`) | openpyxl/LO core: CF, formulas, Rule 3 styling, charts |
| `.claude/rules/openpyxl-infra.md` | `workbook.py`, `reader.py`, `config_sheet.py`, `scripts/**` | Workbook primitives, named-range auto-extend, section-reader guards |
