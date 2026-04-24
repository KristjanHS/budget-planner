# CLAUDE.md — Project Instructions for Claude Code

## What This Is

An openpyxl-based Excel workbook generator for EU project FTE budget planning. The user will iterate on this with you.

<!-- L1=this file (always), L2=docs/ (on demand via trigger table), L3=.claude/rules/ (path-gated) -->

## Reference Index (encounter a problem? check here)

| Trigger | Read |
|---------|------|
| Programme cost model, funding rates, 25% indirect | `docs/grant-domain.md` |
| Sheet builders, stateless pattern | `docs/internals.md` § Sheet Builders |
| Reader breaks, rebuild loses data | `docs/internals.md` § Reader and Rebuild |
| Column offsets, month positioning, styles, row/column shifts, new sheet/column, tab-color conventions | `budget_planner/layout.py` + `budget_planner/styles.py` + `docs/internals.md` § Styling conventions |
| Cross-sheet invariant, trust boundary, derived-metric rule | `docs/design-invariants.md` |
| C4 architecture, workflows, formula graph | `docs/architecture.md` |
| Cloud/container env, LO missing, confidential-workbook debugging | `docs/ops-quickref.md` |
| Code review, /impag, worktrees, MCP tools for workbook inspection | `docs/dev-workflow-review.md` (§ MCP Tools for openpyxl-read / mcp-excel) |
| Add WP, add team member, add actuals, compare scenarios | `docs/common-tasks.md` |
| Timelog Plan, solver, 1.0-FTE logging allocation, WP fill order, merged grid, freeze-point, per-year EUR cols, monthly total, yearly total, yearly budget, vacation month, Pool reduction | `budget_planner/timelog_plan.py` + `budget_planner/timelog_solver.py` + `docs/plans/archived/2026-04-23-solver-freeze-point-design.md` + `docs/plans/archived/2026-04-19-timelog-plan-enhancements-design.md` + `docs/plans/archived/2026-04-19-timelog-plan-v2-design.md` + `docs/plans/archived/2026-04-18-timelog-plan-design.md` |
| Rebuild drops rows, stale Config refs | `docs/plans/archived/2026-04-12-rebuild-shape-drift-design.md` |
| WP schedule single range, migration from legacy workbook, multi-segment error | `docs/plans/archived/2026-04-19-wp-gap-removal-design.md` + `docs/design-invariants.md` § WP schedule — single active range |
| Run commands, generate workbook, compare scenarios, dev setup | `docs/commands.md` |

## Critical Rules

1. **Always use Excel formulas, never hardcode calculated values in Python.** The spreadsheet must remain dynamic.
2. **After every generation: `python scripts/recalc.py output/budget_plan.xlsx` then `pytest tests/ -v`.** Skipping recalc makes `data_only=True` reads return None — tests pass silently on broken output. Verify all pass before presenting to user.
3. **Font color convention is sacred**: Blue = input, Black = formula, Green = cross-sheet ref. Never mix these.
4. **Inactive month cells must be greyed out and empty.** Auditors flag time logged outside WP active periods.
5. **Named ranges `project_start` and `fte_rate`** must always exist and point to Config!$B$2 and Config!$B$4.
6. **New architectural invariants** (cross-sheet constraint, trust boundary, derived-metric rule) must be recorded in `docs/design-invariants.md` before merge. Rationale + scope carve-out in `.claude/rules/dev-workflow.md`.
7. **Skills that emit markdown by default (retro, reflect, kaizen:*, mybrain, senior-architect, architecture-diagram-creator, brutal-honesty-review, deep-research) should output to conversation, not files, unless (a) the artifact survives past this session AND (b) no existing doc covers the topic. Prefer extending `docs/design-invariants.md` or a code docstring over creating a new file. One-off analyses that must be written go to `docs/analysis/` (`.claudeignore`'d).

## Large Files

Thresholds + intent-declaration rule in `~/.claude/rules/reading-large-files.md` (auto-loads on `.py`/`.md` reads). Project-specific addendum: archived plans in `docs/plans/archived/` aren't listed — treat all of them as large by default. List active large files on demand: `wc -l budget_planner/*.py tests/*.py | sort -rn | head`.

## Rules Index (path-gated — auto-loaded when touching matching files)

| Rule file | Triggers on | Key content |
|-----------|-------------|-------------|
| `.claude/rules/dev-workflow.md` | `**/*.py` | Build/test, linting, formatting, invariants (core cross-cutting) |
| `.claude/rules/testing-project.md` | `tests/**`, `**/test_*.py` | Project testing rules (no hardcoded seeds, round-trip, docs-only stages) |
| `.claude/rules/plan-hygiene.md` | `docs/**`, `CLAUDE.md`, `.claude/rules/**` | Plan archival, doc hygiene, bucket migrations, single-file convention |
| `.claude/rules/rule-authoring.md` | `.claude/rules/**` | Copy shipped-fix form verbatim when adding rule examples |
| `.claude/rules/openpyxl.md` | builder `.py` files (not `styles.py`) | openpyxl/LO core: CF, formulas, Rule 3 styling, charts |
| `.claude/rules/openpyxl-infra.md` | `workbook.py`, `reader.py`, `config_sheet.py`, `scripts/**` | Workbook primitives, named-range auto-extend, section-reader guards |
| `.claude/rules/windows-scripts.md` | `*.bat`, `*.cmd` | Wrappers target cmd.exe — no PS subprocess; delegate complex logic to Python |
