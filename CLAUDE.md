# CLAUDE.md — Project Instructions for Claude Code

## What This Is

An openpyxl-based Excel workbook generator for EU project FTE budget planning. The user will iterate on this with you.

<!-- L1=this file (always), L2=docs/ (on demand via trigger table), L3=.claude/rules/ (path-gated) -->

## Reference Index (encounter a problem? check here)

| Trigger | Read |
|---------|------|
| Programme cost model, funding rates, 25% indirect; C4 architecture, workflows, formula graph; sheet builders, stateless pattern; reader breaks, rebuild loses data; column offsets, month positioning, styles, row/column shifts, new sheet/column, tab-color conventions; cross-sheet invariant, trust boundary, derived-metric rule; Timelog Plan, solver, 1.0-FTE logging allocation, WP fill order, merged grid, freeze-point, per-year EUR cols, monthly total, yearly total, yearly budget, vacation month, Pool reduction; rebuild drops rows, stale Config refs; WP schedule single range, migration from legacy workbook, multi-segment error | `docs/spec.md` (also `budget_planner/layout.py` + `budget_planner/styles.py` for live column/row constants; `budget_planner/timelog_plan.py` + `budget_planner/timelog_solver.py` for Timelog solver internals) |
| Run commands, generate workbook, compare scenarios, dev setup, add WP / team member / actuals, cloud/LO env setup, code review & /impag, MCP tools for workbook inspection | `docs/runbook.md` (§ MCP Tools for openpyxl-read / mcp-excel) |
| Audit "which past incident drove this rule" — lazy imports, optional→required params, constant renames, kwarg drops, F401 echoes, hook bounded reads, multiset bystanders | `docs/refactor-incidents.md` |

## Critical Rules

1. **Always use Excel formulas, never hardcode calculated values in Python.** The spreadsheet must remain dynamic.
2. **After every generation: `python scripts/recalc.py output/budget_plan.xlsx` then `pytest tests/ -v`.** Skipping recalc makes `data_only=True` reads return None — tests pass silently on broken output. Verify all pass before presenting to user.
3. **Font color convention is sacred**: Blue = input, Black = formula, Green = cross-sheet ref. Never mix these.
4. **Inactive month cells must be greyed out and empty.** Auditors flag time logged outside WP active periods.
5. **Named ranges `project_start` and `fte_rate`** must always exist and point to Config!$B$2 and Config!$B$4.
6. **New architectural invariants** (cross-sheet constraint, trust boundary, derived-metric rule) must be recorded in `docs/spec.md` before merge. Rationale + scope carve-out in `.claude/rules/dev-workflow.md`.
7. **Skills that emit markdown by default** (retro, reflect, kaizen:*, mybrain, senior-architect, architecture-diagram-creator, brutal-honesty-review, deep-research): prefer conversation output unless the artifact survives past this session. Write-time gates and override sentinels are enforced by hooks in `.claude/hooks/` — see `.claude/hooks/README.md` (7 active) for the full table + troubleshooting. Active gates: `docs-bloat-gate` (reference docs / rules / audit-keyword `docs/*.md`), `docs-only-plan-gate` (`git mv` of plans → `archived/` requires fresh pytest-pass stamp), `external-process-test-gate` (`tests/**/*.py` outside `tests/manual/` blocks docker/kubectl/playwright/selenium/`requests` patterns). One-off analyses go to `docs/analysis/` (`.claudeignore`'d).

## Large Files

Thresholds + intent-declaration rule in `~/.claude/rules/reading-large-files.md` (auto-loads on `.py`/`.md` reads). Project-specific addendum: archived plans in `docs/plans/archived/` aren't listed — treat all of them as large by default. List active large files on demand: `wc -l budget_planner/*.py tests/*.py | sort -rn | head`.

## Rules Index (path-gated — auto-loaded when touching matching files)

| Rule file | Triggers on | Key content |
|-----------|-------------|-------------|
| `.claude/rules/dev-workflow.md` | `**/*.py` | Build/test, linting, formatting, invariants (core cross-cutting) |
| `.claude/rules/testing-project.md` | `tests/**`, `**/test_*.py` | Project testing rules (no hardcoded seeds, round-trip, docs-only stages) |
| `.claude/rules/plan-hygiene.md` | `docs/**`, `CLAUDE.md`, `.claude/rules/**` | Write-gate (no plan doc for ≤2-file/1-session work), delete-not-archive for shipped single-session plans, doc hygiene |
| `.claude/rules/rule-authoring.md` | `.claude/rules/**` | Copy shipped-fix form verbatim when adding rule examples |
| `.claude/rules/openpyxl.md` | builder `.py` files (not `styles.py`) | openpyxl/LO core: CF, formulas, Rule 3 styling, charts |
| `.claude/rules/openpyxl-infra.md` | `workbook.py`, `reader.py`, `config_sheet.py`, `scripts/**` | Workbook primitives, named-range auto-extend, section-reader guards |
| `.claude/rules/windows-scripts.md` | `*.bat`, `*.cmd` | Wrappers target cmd.exe — no PS subprocess; delegate complex logic to Python |
