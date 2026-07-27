# CLAUDE.md — Project Instructions for Claude Code

## What This Is

An openpyxl-based Excel workbook generator for EU project FTE budget planning. The user will iterate on this with you.

<!-- L1=this file (always), L2=.claude/rules/ (path-gated), L3=docs/ (on demand via trigger table) -->

## Reference Index (encounter a problem? check here)

| Trigger | Read |
|---------|------|
| Programme cost model, funding rates, 25% indirect; C4 architecture, workflows, formula graph; sheet builders, stateless pattern; reader breaks, rebuild loses data; column offsets, month positioning, styles, row/column shifts, new sheet/column, tab-color conventions; cross-sheet invariant, trust boundary, derived-metric rule; Timelog Plan, solver, 1.0-FTE logging allocation, WP fill order, merged grid, freeze-point, per-year EUR cols, monthly total, yearly total, yearly budget, vacation month, Pool reduction; rebuild drops rows, stale Config refs; WP schedule single range, migration from legacy workbook, multi-segment error | `docs/spec.md` (also `budget_planner/layout.py` + `budget_planner/styles.py` for live column/row constants; `budget_planner/timelog_plan.py` + `budget_planner/timelog_solver.py` for Timelog solver internals) |
| Run commands, generate workbook, compare scenarios, dev setup, add WP / team member / actuals, cloud/LO env setup, code review & /impag, MCP tools for workbook inspection | `docs/spec.md` (§ Commands / § Ops Quickref / § Dev Workflow — Review, Planning, Worktrees, MCP) |
| Audit "which past incident drove this rule" — lazy imports, optional→required params, constant renames, kwarg drops, F401 echoes, hook bounded reads, multiset bystanders | `docs/spec.md` § Refactor Incidents |

## Critical Rules

1. **Always use Excel formulas, never hardcode calculated values in Python.** The spreadsheet must remain dynamic.
2. **After every generation: `make check`** (= `python scripts/recalc.py output/budget_plan.xlsx` then `pytest tests/ -q`). Skipping recalc makes `data_only=True` reads return None — tests pass silently on broken output. Verify all pass before presenting to user.
3. **Font color convention is sacred**: Blue = input, Black = formula, Green = cross-sheet ref. Never mix these.
4. **Inactive month cells must be greyed out and empty.** Auditors flag time logged outside WP active periods.
5. **Named ranges `project_start` and `fte_rate`** must always exist and point to Config!$B$2 and Config!$B$4.
6. **New architectural invariants** (cross-sheet constraint, trust boundary, derived-metric rule) must be recorded in `docs/spec.md` before merge. Rationale + scope carve-out in `.claude/rules/dev-workflow.md`.
7. **Skills that emit markdown by default** (retro, reflect, kaizen:*, mybrain, senior-architect, architecture-diagram-creator, brutal-honesty-review, deep-research) fall under the global conversation-output default. This project owns no hooks — every gate that fires here is global (`external-process-test-gate`, `docs-bloat-gate`; behaviour + troubleshooting in `~/.claude/hooks/README.md`). Archiving a docs-only plan is no longer machine-gated: run `pytest tests/ -q` yourself first, because `docs/spec.md` is test-coupled. Both that gate and the `docs/analysis/` home for one-off analyses: `.claude/rules/plan-hygiene.md` (§ Doc hygiene).

## Large Files

Archived plans in `docs/plans/archive/` aren't listed anywhere — treat all of them as large by default. List active large files on demand: `wc -l budget_planner/*.py tests/*.py | sort -rn | head`.

## Rules Index (path-gated — auto-loaded when touching matching files)

| Rule file | Triggers on | Key content |
|-----------|-------------|-------------|
| `.claude/rules/dev-workflow.md` | `**/*.py` | Build/test, linting, formatting, invariants (core cross-cutting) |
| `.claude/rules/testing-project.md` | `tests/**`, `**/test_*.py` | Project testing rules (no hardcoded seeds, round-trip, docs-only stages) |
| `.claude/rules/plan-hygiene.md` | `docs/**`, `CLAUDE.md`, `.claude/rules/**` | Write-gate (no plan doc for ≤2-file/1-session work), delete-not-archive for shipped single-session plans, archive mechanics (pathspec both halves, merge-base gate), cite-by-stem, §-heading grep |
| `.claude/rules/openpyxl.md` | builder `.py` files (not `styles.py`) | openpyxl/LO core: CF, formulas, Rule 3 styling, charts |
| `.claude/rules/openpyxl-infra.md` | `workbook.py`, `reader.py`, `config_sheet.py`, `scripts/**` | Workbook primitives, named-range auto-extend, section-reader guards |
| `.claude/rules/windows-scripts.md` | `*.bat`, `*.cmd` | Wrappers target cmd.exe — no PS subprocess; delegate complex logic to Python |
