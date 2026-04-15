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
| Column offsets, month positioning, styles, row/column shifts, new sheet/column | `docs/internals.md` § Layout and Styling + `budget_planner/layout.py` | `layout.py` is single source; structure lives in code + internals.md; no prose duplicate |
| Cross-sheet invariant, trust boundary, derived-metric rule | `docs/design-invariants.md` | Two-cadence actuals, single-month-cell rule; update before merge |
| C4 architecture, workflows, formula graph | `docs/architecture.md` | Mermaid diagrams, build order, cross-sheet dependencies |
| Directory map (function → file) | `docs/directory-map.md` | Quick orientation: which file does what |
| Cloud/container env, LO missing | `docs/cloud-env-setup.md` | Pre-flight checks, failure modes |
| Code review, /impag, worktrees, MCP tools for workbook inspection | `docs/dev-workflow-review.md` (§ MCP Tools for openpyxl-read / mcp-excel) | Grep-cascade checks, planning protocol, worktree sync, MCP usage |
| Add WP, add team member, add actuals, compare scenarios | `docs/common-tasks.md` | Step-by-step SOPs for common data tasks |
| Rebuild drops rows, stale Config refs | `docs/plans/archived/2026-04-12-rebuild-shape-drift-design.md` | Readers use named ranges; tolerate internal blanks |
| Debugging against user's confidential workbook | `docs/debugging-your-real-file.md` | Run `make_reproducer.py`; never send the real file |
| "record this", "remember this", organizing docs | `~/.claude/references/recording-principles.md` | L1/L2/L3 criteria, what belongs where |

## Key Commands

```bash
# Set up dev environment
uv venv && uv pip install -e ".[dev]"

# Create a fresh empty workbook from config.py seed data
python -m budget_planner init

# Create a workbook with demo Capacity + Actuals data
python -m budget_planner init --demo

# Rebuild from existing workbook (preserves assignments + Actuals EUR)
python -m budget_planner rebuild

# Snapshot / restore for scenario comparison
python -m budget_planner --snapshot output/budget_plan.xlsx
python -m budget_planner --restore output/budget_plan.xlsx

# Recalculate formulas (MANDATORY after any generation)
python scripts/recalc.py output/budget_plan.xlsx

# Run full test suite
pytest tests/ -v
```

## Critical Rules

1. **Always use Excel formulas, never hardcode calculated values in Python.** The spreadsheet must remain dynamic.
2. **After every generation, run `pytest tests/ -v`** and verify all pass before presenting to user.
3. **Font color convention is sacred**: Blue = input, Black = formula, Green = cross-sheet ref. Never mix these.
4. **Inactive month cells must be greyed out and empty.** Auditors flag time logged outside WP active periods.
5. **Named ranges `project_start` and `fte_rate`** must always exist and point to Config!$B$2 and Config!$B$4.

## Rules Index (path-gated — auto-loaded when touching `**/*.py`)

| Rule file | Key content |
|-----------|-------------|
| `.claude/rules/dev-workflow.md` | Build/test, linting, formatting, testing, invariants (incl. design-invariants.md update-before-merge rule) |
| `.claude/rules/openpyxl-patterns.md` | openpyxl/LO gotchas (builders, styles, workbook only) |

## Reference Index (tail echo)

| Trigger | Read |
|---------|------|
| Programme cost model / funding rates | `docs/grant-domain.md` |
| Sheet builders, readers, rebuild, layout | `docs/internals.md` + `budget_planner/layout.py` |
| Cross-sheet invariants / trust boundaries | `docs/design-invariants.md` |
| Architecture, formula graph, workflows | `docs/architecture.md` |
| Directory map (function → file) | `docs/directory-map.md` |
| Cloud/container env setup | `docs/cloud-env-setup.md` |
| Code review, /impag, worktrees, MCP tools | `docs/dev-workflow-review.md` |
| Common data tasks (add WP, team member, actuals) | `docs/common-tasks.md` |
| Rebuild drops rows / stale Config refs | `docs/plans/archived/2026-04-12-rebuild-shape-drift-design.md` |
| Debugging against confidential real workbook | `docs/debugging-your-real-file.md` |
| Recording principles (L1/L2/L3) | `~/.claude/references/recording-principles.md` |
