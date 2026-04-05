# CLAUDE.md — Project Instructions for Claude Code

## What This Is

An openpyxl-based Excel workbook generator for EU project FTE budget planning. The user will iterate on this with you.

## Information Recording Principles

This file uses **progressive disclosure**. Keep it focused on what every conversation needs.

### Level 1 (this file) — always loaded

| Type | Example |
|------|---------|
| Core commands | `python -m budget_planner init` |
| Iron rules / prohibitions | Never hardcode calculated values |
| Error diagnostics | Symptom -> cause -> fix |
| Code patterns | Copy-paste ready snippets |
| Directory map | Function -> file |
| Trigger indexes | Pointers to Level 2 docs |

### Level 2 (`docs/`) — load on demand

| Type | Example |
|------|---------|
| Domain reference | programme cost categories, funding rates |
| Architecture deep-dives | Builder patterns, reader fragility |
| Historical plans | Completed and draft plans |

### When asked to record information

1. High-frequency or iron-rule? -> Add to this file (Level 1)
2. Detailed SOP, edge case, or reference? -> Add to `docs/` (Level 2)
3. Every Level 2 reference in this file must include a **trigger condition** (when to read it)

## Reference Index (encounter a problem? check here)

| Trigger | Document | Key content |
|---------|----------|-------------|
| Programme cost model, funding rates, 25% indirect | `docs/grant-domain.md` | Cost categories, PM->EUR chain, reallocation rules |
| How sheets are built, builder interface | `docs/deep-dive/Sheet Builders.md` | Builder contract, stateless pattern |
| Reader breaks, rebuild loses data | `docs/deep-dive/Reader and Rebuild.md` | Formula reference chains, fragile parsing |
| Column offsets, month positioning, styles | `docs/deep-dive/Layout and Styling.md` | `layout.py` constants, `styles.py` patterns |
| Project overview, tech stack | `docs/1. Project Overview.md` | Purpose, technology, actors |
| C4 architecture, component diagram | `docs/2. Architecture Overview.md` | System context, component graph |
| Init vs rebuild workflow | `docs/3. Workflow Overview.md` | CLI flow, data preservation |
| Brittleness, hardening tasks | `docs/plans/hardening-plan.md` | Duplicated constants, silent corruption risks |

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
6. **_META sheet must be updated** whenever you change the workbook structure. It's the primary documentation for future AI agents.
7. **Cell protection** *(temporarily disabled)*: When re-enabled, all sheets are protected (password "grant"). Only blue input cells are unlocked. See `_apply_protection()` in workbook.py.

## Before Modifying Code

| You want to change | Read first | Key trap |
|--------------------|-----------|----------|
| Sheet builder logic | `docs/deep-dive/Sheet Builders.md` | Builders are stateless; don't add state |
| Reader / rebuild | `docs/deep-dive/Reader and Rebuild.md` | Most fragile module; formula chain resolution |
| Column layout, month offsets | `docs/deep-dive/Layout and Styling.md` | `layout.py` is the single source; don't duplicate constants |
| Cost model, funding rules | `docs/grant-domain.md` | 25% indirect is extra budget, not deducted from ceiling |
| Hardening / refactoring | `docs/plans/hardening-plan.md` | Check if brittleness already identified there |
| Config column layout | `.claude/rules/dev-workflow.md` | Grep `Config!{letter}` in ALL builders + tests when columns shift |

## Rules Index (path-gated — auto-loaded when touching `**/*.py`)

| Rule file | Key content |
|-----------|-------------|
| `.claude/rules/dev-workflow.md` | Build/test, linting, openpyxl patterns, worktree protocol |
| `.claude/rules/architecture.md` | Directory map (function → file) |
| `.claude/rules/common-tasks.md` | Step-by-step for "add WP", "add team member", etc. |
