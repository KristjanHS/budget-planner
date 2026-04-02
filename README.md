# Budget Planner

Excel workbook generator for tracking planned and actual FTE allocation per work package across a multi-year EU-funded project. Includes per-person capacity planning with overcommitment detection.

## Quick Start

```bash
# Install (requires Python 3.10+)
uv venv && uv pip install -e ".[dev]"

# Generate workbook (demo mode -- pre-filled with seed allocations)
python -m budget_planner

# Generate skeleton workbook (empty Capacity for manual entry)
python -m budget_planner --skeleton

# Update existing workbook (preserves user-entered data across all sheets)
python -m budget_planner --update output/budget_plan.xlsx

# Snapshot / restore for scenario comparison
python -m budget_planner --snapshot output/budget_plan.xlsx
python -m budget_planner --restore output/budget_plan.xlsx

# Recalculate formulas (requires LibreOffice)
python scripts/recalc.py output/budget_plan.xlsx

# Run tests
pytest tests/ -v
```

## Features

- **Plan sheet**: WP config table, planned FTE grid (derived from Capacity via cross-sheet SUM)
- **Team sheet**: Team roster (name, max FTE) and WP share seed matrix
- **Capacity sheet**: Per-person monthly FTE allocation across WPs and external projects, with TOTAL/REMAINING rows and overcommitment alerts (amber >80%, red >100%)
- **Actuals sheet**: EUR actuals per WP per month from accountant, with ceiling validation
- **Trends sheet**: Burn rate analysis, cumulative spend chart, person load overview
- **Summary sheet**: Annual aggregation with budget health indicators
- **_META sheet**: Self-documenting AI-agent-readable workbook structure
- **Round-trip updates**: `--update` preserves all user-entered data across regeneration
- **Scenario comparison**: `--snapshot` / `--restore` for before/after comparison

## Business Context

This tool tracks a single organisation's share of a multi-partner EU-funded (programme) project. It derives euro amounts from FTE using a configurable rate. Key constraints:

- **User**: Manages an R&D team's FTE across EU project work packages
- **Funder constraint**: Total FTE per WP must stay within budget. Auditors flag time logged to a WP outside its active period. Annual planned-vs-actual is informational, not strictly enforced.
- **Primary unit**: FTE (person-months). Euros are derived internally (FTE x monthly rate), not reported to funder.
- **Single org**: Only our partner's share is tracked. No multi-partner view needed.

## Project Structure

```
budget-planner/
├── README.md                 # This file
├── CONVENTIONS.md            # Spreadsheet conventions and AI-agent instructions
├── CLAUDE.md                 # Claude Code project instructions
├── budget_planner/
│   ├── __main__.py           # CLI entry point
│   ├── config.py             # WPs, team, external commitments, constants
│   ├── styles.py             # Font/fill/border constants
│   ├── plan.py               # Plan sheet builder
│   ├── team.py               # Team sheet builder
│   ├── capacity.py           # Capacity sheet builder
│   ├── actuals.py            # Actuals sheet builder
│   ├── trends.py             # Trends sheet builder
│   ├── summary.py            # Summary sheet builder
│   ├── meta.py               # _META sheet builder
│   ├── reader.py             # Round-trip reader (v1 + v2 format)
│   └── workbook.py           # Orchestrator (create/update/snapshot/restore)
├── tests/                    # pytest suite
├── scripts/
│   ├── recalc.py             # LibreOffice formula recalculator
│   └── office/               # LibreOffice helpers
└── output/
    └── budget_plan.xlsx  # Generated workbook (gitignored)
```

## Modifying

1. Edit `config.py` for WP data, team roster, external commitments, project duration, rates
2. Re-run: `python -m budget_planner && python scripts/recalc.py output/budget_plan.xlsx`
3. Verify: `pytest tests/ -v`

## Design Decisions

- Single flat Plan sheet with separate Planned/Actual blocks (not stacked rows per WP)
- Plan Planned FTE is derived from Capacity via cross-sheet SUM formulas
- Team + Capacity sheets for per-person FTE breakdown with overcommitment detection
- Even FTE spread across active months as default, with manual override per cell
- Single FTE rate -- all EUR derived from FTE x rate (cell E2)
- Inactive months greyed out and empty to prevent audit issues
- Demo mode seeds with calculated allocations; skeleton mode leaves cells empty
- `_META` sheet inside the workbook documents structure for AI agents

## Dependencies

- Python 3.10+
- `openpyxl` >= 3.1.0
- `python-dateutil`
- LibreOffice (for formula recalculation via `scripts/recalc.py`)
