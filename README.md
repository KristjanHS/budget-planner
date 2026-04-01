# Budget Planner

Excel workbook generator for tracking planned and actual FTE allocation per work package across a multi-year EU-funded project. Includes per-person capacity planning with overcommitment detection.

## Quick Start

```bash
# Install (requires Python 3.10+)
uv venv && uv pip install -e ".[dev]"

# Generate workbook (demo mode — pre-filled with seed allocations)
python -m budget_planner

# Generate skeleton workbook (empty Capacity for manual entry)
python -m budget_planner --skeleton

# Recalculate formulas (requires LibreOffice)
python scripts/recalc.py output/budget_plan.xlsx

# Run tests (165 tests)
pytest tests/ -v
```

## Features

- **Plan sheet**: WP config table, planned FTE grid, actual FTE grid with budget validation
- **Team sheet**: Team roster (name, max FTE) and WP share seed matrix
- **Capacity sheet**: Per-person monthly FTE allocation across WPs and external projects, with TOTAL/REMAINING rows and overcommitment alerts (amber >80%, red >100%)
- **Summary sheet**: Annual aggregation with budget health indicators
- **_META sheet**: Self-documenting AI-agent-readable workbook structure
- **Round-trip updates**: `--update` preserves all user-entered data across regeneration
- **Scenario comparison**: `--snapshot` / `--restore` for before/after comparison

## Project Structure

```
budget-planner/
├── README.md                 # This file
├── SPEC.md                   # Full project specification
├── CONVENTIONS.md            # Spreadsheet conventions and AI-agent instructions
├── DESIGN_DECISIONS.md       # Brainstorming session record
├── CLAUDE.md                 # Claude Code project instructions
├── budget_planner/
│   ├── __main__.py           # CLI entry point
│   ├── config.py             # WPs, team, external commitments, constants
│   ├── styles.py             # Font/fill/border constants
│   ├── plan.py               # Plan sheet builder
│   ├── team.py               # Team sheet builder
│   ├── capacity.py           # Capacity sheet builder
│   ├── summary.py            # Summary sheet builder
│   ├── meta.py               # _META sheet builder
│   ├── reader.py             # Round-trip reader (v1 + v2 format)
│   └── workbook.py           # Orchestrator (create/update/snapshot/restore)
├── tests/                    # 165 tests (pytest)
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

See `SPEC.md` for full context. Key points:

- Single flat Plan sheet with separate Planned/Actual blocks (not stacked rows)
- Team + Capacity sheets for per-person FTE breakdown with overcommitment detection
- Plan Planned FTE is derived from Capacity via cross-sheet SUM formulas
- Summary sheet aggregates by year
- `_META` sheet inside the workbook documents structure for AI agents
- All EUR derived from FTE × rate (cell E2)
- Inactive months greyed out to prevent audit issues
- Demo mode seeds with calculated allocations; skeleton mode leaves cells empty

## Dependencies

- Python 3.10+
- `openpyxl` >= 3.1.0
- `python-dateutil`
- LibreOffice (for formula recalculation via `scripts/recalc.py`)
