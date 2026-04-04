# Budget Planner

Excel workbook generator for tracking planned FTE allocation and actual EUR spend per work package across a multi-year EU-funded project. Includes per-person capacity planning with overcommitment detection.

## Quick Start

```bash
# Install (requires Python 3.10+)
uv venv && uv pip install -e ".[dev]"

# Create a fresh workbook from config.py seed data
python -m budget_planner init

# Rebuild from existing workbook (preserves Capacity FTE + Actuals EUR)
python -m budget_planner rebuild

# Snapshot / restore for scenario comparison
python -m budget_planner --snapshot output/budget_plan.xlsx
python -m budget_planner --restore output/budget_plan.xlsx

# Recalculate formulas (requires LibreOffice)
python scripts/recalc.py output/budget_plan.xlsx

# Run tests
pytest tests/ -v
```

## Features

- **Config sheet**: Project parameters + WP definitions in an Excel Table (`tbl_WorkPackages`); the living source of truth for rebuilds
- **Plan sheet**: WP config table, planned FTE grid (derived from Capacity via cross-sheet SUM)
- **Team sheet**: Team roster (name, max FTE) and WP share seed matrix
- **Capacity sheet**: Per-person monthly FTE allocation across WPs and external projects, with TOTAL/REMAINING rows and overcommitment alerts (amber >80%, red >100%)
- **Actuals sheet**: EUR actuals per WP per month from accountant, with ceiling validation
- **Trends sheet**: Burn rate analysis, cumulative spend chart, person load overview
- **Summary sheet**: Annual aggregation with budget health indicators
- **_META sheet**: Self-documenting AI-agent-readable workbook structure
- **Round-trip rebuild**: `rebuild` reads Config/Team/Capacity/Actuals from the existing workbook and regenerates all sheets, preserving manual Capacity FTE and Actuals EUR entries
- **Cell protection**: All sheets protected (password "grant"), only blue input cells are unlocked for editing
- **Column grouping**: Month columns grouped by year with auto-collapse; current year expanded on rebuild
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
├── CONVENTIONS.md            # Spreadsheet formatting, structural patterns, cross-sheet refs
├── CLAUDE.md                 # Claude Code project instructions
├── budget_planner/
│   ├── __main__.py           # CLI: init, rebuild subcommands
│   ├── config.py             # WPs, team, external commitments, constants (seed for init)
│   ├── config_sheet.py       # Config sheet builder + reader
│   ├── styles.py             # Font/fill/border constants, column grouping
│   ├── plan.py               # Plan sheet builder
│   ├── team.py               # Team sheet builder
│   ├── capacity.py           # Capacity sheet builder
│   ├── actuals.py            # Actuals sheet builder
│   ├── trends.py             # Trends sheet builder
│   ├── summary.py            # Summary sheet builder
│   ├── meta.py               # _META sheet builder
│   ├── reader.py             # Round-trip reader (Config, Team, Capacity, Actuals)
│   └── workbook.py           # Orchestrator: init_workbook() + rebuild_workbook()
├── tests/                    # pytest suite
├── scripts/
│   ├── recalc.py             # LibreOffice formula recalculator
│   └── office/               # LibreOffice helpers
└── output/
    └── budget_plan.xlsx  # Generated workbook (gitignored)
```

## Modifying

1. Edit the **Config sheet** (project params, work packages) or **Team sheet** (roster, WP shares) in Excel
2. Run: `python -m budget_planner rebuild && python scripts/recalc.py output/budget_plan.xlsx`
3. Capacity FTE allocations and Actuals EUR values are preserved across rebuilds
4. Verify: `pytest tests/ -v`

## Design Decisions

- **Config sheet is the living source of truth** -- project params and WP definitions live in Excel, not Python; `config.py` is only used as seed data for `init`
- **Round-trip preservation** -- `rebuild` reads the existing workbook and re-injects Capacity FTE and Actuals EUR values, matched by WP ID + person name + month index
- **Cell protection** -- all sheets protected (password "grant"), only blue input cells unlocked; prevents accidental formula overwrites
- **Column grouping** -- month columns grouped by year with auto-collapse for easier navigation of 60-column grids
- Plan sheet has a single Planned FTE block; actuals are tracked in a separate Actuals sheet (EUR from accountant)
- Plan Planned FTE is derived from Capacity via cross-sheet SUM formulas
- Team + Capacity sheets for per-person FTE breakdown with overcommitment detection
- Even FTE spread across active months as default, with manual override per cell
- Single FTE rate -- all EUR derived from FTE x rate
- Inactive months greyed out and empty to prevent audit issues
- Demo mode seeds with calculated allocations (`init --demo`); plain `init` leaves Capacity cells empty
- `_META` sheet inside the workbook documents structure for AI agents

## Windows Quick Start (no WSL)

Only Python is required. Excel or LibreOffice is needed to view and recalculate the workbook.

1. Install **Python 3.10+** from [python.org](https://www.python.org/downloads/) — check **"Add to PATH"** during install
2. Download and extract the repository ZIP
3. Double-click the batch files:

| File | What it does |
|------|-------------|
| `setup.bat` | Creates a virtual environment and installs dependencies (run once) |
| `init.bat` | Generates an empty workbook |
| `demo.bat` | Generates a workbook with demo data |
| `rebuild.bat` | Rebuilds from an existing workbook, preserving your data |

After each generation, open `output\budget_plan.xlsx` in Excel (recalculates automatically) or LibreOffice Calc (`Ctrl+Shift+F9`), then save.

For snapshot/restore and other commands, open a terminal in the project folder:

```powershell
.venv\Scripts\activate
python -m budget_planner --snapshot output/budget_plan.xlsx
python -m budget_planner --restore output/budget_plan.xlsx
```

## Dependencies

- Python 3.10+
- `openpyxl` >= 3.1.0
- `python-dateutil`
- LibreOffice (for formula recalculation via `scripts/recalc.py` — Linux/macOS only)
