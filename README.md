# Budget Planner

Excel workbook generator for tracking planned FTE allocation and actual EUR spend per work package across a multi-year EU-funded project. Includes per-person capacity planning with overcommitment detection.

## Quick Start

```bash
# Install (requires Python 3.14+)
uv venv && uv pip install -e ".[dev]"

# Create a fresh workbook from config.py seed data
python -m budget_planner init

# Rebuild from existing workbook (preserves Capacity FTE + Actuals EUR)
python -m budget_planner rebuild

# Generate change workbook (rebuild e2e test with modified config)
pytest tests/test_rebuild_e2e.py -q

# Snapshot / restore for scenario comparison
python -m budget_planner --snapshot output/budget_plan.xlsx
python -m budget_planner --restore output/budget_plan.xlsx

# Recalculate formulas (requires LibreOffice)
python scripts/recalc.py output/budget_plan.xlsx

# Run tests
pytest tests/ -v
```

## Features

### Sheets (data flow order)

- **Config**: Project parameters + WP definitions in an Excel Table (`tbl_WorkPackages`); the living source of truth for rebuilds
- **Team**: Team roster (name, grade, max FTE) and WP share seed matrix
- **Capacity**: Per-person monthly FTE allocation across WPs and external projects, with TOTAL/REMAINING rows and overcommitment alerts (amber >80%, red >100%)
- **Plan**: Planned FTE grid per WP (derived from Capacity via cross-sheet SUM)
- **Actuals**: EUR actuals per WP per month from accountant, with ceiling validation
- **Trends**: Burn rate analysis, cumulative spend chart, person load overview
- **Summary**: Annual aggregation with budget health indicators
- **_META**: Self-documenting AI-agent-readable workbook structure

### How it works

- **Round-trip rebuild**: `rebuild` reads the existing workbook and regenerates all sheets, preserving manual Capacity FTE and Actuals EUR entries
- **Cell protection**: All sheets protected (password "grant"), only blue input cells unlocked for editing
- **Column grouping**: Month columns grouped by year with auto-collapse; current year expanded on rebuild
- **Scenario comparison**: `--snapshot` / `--restore` for before/after comparison

## Business Context

This tool tracks a single organisation's share of a multi-partner EU-funded (programme) project. It derives euro amounts from FTE using a configurable rate. Key constraints:

- **User**: Manages an R&D team's FTE across EU project work packages
- **Funder constraint**: Total FTE per WP must stay within budget. Auditors flag time logged to a WP outside its active period. Annual planned-vs-actual is informational, not strictly enforced.
- **Primary unit**: FTE (person-months). Euros are derived internally (FTE x monthly rate), not reported to funder.
- **Single org**: Only our partner's share is tracked. No multi-partner view needed.

## Modifying

1. Edit the **Config sheet** (project params, work packages) or **Team sheet** (roster, WP shares) in Excel
2. Run: `python -m budget_planner rebuild && python scripts/recalc.py output/budget_plan.xlsx`
3. Capacity FTE allocations and Actuals EUR values are preserved across rebuilds
4. Verify: `pytest tests/ -v`

## Windows Quick Start (no WSL)

Only Python is required. Excel or LibreOffice is needed to view and recalculate the workbook.

1. Install **Python 3.14+** from [python.org](https://www.python.org/downloads/) — check **"Add to PATH"** during install
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

## Documentation

| Document | Description |
|----------|-------------|
| [Architecture](docs/architecture.md) | Project structure, C4 diagrams, build order, workflow sequences, cross-sheet formula graph |
| [Internals](docs/internals.md) | Sheet builders, reader/rebuild, layout/styling, CF rules, cross-sheet refs |
| [Programme Domain Reference](docs/grant-domain.md) | Cost categories, funding rates, 25% indirect rule |
| [Cloud Environment Setup](docs/cloud-env-setup.md) | LibreOffice install, pre-flight checks, failure modes |

## Dependencies

- Python 3.14+
- `openpyxl` >= 3.1.0
- `python-dateutil`
- LibreOffice (for formula recalculation via `scripts/recalc.py` — Linux/macOS only)
