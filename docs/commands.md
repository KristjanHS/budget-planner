# Commands

## Dev environment setup

```bash
uv venv && uv pip install -e ".[dev]"
```

## Generate a workbook

```bash
# Empty workbook from config.py seeds
python -m budget_planner init

# Demo workbook with Capacity + Actuals data
python -m budget_planner init --demo

# Rebuild from existing workbook (preserves assignments + Actuals EUR)
python -m budget_planner rebuild
```

After any generation, run recalc + tests — see CLAUDE.md Critical Rule #2.

## Scenario comparison

```bash
python -m budget_planner --snapshot output/budget_plan.xlsx
python -m budget_planner --restore output/budget_plan.xlsx
```

## Recalc + test

```bash
python scripts/recalc.py output/budget_plan.xlsx
pytest tests/ -v
```
