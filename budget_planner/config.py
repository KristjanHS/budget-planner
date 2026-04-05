"""
Budget Planner — Project Configuration

Edit this file to change project-specific data.
All builder modules import from here.

After editing, re-run: python -m budget_planner && pytest tests/ -v
"""

from datetime import datetime

# ── PROJECT PARAMETERS ──

PROJECT_TITLE = "Programme Project Budget Plan"
PROJECT_START = datetime(2026, 11, 1)  # First month of the project
PROJECT_MONTHS = 14                    # Total project duration in months
DEFAULT_RATE = 7500                    # EUR per person-month (internal rate)
EU_ANNUAL_CEILING_PM = 12.0              # Max PM per calendar year across all EU grants (≈215 days)

# ── WORK PACKAGES ──
# Each WP dict must have: id, name, fte (total person-months), start (month number 1-based), end (month number 1-based),
# type (distribution curve: "flat", "front-loaded", "bell", "back-loaded")
# start and end define the active period — FTE can only be allocated within this range
# The funder audits that no time is logged outside a WP's active period

WORK_PACKAGES = [
    {"id": "WP1", "name": "Project Management",              "fte": 4.0, "start": 1,  "end": 14, "type": "flat"},
    {"id": "WP2", "name": "Requirements and concept",     "fte": 6.0,  "start": 1,  "end": 6, "type": "front-loaded"},
    {"id": "WP3", "name": "Research and analysis", "fte": 6.0, "start": 1,  "end": 7, "type": "front-loaded"},
    {"id": "WP4", "name": "Architecture design",             "fte": 12.0, "start": 2,  "end": 11, "type": "bell"},
    {"id": "WP5", "name": "Component design",               "fte": 6.0, "start": 4, "end": 13, "type": "bell"},
    {"id": "WP6", "name": "Subsystem design",  "fte": 3.0, "start": 4, "end": 13, "type": "bell"},
    {"id": "WP7", "name": "Software development",         "fte": 3.0, "start": 6, "end": 14, "type": "bell"},
    {"id": "WP8", "name": "Prototyping",        "fte": 3.0,  "start": 10, "end": 14, "type": "back-loaded"},
    {"id": "WP9", "name": "Validation and demo",       "fte": 3.0,  "start": 12, "end": 14, "type": "back-loaded"},
]

# ── TEAM ──
# Seed data only — Config sheet roster is primary after first generation
TEAM = [
    {"name": "Alice", "max_fte": 1.0},
    {"name": "Bob",   "max_fte": 1.0},
    {"name": "Carol", "max_fte": 1.0},
    {"name": "Dave",  "max_fte": 1.0},
]

# ── EXTERNAL COMMITMENTS ──
# Seed data only — Excel Capacity sheet is primary after first generation
EXTERNAL_COMMITMENTS = [
    {"person": "Alice", "project": "ProjectX", "fte": 2.0, "start": 1, "end": 11},
    {"person": "Bob",   "project": "ProjectX", "fte": 2.0, "start": 1, "end": 14},
]

# ── DEMO MODE ──
# Months of actuals to seed in demo mode (end of Year 1)
DEMO_CURRENT_MONTH = 12

# Capacity assignments for demo mode — passed to init_workbook(assignments=...).
# Each entry: person, WP id, total FTE, start month, end month.
DEMO_CAPACITY_ASSIGNMENTS = [
    {"person": "Alice", "wp": "WP1", "start": 1,  "end": 14, "total_fte": 4.0},
    {"person": "Alice", "wp": "WP2", "start": 1,  "end": 6, "total_fte": 3.0},
    {"person": "Alice", "wp": "WP3", "start": 1,  "end": 7, "total_fte": 3.0},
    {"person": "Bob",   "wp": "WP4", "start": 2,  "end": 11, "total_fte": 2.0},
    {"person": "Bob",   "wp": "WP5", "start": 4, "end": 13, "total_fte": 5.0},
    {"person": "Bob",   "wp": "WP6", "start": 4, "end": 13, "total_fte": 2.0},
    {"person": "Carol", "wp": "WP7", "start": 6, "end": 14, "total_fte": 3.0},
    {"person": "Carol", "wp": "WP8", "start": 10, "end": 14, "total_fte": 3.0},
    {"person": "Carol", "wp": "WP9", "start": 12, "end": 14, "total_fte": 3.0},
    {"person": "Dave",  "wp": "WP3", "start": 1,  "end": 7, "total_fte": 6.0},
    {"person": "Dave",  "wp": "WP5", "start": 4, "end": 13, "total_fte": 1.0},
    {"person": "Dave",  "wp": "WP6", "start": 4, "end": 13, "total_fte": 1.0},
    {"person": "Dave",  "wp": "WP8", "start": 10, "end": 14, "total_fte": 3.0},
]

# Spending multipliers per WP index for demo actuals.
# Each multiplier is applied to the planned monthly EUR to create
# diverse Trends scenarios: OK, OVER, SLOW, WATCH.
# WPs not yet active at DEMO_CURRENT_MONTH naturally show "—".
DEMO_ACTUALS_MULTIPLIERS = [
    1.00,  # WP1 (Project Management)              — on-track  → "OK"
    1.28,  # WP2 (Requirements and concept)     — overspend → "OVER"
    0.30,  # WP3 (Research and analysis) — underspend → "SLOW"
    1.00,  # WP4 (Architecture design)              — N/A (starts M2)
    1.00,  # WP5 (Component design)                — N/A (starts M4)
    1.00,  # WP6 (Subsystem design)   — N/A (starts M4)
    1.00,  # WP7 (Software development)          — N/A (starts M6)
    1.00,  # WP8 (Prototyping)         — N/A (starts M10)
    1.00,  # WP9 (Validation and demo)        — N/A (starts M12)
]

# ── LAYOUT CONSTANTS ──
from budget_planner.layout import MONTH_COL_START, val_columns as _val_columns  # noqa: E402

# ── DERIVED (do not edit) ──
NUM_WPS = len(WORK_PACKAGES)
MONTH_COL_END = MONTH_COL_START + PROJECT_MONTHS - 1
_vcols = _val_columns(PROJECT_MONTHS)
VAL_COL_SUM = _vcols["sum"]
VAL_COL_BUDGET = _vcols["budget"]
VAL_COL_DIFF = _vcols["diff"]
VAL_COL_CUM = _vcols["cum"]
VAL_COL_PCT_CEIL = _vcols["pct_ceil"]
VAL_COL_BREACH = _vcols["breach"]


def default_project_params():
    """Return default project parameters dict from module-level constants."""
    return {
        "title": PROJECT_TITLE,
        "start": PROJECT_START,
        "months": PROJECT_MONTHS,
        "rate": DEFAULT_RATE,
        "eu_ceiling_pm": EU_ANNUAL_CEILING_PM,
    }
