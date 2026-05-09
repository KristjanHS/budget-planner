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
DEFAULT_RATE = 10000                    # EUR per person-month (per-WP default rate)
EU_ANNUAL_CEILING_PM = 12.0              # Max PM per calendar year across all EU grants (≈215 days)
CFS_THRESHOLD = 430_000                  # EUR — CFS audit certificate required above this
DEFAULT_RP_LENGTH = 6                   # Months per reporting period (programme typical: 12-18)

# ── GRADE RATES ──
# EUR per person-month by job grade (bookkeeping-provided bands)
GRADE_RATES = {"Senior": 12500, "Mid": 10000, "Junior": 7500}
DEFAULT_GRADE = "Mid"

# ── WORK PACKAGES ──
# Each WP dict has: id, name, fte (total person-months), start (1-indexed
# project month), end (1-indexed project month), type (distribution curve:
# "flat", "front-loaded", "bell", "back-loaded"). Optional per-WP rate override.
# Invariant: exactly one entry per WP id; 1 <= start <= end <= PROJECT_MONTHS.
# Multi-segment schedules are not supported — every WP has a single contiguous
# active range. The funder audits that no time is logged outside this range.
# NOTE: WP3 seed was historically split M1-M3 + M5-M7 (demo "mid-WP gap");
# collapsed to M1-M7 for the single-range schema.

WORK_PACKAGES = [
    {"id": "WP1", "name": "Project Management",              "fte": 4.0,
     "start":  1, "end": 14, "type": "flat"},
    {"id": "WP2", "name": "Requirements and concept",     "fte":  6.0,
     "start":  1, "end": 18, "type": "front-loaded"},
    {"id": "WP3", "name": "Research and analysis", "fte": 6.0,
     "start":  1, "end": 7, "type": "front-loaded"},
    {"id": "WP4", "name": "Architecture design",             "fte": 12.0,
     "start":  2, "end": 11, "type": "bell"},
    {"id": "WP5", "name": "Component design",               "fte": 6.0,
     "start": 4, "end": 13, "type": "bell"},
    {"id": "WP6", "name": "Subsystem design",  "fte": 3.0,
     "start": 4, "end": 13, "type": "bell"},
    {"id": "WP7", "name": "Software development",         "fte": 3.0,
     "start": 6, "end": 14, "type": "bell"},
    {"id": "WP8", "name": "Prototyping",        "fte":  3.0,
     "start": 10, "end": 14, "type": "back-loaded"},
    {"id": "WP9", "name": "Validation and demo",       "fte":  3.0,
     "start": 12, "end": 14, "type": "back-loaded"},
]

# ── TEAM ──
# Seed data only — Config sheet roster is primary after first generation
TEAM = [
    {"name": "Alice", "max_fte": 1.0, "grade": "Senior"},
    {"name": "Bob",   "max_fte": 1.0, "grade": "Mid"},
    {"name": "Carol", "max_fte": 1.0, "grade": "Junior"},
    {"name": "Dave",  "max_fte": 1.0, "grade": "Mid"},
]

# ── EXTERNAL COMMITMENTS ──
# Seed data only — Excel Capacity sheet is primary after first generation
EXTERNAL_COMMITMENTS = [
    {"person": "Alice", "project": "ProjectX",        "fte": 2.0, "start": 1, "end": 11, "is_eu": False},
    {"person": "Bob",   "project": "ProjectX",        "fte": 2.0, "start": 1, "end": 14, "is_eu": False},
    {"person": "Bob",   "project": "HorizonEU-Alpha", "fte": 2.0, "start": 1, "end": 7, "is_eu": True},
]

# ── REPORTING PERIODS ──
# Seed data only — Config sheet is primary after first generation
# Boundaries define when Form C / financial statements are due
REPORTING_PERIODS = [
    {"id": "RP1", "start": 1, "end": 6},
    {"id": "RP2", "start": 7, "end": 12},
    {"id": "RP3", "start": 13, "end": 14},
]

# ── DEMO MODE ──
# Months of actuals to seed in demo mode (end of Year 1)
DEMO_CURRENT_MONTH = 12

# Capacity assignments for demo mode — passed to init_workbook(assignments=...).
# Each entry: person, WP id, total FTE, start month, end month.
#
# Engineered to surface multiple Timelog Solver stop-reason branches so the
# Stop reason column on Timelog Plan can be sanity-checked manually against
# the STOP REASONS reference table on the Config sheet:
#   WP1 → "+12% cap"      (Alice fills to 71400; 8th placement would breach +12%)
#   WP2 → "No window"     (Dave's row [20,24] does not overlap WP2's [1,18])
#   WP4 → "No candidates" (no PERSON-WP ASSIGNMENTS row references WP4)
#   WP9 → "-11% override" (Alice forced past +12% cap; skipping would breach -11%)
#   WP3, WP5, WP6, WP7, WP8 → "Filled"
DEMO_CAPACITY_ASSIGNMENTS = [
    {"person": "Alice", "wp": "WP1", "start": 1,  "end": 14, "total_fte": 4.0},
    {"person": "Alice", "wp": "WP3", "start": 1,  "end": 7, "total_fte": 3.0},
    {"person": "Alice", "wp": "WP9", "start": 12, "end": 14, "total_fte": 3.0},
    {"person": "Bob",   "wp": "WP5", "start": 4, "end": 13, "total_fte": 5.0},
    {"person": "Bob",   "wp": "WP6", "start": 4, "end": 13, "total_fte": 2.0},
    {"person": "Carol", "wp": "WP7", "start": 6, "end": 14, "total_fte": 3.0},
    {"person": "Carol", "wp": "WP8", "start": 10, "end": 14, "total_fte": 3.0},
    {"person": "Dave",  "wp": "WP2", "start": 6, "end": 7, "total_fte": 1.00},
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
# Posted-cumulative redesign (current naming):
VAL_COL_SLOPE = _vcols["slope"]
VAL_COL_BUDGET = _vcols["budget"]
VAL_COL_DIFF = _vcols["diff"]
VAL_COL_POSTED = _vcols["posted"]
VAL_COL_PCT_CEIL = _vcols["pct_ceil"]
VAL_COL_BREACH = _vcols["breach"]
# Backward-compat aliases (transitional; retire as Stages B/C/D migrate the
# remaining consumers — Plan / Capacity / Actuals-Hours / Plan-Forecast /
# Summary). Point at the same column slots.
VAL_COL_SUM = VAL_COL_SLOPE
VAL_COL_CUM = VAL_COL_POSTED


def default_project_params():
    """Return default project parameters dict from module-level constants."""
    return {
        "title": PROJECT_TITLE,
        "start": PROJECT_START,
        "months": PROJECT_MONTHS,
        "rate": DEFAULT_RATE,
        "eu_ceiling_pm": EU_ANNUAL_CEILING_PM,
        # Uniform vacation month: 7 = July. Matches the programme enhancement doc
        # default; readers fall back to 7 when the Config scalar is missing.
        "vacation_month_num": 7,
    }


def default_reporting_periods(project_months, rp_length=DEFAULT_RP_LENGTH):
    """Generate default reporting periods from project duration.

    Returns list of dicts with keys: id, start, end.
    """
    periods = []
    m = 1
    rp_num = 1
    while m <= project_months:
        end = min(m + rp_length - 1, project_months)
        periods.append({"id": f"RP{rp_num}", "start": m, "end": end})
        m = end + 1
        rp_num += 1
    return periods
