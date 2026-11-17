"""
Budget Planner — Project Configuration

Edit this file to change project-specific data.
All builder modules import from here.

After editing, re-run: python -m budget_planner && pytest tests/ -v
"""

import math
from datetime import datetime

from budget_planner.overlay_settings import load_overlay, overlay_dir_from_env

# Optional local overlay; resolution rules live in overlay_dir_from_env().
_OVERLAY = load_overlay(overlay_dir_from_env())

# ── PROJECT PARAMETERS ──

PROJECT_TITLE: str = _OVERLAY.get("PROJECT_TITLE", "Programme Project Budget Plan")
PROJECT_START: datetime = _OVERLAY.get("PROJECT_START", datetime(2026, 11, 1))  # First month (day must be 1)
# Total duration in months; the default 14 spans two calendar years, the first one partial (Nov-Dec)
PROJECT_MONTHS: int = _OVERLAY.get("PROJECT_MONTHS", 14)
DEFAULT_RATE: int = _OVERLAY.get("DEFAULT_RATE", 10000)  # EUR per person-month (per-WP default rate)
EU_ANNUAL_CEILING_PM = 12.0  # Max PM per calendar year across all EU grants (≈215 days)
CFS_THRESHOLD = 430_000  # EUR — CFS audit certificate required above this
DEFAULT_RP_LENGTH: int = _OVERLAY.get("RP_LENGTH", 6)  # Months per reporting period

# ── GRADE RATES ──
# EUR per person-month by job grade (bookkeeping-provided bands)
GRADE_RATES = {
    "Senior": _OVERLAY.get("GRADE_RATE_SENIOR", 12500),
    "Mid": _OVERLAY.get("GRADE_RATE_MID", 10000),
    "Junior": _OVERLAY.get("GRADE_RATE_JUNIOR", 7500),
}
DEFAULT_GRADE = "Mid"


def _months(lo, hi):
    """Window over the fraction (lo, hi] of PROJECT_MONTHS, e.g. _months(1 / 2, 1) = second half."""
    start = min(PROJECT_MONTHS, math.floor(round(lo * PROJECT_MONTHS, 9)) + 1)
    return {"start": start, "end": max(start, math.floor(round(hi * PROJECT_MONTHS, 9)))}


# ── WORK PACKAGES ──
# Each WP dict has: id, name, fte (total person-months), start (1-indexed
# project month), end (1-indexed project month), type (distribution curve:
# "flat", "front-loaded", "bell", "back-loaded"). Optional per-WP rate override.
# Invariant: exactly one entry per WP id; 1 <= start <= end <= PROJECT_MONTHS.
# Multi-segment schedules are not supported — every WP has a single contiguous
# active range. The funder audits that no time is logged outside this range.


def _wp_name(n, default):
    return _OVERLAY.get(f"WP{n}_NAME", default)


WORK_PACKAGES = [
    {"id": "WP1", "name": _wp_name(1, "Project Management"), "fte": 4.0, **_months(0, 1), "type": "flat"},
    {
        "id": "WP2",
        "name": _wp_name(2, "Requirements and concept"),
        "fte": 6.0,
        **_months(0, 1 / 2),
        "type": "front-loaded",
    },
    {
        "id": "WP3",
        "name": _wp_name(3, "Research and analysis"),
        "fte": 6.0,
        **_months(0, 1 / 2),
        "type": "front-loaded",
    },
    {"id": "WP4", "name": _wp_name(4, "Architecture design"), "fte": 12.0, **_months(1 / 4, 3 / 4), "type": "bell"},
    {"id": "WP5", "name": _wp_name(5, "Component design"), "fte": 6.0, **_months(1 / 3, 1), "type": "bell"},
    {
        "id": "WP6",
        "name": _wp_name(6, "Subsystem design"),
        "fte": 3.0,
        **_months(1 / 3, 1),
        "type": "bell",
    },
    {"id": "WP7", "name": _wp_name(7, "Software development"), "fte": 3.0, **_months(1 / 2, 1), "type": "bell"},
    {
        "id": "WP8",
        "name": _wp_name(8, "Prototyping"),
        "fte": 3.0,
        **_months(1 / 2, 1),
        "type": "back-loaded",
    },
    {
        "id": "WP9",
        "name": _wp_name(9, "Validation and demo"),
        "fte": 3.0,
        **_months(3 / 4, 1),
        "type": "back-loaded",
    },
]

# ── TEAM ──
# Seed data only — Config sheet roster is primary after first generation
TEAM = [
    {"name": "Alice", "max_fte": 1.0, "grade": "Senior"},
    {"name": "Bob", "max_fte": 1.0, "grade": "Mid"},
    {"name": "Carol", "max_fte": 1.0, "grade": "Junior"},
    {"name": "Dave", "max_fte": 1.0, "grade": "Mid"},
]

# ── EXTERNAL COMMITMENTS ──
# Seed data only — Excel Capacity sheet is primary after first generation
EXTERNAL_COMMITMENTS = [
    {"person": "Alice", "project": "ProjectX", "fte": 2.0, **_months(0, 1 / 2), "is_eu": False},
    {"person": "Bob", "project": "ProjectX", "fte": 2.0, **_months(0, 1), "is_eu": False},
    {"person": "Bob", "project": "HorizonEU-Alpha", "fte": 2.0, **_months(0, 1 / 2), "is_eu": True},
]


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


# ── REPORTING PERIODS ──
# Seed data only — Config sheet is primary after first generation
# Boundaries define when Form C / financial statements are due
REPORTING_PERIODS = default_reporting_periods(PROJECT_MONTHS)

# ── DEMO MODE ──
# Months of actuals to seed in demo mode: two thirds in, so the last third stays unseeded (forecast-only)
DEMO_CURRENT_MONTH = max(1, PROJECT_MONTHS * 2 // 3)

# Capacity assignments for demo mode — passed to init_workbook(assignments=...).
# Each entry: person, WP id, total FTE, start month, end month.
#
# Engineered to surface multiple Timelog Solver stop-reason branches so the
# Stop reason column on Timelog Plan can be sanity-checked manually against
# the STOP REASONS reference table on the Config sheet:
#   WP1 → "+12% cap"      (sole candidate Alice at Senior rate; a 4th month would breach +12%)
#   WP2 → "No window"     (Dave's row covers the second half; WP2 runs in the first half)
#   WP4 → "No candidates" (no PERSON-WP ASSIGNMENTS row references WP4)
#   WP9 → "-11% override" (Alice forced past +12% cap; skipping would breach -11%)
#   WP3, WP5, WP6, WP7, WP8 → "Filled" (verified at the default PROJECT_MONTHS and rates)
DEMO_CAPACITY_ASSIGNMENTS = [
    {"person": "Alice", "wp": "WP1", **_months(0, 1), "total_fte": 4.0},
    {"person": "Alice", "wp": "WP9", **_months(3 / 4, 1), "total_fte": 3.0},
    {"person": "Bob", "wp": "WP5", **_months(1 / 3, 1), "total_fte": 5.0},
    {"person": "Bob", "wp": "WP6", **_months(1 / 3, 1), "total_fte": 2.0},
    {"person": "Carol", "wp": "WP7", **_months(1 / 2, 1), "total_fte": 3.0},
    {"person": "Dave", "wp": "WP2", **_months(1 / 2, 1), "total_fte": 1.0},
    {"person": "Dave", "wp": "WP3", **_months(0, 1 / 2), "total_fte": 6.0},
    {"person": "Dave", "wp": "WP5", **_months(1 / 3, 1), "total_fte": 1.0},
    {"person": "Dave", "wp": "WP6", **_months(1 / 3, 1), "total_fte": 1.0},
    {"person": "Dave", "wp": "WP8", **_months(1 / 2, 1), "total_fte": 3.0},
]

# Spending multipliers per WP index for demo actuals.
# Each multiplier is applied to the planned monthly EUR to create
# diverse Trends scenarios: OK, OVER, SLOW, WATCH.
# WPs not yet active at DEMO_CURRENT_MONTH naturally show "—".
DEMO_ACTUALS_MULTIPLIERS = [
    1.00,  # WP1 (Project Management)              — on-track  → "OK"
    1.28,  # WP2 (Requirements and concept)     — overspend → "OVER"
    0.30,  # WP3 (Research and analysis) — underspend → "SLOW"
    1.00,  # WP4 (Architecture design)              — neutral
    1.00,  # WP5 (Component design)                — neutral
    1.00,  # WP6 (Subsystem design)   — neutral
    1.00,  # WP7 (Software development)          — neutral
    1.00,  # WP8 (Prototyping)         — neutral
    1.00,  # WP9 (Validation and demo)        — neutral
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
# Cross-semantic aliases — Plan / Capacity / Actuals-Hours / Summary read
# this column slot as a row-wide monthly-SUM, while Actuals reads it as
# Slope. Both names are canonical for their respective sheet contexts;
# tests for non-Actuals sheets import the SUM/CUM forms.
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
