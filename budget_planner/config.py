"""
Budget Planner — Project Configuration

Edit this file to change project-specific data.
All builder modules import from here.

After editing, re-run: python -m budget_planner && pytest tests/ -v
"""

from datetime import datetime

# ── PROJECT PARAMETERS ──

PROJECT_TITLE = "Programme Project Budget Plan"
PROJECT_START = datetime(2026, 7, 1)  # First month of the project
PROJECT_MONTHS = 60                    # Total project duration in months
DEFAULT_RATE = 7500                    # EUR per person-month (internal rate)

# ── WORK PACKAGES ──
# Each WP dict must have: id, name, fte (total person-months), start (month number 1-based), end (month number 1-based)
# start and end define the active period — FTE can only be allocated within this range
# The funder audits that no time is logged outside a WP's active period

WORK_PACKAGES = [
    {"id": "WP1", "name": "Project Management",    "fte": 4.0, "start": 1,  "end": 60},
    {"id": "WP2", "name": "Requirements & Design",  "fte": 8.0,  "start": 1,  "end": 18},
    {"id": "WP3", "name": "Core Development",       "fte": 24.0, "start": 7,  "end": 42},
    {"id": "WP4", "name": "Integration & Testing",  "fte": 15.0, "start": 19, "end": 48},
    {"id": "WP5", "name": "Pilot Deployment",       "fte": 10.0, "start": 37, "end": 54},
    {"id": "WP6", "name": "Dissemination",          "fte": 6.0,  "start": 13, "end": 60},
    {"id": "WP7", "name": "Ethics & Compliance",    "fte": 4.0,  "start": 1,  "end": 60},
    {"id": "WP8", "name": "Exploitation",           "fte": 5.0,  "start": 25, "end": 60},
]

# ── TEAM ──
# Seed data only — Excel Team sheet is primary after first generation
TEAM = [
    {"name": "Alice", "max_fte": 1.0, "wps": {"WP1": 0.25, "WP2": 0.50, "WP3": 0.25}},
    {"name": "Bob",   "max_fte": 1.0, "wps": {"WP1": 0.25, "WP3": 0.50, "WP7": 0.25}},
    {"name": "Carol", "max_fte": 1.0, "wps": {"WP4": 0.40, "WP5": 0.40, "WP6": 0.20}},
    {"name": "Dave",  "max_fte": 1.0, "wps": {"WP3": 0.25, "WP4": 0.60, "WP8": 0.15}},
]

# ── EXTERNAL COMMITMENTS ──
# Seed data only — Excel Capacity sheet is primary after first generation
EXTERNAL_COMMITMENTS = [
    {"person": "Alice", "project": "ProjectX", "fte": 2.0, "start": 1, "end": 11},
    {"person": "Bob",   "project": "ProjectX", "fte": 2.0, "start": 1, "end": 60},
]

# ── LAYOUT CONSTANTS ──
# Column where month data starts (C = 3). A = WP ID, B = WP Name.
MONTH_COL_START = 3

# ── DERIVED (do not edit) ──
NUM_WPS = len(WORK_PACKAGES)
MONTH_COL_END = MONTH_COL_START + PROJECT_MONTHS - 1
VAL_COL_SUM = MONTH_COL_END + 1
VAL_COL_BUDGET = MONTH_COL_END + 2
VAL_COL_DIFF = MONTH_COL_END + 3
