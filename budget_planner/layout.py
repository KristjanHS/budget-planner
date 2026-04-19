"""Centralized layout constants — single source of truth for column structure.

Every builder imports from here instead of defining its own ``month_col_start``.
Validation-column offsets are relative to the last month column so that adding
or removing months never requires touching builder arithmetic.
"""

from typing import TypedDict

from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from dateutil.relativedelta import relativedelta

from budget_planner.styles import (
    style_cell,
    HEADER_FONT, LIGHT_BLUE_FILL, CENTER,
)


# Config!B-column parameter rows (Section 1 of Config sheet). Keys mirror the
# literal label strings written into col A — using a TypedDict here surfaces
# typos in consumer lookups at pyright time instead of at runtime.
ConfigParamRows = TypedDict("ConfigParamRows", {
    "Project Title": int,
    "Project Start": int,
    "Project Months": int,
    "FTE Rate (EUR)": int,
    "EU Annual Ceiling (PM/FTE/year)": int,
    "Hours per FTE Month": int,
    "Vacation Month": int,
    "Num Vacation Months": int,
    "Version Stamp": int,
})

# ── Sheet name constants ──
SHEET_CONFIG = "Config"
SHEET_PLAN = "Plan"
SHEET_CAPACITY = "Capacity"
SHEET_ACTUALS = "Actuals"
SHEET_ACTUALS_HOURS = "Actuals-Hours"
SHEET_PLAN_FORECAST = "Plan-Forecast"
SHEET_SUMMARY = "Summary"
SHEET_TRENDS = "Trends"
SHEET_TIMELOG_PLAN = "Timelog Plan"

# Column structure: A=ID, B=Name, C onwards = months
MONTH_COL_START = 3

# ── Timelog Plan column structure (merged-grid layout) ──
# A = Person (cross-ref to Team Roster)
# B..B+months-1 = month grid — single editable grid (v2 layout collapsed the
#     v1 PROPOSED + ASSIGNED twin bands into one). On rebuild, re-solve or
#     preserve is governed by the `timelog_resolve_flag` toggle cell.
# PER-PERSON SUMMARY band reuses col A for Person; metrics live at B..H
# (Assigned, Pool, Over-pool?, Switches, First log, Last log, Rate). Rate is
# duplicated only once per person — on the summary row.
TIMELOG_MONTH_COL_START = 2   # col B

# ── Section header contract ──
# Every labelled section on the Config sheet uses: title row, header row,
# then data starting at offset +2. Readers rely on this exact layout.
SECTION_HEADER_ROWS = 2

SECTION_HEADER_LABELS = {
    "wp_budget": "WP BUDGET",
    "wp_schedule": "WP SCHEDULE",
    "team_roster": "TEAM ROSTER",
    "rate_by_grade": "RATE BY GRADE",
    "reporting_periods": "REPORTING PERIODS",
    "assignments": "PERSON-WP ASSIGNMENTS",
    "externals": "EXTERNAL COMMITMENTS",
    "planned_fte": "PLANNED FTE (PERSON-MONTHS)",
    "wp_id_column": "WP ID",  # column-header marker in Capacity/Actuals month header row
}

# ── Config sheet column constants ──
# WP Budget section (Config sheet)
COL_WP_ID = 1
COL_WP_NAME = 2
COL_WP_FTE = 3
COL_WP_RATE = 4
COL_WP_CEILING = 5
COL_WP_ASSIGNED_PM = 6  # SUMIFS formula: total assigned PM per WP
COL_WP_VARIANCE = 7     # FTE Budget - Assigned PM
WP_BUDGET_LAST_COL = 7  # for table ref / section header merge

# WP Schedule section (Config sheet)
COL_SCHED_START = 2
COL_SCHED_END = 3
COL_SCHED_TYPE = 4
SCHED_LAST_COL = 4

# Team Roster section (Config sheet)
COL_ROSTER_FTE = 2
COL_ROSTER_GRADE = 3
COL_ROSTER_RATE = 4
COL_ROSTER_PRIORITY = 5
ROSTER_LAST_COL = 5

# Reporting Periods section (Config sheet)
COL_RP_START = 2
COL_RP_END = 3
RP_LAST_COL = 3

# Rate by Grade section (Config sheet)
COL_GRADE_RATE = 2
GRADE_LAST_COL = 2

# Person-WP Assignments (Config sheet)
COL_ASSIGN_PERSON = 1
COL_ASSIGN_WP = 2
COL_ASSIGN_FTE = 3
COL_ASSIGN_START = 4
COL_ASSIGN_END = 5
COL_ASSIGN_SIG_ASSIGNED = 6  # SUMIFS: total assigned FTE for this person
COL_ASSIGN_SIG_EXTERNAL = 7  # SUMIFS: total external FTE for this person
COL_ASSIGN_BALANCE = 8       # Available PM - Assigned - External
ASSIGN_LAST_COL = 8
# Minimum number of rows in the assignments section (populated + blank template
# rows). Ensures named ranges span a realistic area even when assignments=[],
# and avoids the "partial merge" error when users insert rows adjacent to the
# EXTERNAL COMMITMENTS section header.
ASSIGN_MIN_ROWS = 5

# External Commitments (Config sheet)
COL_EXT_PERSON = 1
COL_EXT_PROJECT = 2
COL_EXT_FTE = 3
COL_EXT_START = 4
COL_EXT_END = 5
COL_EXT_EU = 6
EXT_LAST_COL = 6

# Validation column offsets (relative to month_col_end)
VAL_OFFSET_SUM = 1
VAL_OFFSET_BUDGET = 2
VAL_OFFSET_DIFF = 3
# Actuals-only additional columns
VAL_OFFSET_CUM = 4
VAL_OFFSET_PCT_CEIL = 5
VAL_OFFSET_BREACH = 6


# Alert thresholds
TREND_OK_BAND = 0.05
TREND_OVER_THRESHOLD = 0.15
TREND_WATCH_THRESHOLD = 0.05   # same as OK_BAND, used for WATCH vs OK boundary
TREND_SLOW_THRESHOLD = -0.10
CAPACITY_AMBER_LOAD = 0.80
CAPACITY_VARIANCE_FTE_RED = 0.2  # |Planned FTE - Actual FTE| > this → red
ACTUALS_AMBER_PCT_CEIL = 0.85
ACTUALS_AMBER_BREACH_MONTHS = 12
ACTUALS_RED_BREACH_MONTHS = 6


def cell_ref(sheet: str, row: int, col: int, *, absolute: bool = True) -> str:
    """Return an Excel cell reference like ``Sheet!$B$3`` from numeric coordinates.

    Prefer this over hand-writing ``f"{SHEET}!$B${row}"`` so that layout column
    changes in this module propagate to every cross-sheet formula automatically.
    """
    letter = get_column_letter(col)
    if absolute:
        return f"{sheet}!${letter}${row}"
    return f"{sheet}!{letter}{row}"


def month_range_ref(row: int, start_col: int, end_col: int) -> str:
    """Return ``"A5:Z5"`` for a horizontal range on one row.

    Use for SUM-across-months formulas and range-scoped CF targets to avoid
    hand-typing matched column-letter pairs.
    """
    return f"{get_column_letter(start_col)}{row}:{get_column_letter(end_col)}{row}"


def month_col_end(project_months: int) -> int:
    """Last month column number."""
    return MONTH_COL_START + project_months - 1


def val_columns(project_months: int) -> dict[str, int]:
    """Return dict of validation column numbers."""
    end = month_col_end(project_months)
    return {
        "sum": end + VAL_OFFSET_SUM,
        "budget": end + VAL_OFFSET_BUDGET,
        "diff": end + VAL_OFFSET_DIFF,
        "cum": end + VAL_OFFSET_CUM,
        "pct_ceil": end + VAL_OFFSET_PCT_CEIL,
        "breach": end + VAL_OFFSET_BREACH,
    }


# ── Forecast monitoring column offsets (Actuals sheet) ──
# Placed after the Actual-EUR validation block (``val_columns``) with a
# 1-column gutter so the groupable forecast band is visually separable.
FORECAST_GUTTER = 1
FORECAST_OFFSET_CUM = FORECAST_GUTTER + 1
FORECAST_OFFSET_VAR = FORECAST_GUTTER + 2
FORECAST_OFFSET_PCT_CEIL = FORECAST_GUTTER + 3
FORECAST_OFFSET_BREACH = FORECAST_GUTTER + 4


def forecast_val_columns(project_months: int) -> dict[str, int]:
    """Return dict of forecast-monitoring column numbers on Actuals.

    Positioned after the Actual-EUR ``val_columns`` breach column, with a
    1-column gutter. The four columns mirror the Actual-EUR shape:
    ``cum_forecast``, ``var_forecast``, ``pct_ceil_forecast``,
    ``breach_forecast``.
    """
    breach = val_columns(project_months)["breach"]
    return {
        "cum_forecast": breach + FORECAST_OFFSET_CUM,
        "var_forecast": breach + FORECAST_OFFSET_VAR,
        "pct_ceil_forecast": breach + FORECAST_OFFSET_PCT_CEIL,
        "breach_forecast": breach + FORECAST_OFFSET_BREACH,
    }


def wp_active_months(wp_id: str, schedule: list[dict]) -> frozenset[int]:
    """Return frozenset of active month numbers for a WP from its schedule entries."""
    return frozenset(
        m
        for entry in schedule if entry["id"] == wp_id
        for m in range(entry["start"], entry["end"] + 1)
    )


def wp_month_bounds(wp_id: str, schedule: list[dict]) -> tuple[int, int]:
    """Return (min_start, max_end) across all schedule entries for a WP."""
    entries = [e for e in schedule if e["id"] == wp_id]
    if not entries:
        return (1, 1)
    return (min(e["start"] for e in entries), max(e["end"] for e in entries))


def _calendar_year_groups(project_start, project_months):
    """Yield (group_start_month_index, span) tuples for calendar-year boundaries.

    For a project starting Nov 2026 with 14 months:
      2026: M1-M2 (indices 0-1, span=2)
      2027: M3-M14 (indices 2-13, span=12)
    """
    m = 0
    while m < project_months:
        cal_date = project_start + relativedelta(months=m)
        # Months remaining in this calendar year
        months_left_in_year = 12 - cal_date.month + 1
        span = min(months_left_in_year, project_months - m)
        yield m, span
        m += span


def build_year_month_headers(ws, start_row, project_params, month_col_start,
                             val_headers=None):
    """Write year + month header rows. Returns (year_row, month_row, next_row).

    Args:
        ws: worksheet
        start_row: row number to start at
        project_params: dict with 'months' and 'start' keys
        month_col_start: first month column number
        val_headers: optional list of (col, label) tuples for validation columns.
            Each tuple writes the label in the year row and an empty
            LIGHT_BLUE_FILL cell in the month row at the given column.
    """
    if val_headers is None:
        val_headers = []

    r = start_row
    project_months = project_params["months"]
    project_start = project_params["start"]

    # ── Year header row (calendar-year boundaries) ──
    year_row = r
    style_cell(ws, r, 1, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 2, "", fill=LIGHT_BLUE_FILL)
    for group_start, span in _calendar_year_groups(project_start, project_months):
        year_start_col = month_col_start + group_start
        year_end_col = year_start_col + span - 1
        cal_date = project_start + relativedelta(months=group_start)
        year_label = str(cal_date.year)
        if span > 1:
            ws.merge_cells(start_row=r, start_column=year_start_col,
                           end_row=r, end_column=year_end_col)
        style_cell(ws, r, year_start_col, year_label, font=HEADER_FONT,
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)
        for c2 in range(year_start_col + 1, year_end_col + 1):
            style_cell(ws, r, c2, fill=LIGHT_BLUE_FILL)
    for col, label in val_headers:
        style_cell(ws, r, col, label, font=HEADER_FONT,
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)
    r += 1

    # ── Month header row ──
    month_row = r
    style_cell(ws, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    style_cell(ws, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    for mi in range(project_months):
        c = month_col_start + mi
        cal_date = project_start + relativedelta(months=mi)
        mlabel = f"M{mi + 1}-{cal_date.strftime('%b%y')}"
        style_cell(ws, r, c, mlabel, font=Font(name="Arial", size=8),
                   fill=LIGHT_BLUE_FILL, alignment=CENTER)
    for col, _label in val_headers:
        style_cell(ws, r, col, "", fill=LIGHT_BLUE_FILL)
    r += 1

    return year_row, month_row, r
