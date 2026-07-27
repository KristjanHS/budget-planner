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
    HEADER_FONT,
    LIGHT_BLUE_FILL,
    CENTER,
)


# Config!B-column parameter rows (Section 1 of Config sheet). Keys mirror the
# literal label strings written into col A — using a TypedDict here surfaces
# typos in consumer lookups at pyright time instead of at runtime.
ConfigParamRows = TypedDict(
    "ConfigParamRows",
    {
        "Project Title": int,
        "Project Start": int,
        "Project Months": int,
        "FTE Rate (EUR)": int,
        "EU Annual Ceiling (PM/FTE/year)": int,
        "Hours per FTE Month": int,
        "Vacation Month": int,
        "Num Vacation Months": int,
        "Rebuild from month #": int,
        "Version Stamp": int,
    },
)

# ── Sheet name constants ──
SHEET_CONFIG = "Config"
SHEET_PLAN = "Plan"
SHEET_CAPACITY = "Capacity"
SHEET_ACTUALS = "Actuals"
SHEET_ACTUALS_HOURS = "Actuals-Hours"
SHEET_SUMMARY = "Summary"
SHEET_TRENDS = "Trends"
SHEET_TIMELOG_PLAN = "Timelog Plan"

# Column structure: A=ID, B=Name, C onwards = months
MONTH_COL_START = 3

# ── Timelog Plan column structure (merged-grid layout) ──
# A = Person (cross-ref to Team Roster)
# B..B+months-1 = month grid — single editable grid (v2 layout collapsed the
#     v1 PROPOSED + ASSIGNED twin bands into one). On rebuild, the freeze
#     point is governed by the `timelog_rebuild_from_month` integer Config
#     param (Config!B9): months < N frozen verbatim, months ≥ N re-solved.
# PER-PERSON SUMMARY band reuses col A for Person; metrics live at B..H
# (Assigned, Pool, Over-pool?, Switches, First log, Last log, Rate). Rate is
# duplicated only once per person — on the summary row.
TIMELOG_MONTH_COL_START = 2  # col B

# PER-WP SUMMARY base-column span: A..J (10 cols: WP code + 9 metrics).
# Year cols (``Y{yyyy} +i``) start at TIMELOG_WP_SUMMARY_BASE_LAST_COL + 1 = K.
# Aiko-port T3 (2026-04-23) inserted Buffer % at col E; cols E→F, F→G, G→H, H→I
# shifted right by one. 2026-04-26 appended "Stop reason" at col J; year band
# shifted J..N → K..O.
TIMELOG_WP_SUMMARY_BASE_LAST_COL = 10

# ``Yearly Assigned +i`` / ``Yearly Budget +i`` row labels sit at col J,
# stacked beneath the "Stop reason" header col, immediately left of the
# per-year breakdown cols (K..) and right-aligned so the label hugs the
# year value it describes. Shift from col A decided in plan D9 (2026-04-22);
# post-Aiko-port T3 shifted H → I when Buffer % was inserted; 2026-04-28
# shifted I → J under "Stop reason" with right alignment.
TIMELOG_YEARLY_LABEL_COL = 10

# Timelog Plan grid section title — written by the builder into col A and used
# by the reader as the row anchor for the merged grid band. Lives here (not in
# timelog_plan) so the reader doesn't depend on a sheet builder for layout.
SECTION_TITLE_GRID = "TIMELOG GRID (edit freely — rebuild re-solves from Config!B9 onward)"

# ── Section header contract ──
# Every labelled section on the Config sheet uses: title row, header row,
# then data starting at offset +2. Readers rely on this exact layout.
SECTION_HEADER_ROWS = 2

SECTION_HEADER_LABELS = {
    "wp_budget": "WP BUDGET",
    "team_roster": "TEAM ROSTER",
    "rate_by_grade": "RATE BY GRADE",
    "reporting_periods": "REPORTING PERIODS",
    "assignments": "PERSON-WP ASSIGNMENTS",
    "externals": "EXTERNAL COMMITMENTS",
    "planned_fte": "PLANNED FTE (PERSON-MONTHS)",
    "stop_reasons": "TIMELOG SOLVER — STOP REASONS",
    "wp_id_column": "WP ID",  # column-header marker in Capacity/Actuals month header row
}

# ── Config sheet column constants ──
# WP Budget section (Config sheet). Indirect-view layout:
#   Col 5 = Budget -i (black-formula = Budget +i / 1.25, without 25% indirect).
#   Col 6 = Budget +i (black-formula, Aiko-port T1: structured-ref
#           =tbl_WorkPackages[[#This Row],[FTE Budget]]*...[Rate (€/PM) -i]*1.25;
#           recalculates on edit — no user-typed literal).
# Col 5's numeric value matches the legacy ``Ceiling = FTE * Rate`` output so
# deferred consumers (Summary/Trends/Plan) that read by position keep working.
COL_WP_ID = 1
COL_WP_NAME = 2
COL_WP_FTE = 3
COL_WP_RATE = 4
COL_WP_BUDGET_MINUS_I = 5  # black-formula = Budget +i / 1.25 (orange header)
COL_WP_BUDGET_PLUS_I = 6  # black-formula (structured-ref, Aiko-port T1); green header
COL_WP_ASSIGNED_PM = 7  # SUMIFS formula: total assigned PM per WP
COL_WP_VARIANCE = 8  # FTE Budget - Assigned PM
COL_WP_START = 9  # project-month start (blue input, 1-indexed)
COL_WP_END = 10  # project-month end (blue input, 1-indexed)
COL_WP_TYPE = 11  # distribution curve (blue input, data-validation list)
WP_BUDGET_LAST_COL = 11  # for table ref / section header merge

# Compat alias for out-of-T1-scope consumers (actuals.py, tests/test_layout.py)
# that still reference the legacy ``COL_WP_CEILING`` name. Resolves to the same
# column as ``COL_WP_BUDGET_MINUS_I``. Remove when all sheets migrate.
COL_WP_CEILING = COL_WP_BUDGET_MINUS_I

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
COL_ASSIGN_BALANCE = 8  # Available PM - Assigned - External
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

# Validation column offsets (relative to month_col_end).
# Offsets 1..6 carry the without-indirect (-i) rollups; 7..12 carry the
# with-indirect (+i) mirrors. The +i block is outside the collapsible -i
# band and forms the default-visible rollup surface for external-facing
# use. See ``docs/spec.md`` § Indirect-view EUR header suffix.
#
# Actuals semantic: offset 1 = Slope (avg-burn = posted_value/posted_idx);
# offset 4 = Posted (latest-filled cumulative through current_month).
# Posted-cumulative redesign (2026-05-09) replaced Sum/Cum on the Actuals
# sheet — see ``docs/plans/archive/2026-05-09-actuals-posted-cumulative-design.md``.
VAL_OFFSET_SLOPE = 1
VAL_OFFSET_BUDGET = 2
VAL_OFFSET_DIFF = 3
VAL_OFFSET_POSTED = 4
VAL_OFFSET_PCT_CEIL = 5
VAL_OFFSET_BREACH = 6
VAL_OFFSET_SLOPE_PLUS_I = 7
VAL_OFFSET_BUDGET_PLUS_I = 8
VAL_OFFSET_DIFF_PLUS_I = 9
VAL_OFFSET_POSTED_PLUS_I = 10
VAL_OFFSET_PCT_CEIL_PLUS_I = 11
VAL_OFFSET_BREACH_PLUS_I = 12

# Hidden helpers (Actuals-only; -i side; outline-grouped, black-formula).
# `posted_idx` = banned-pattern-safe SUMPRODUCT-MAX scan returning the latest
# 1-based month index with a numeric posting; `posted_value` = INDEX into
# row at that index. Consumed by the visible Posted/Slope/Variance/%/Breach
# formulas. +i side does NOT mirror — Posted +i = ROUND(-i × 1.25, 0).
VAL_OFFSET_POSTED_IDX = 13
VAL_OFFSET_POSTED_VALUE = 14

# Cross-sheet column-position aliases. The same physical column slots host
# different semantic content per sheet: Actuals uses Slope / Posted (per the
# 2026-05-09 posted-cumulative redesign), while Plan / Capacity / Actuals-
# Hours / Summary use the slot for row-wide monthly-SUM. Both names are
# canonical in their own sheet's context — these aliases let each builder
# read the slot under the name that matches its semantics.
VAL_OFFSET_SUM = VAL_OFFSET_SLOPE
VAL_OFFSET_SUM_PLUS_I = VAL_OFFSET_SLOPE_PLUS_I
VAL_OFFSET_CUM = VAL_OFFSET_POSTED
VAL_OFFSET_CUM_PLUS_I = VAL_OFFSET_POSTED_PLUS_I


# Alert thresholds
TREND_OK_BAND = 0.05
TREND_OVER_THRESHOLD = 0.15
TREND_WATCH_THRESHOLD = 0.05  # same as OK_BAND, used for WATCH vs OK boundary
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
    """Return dict of validation column numbers.

    Primary keys (-i side): ``slope`` / ``budget`` / ``diff`` / ``posted`` /
    ``pct_ceil`` / ``breach`` at offsets 1..6 from ``month_col_end``.
    +i mirrors at offsets 7..12 (``*_plus_i``). Hidden helpers
    ``posted_idx`` / ``posted_value`` at offsets 13..14 (Actuals only).

    Cross-semantic keys ``sum`` / ``cum`` / ``sum_plus_i`` / ``cum_plus_i``
    resolve to the same column slots as ``slope`` / ``posted`` (offsets
    1 / 4 / 7 / 10). Each sheet uses the name that matches its semantic:
    Actuals reads as ``slope`` / ``posted``, while Plan / Capacity /
    Actuals-Hours / Summary read as ``sum`` (row-wide monthly-SUM).
    """
    end = month_col_end(project_months)
    return {
        # Posted-cumulative redesign (current naming):
        "slope": end + VAL_OFFSET_SLOPE,
        "budget": end + VAL_OFFSET_BUDGET,
        "diff": end + VAL_OFFSET_DIFF,
        "posted": end + VAL_OFFSET_POSTED,
        "pct_ceil": end + VAL_OFFSET_PCT_CEIL,
        "breach": end + VAL_OFFSET_BREACH,
        "slope_plus_i": end + VAL_OFFSET_SLOPE_PLUS_I,
        "budget_plus_i": end + VAL_OFFSET_BUDGET_PLUS_I,
        "diff_plus_i": end + VAL_OFFSET_DIFF_PLUS_I,
        "posted_plus_i": end + VAL_OFFSET_POSTED_PLUS_I,
        "pct_ceil_plus_i": end + VAL_OFFSET_PCT_CEIL_PLUS_I,
        "breach_plus_i": end + VAL_OFFSET_BREACH_PLUS_I,
        "posted_idx": end + VAL_OFFSET_POSTED_IDX,
        "posted_value": end + VAL_OFFSET_POSTED_VALUE,
        # Cross-semantic aliases — non-Actuals sheets use sum/cum names
        # for the same physical column slots (see VAL_OFFSET_SUM/CUM).
        "sum": end + VAL_OFFSET_SUM,
        "sum_plus_i": end + VAL_OFFSET_SUM_PLUS_I,
        "cum": end + VAL_OFFSET_CUM,
        "cum_plus_i": end + VAL_OFFSET_CUM_PLUS_I,
    }


def wp_active_months(wp_id: str, work_packages: list[dict]) -> frozenset[int]:
    """Return frozenset of active month numbers for a WP.

    Single-range schema: exactly one entry per id in work_packages.
    """
    for wp in work_packages:
        if wp["id"] == wp_id:
            return frozenset(range(wp["start"], wp["end"] + 1))
    return frozenset()


def wp_month_bounds(wp_id: str, work_packages: list[dict]) -> tuple[int, int]:
    """Return (start, end) for a WP. Single-range schema: one entry per id."""
    for wp in work_packages:
        if wp["id"] == wp_id:
            return (int(wp["start"]), int(wp["end"]))
    return (1, 1)


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


def build_year_month_headers(ws, start_row, project_params, month_col_start, val_headers=None):
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
            ws.merge_cells(start_row=r, start_column=year_start_col, end_row=r, end_column=year_end_col)
        style_cell(ws, r, year_start_col, year_label, font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
        for c2 in range(year_start_col + 1, year_end_col + 1):
            style_cell(ws, r, c2, fill=LIGHT_BLUE_FILL)
    for col, label in val_headers:
        style_cell(ws, r, col, label, font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    r += 1

    # ── Month header row ──
    month_row = r
    style_cell(ws, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    for mi in range(project_months):
        c = month_col_start + mi
        cal_date = project_start + relativedelta(months=mi)
        mlabel = f"M{mi + 1}-{cal_date.strftime('%b%y')}"
        style_cell(ws, r, c, mlabel, font=Font(name="Arial", size=8), fill=LIGHT_BLUE_FILL, alignment=CENTER)
    for col, _label in val_headers:
        style_cell(ws, r, col, "", fill=LIGHT_BLUE_FILL)
    r += 1

    return year_row, month_row, r
