"""End-to-end: hours -> forecast EUR chain through rebuild + recalc.

Closes the loop introduced by Tasks 3-13 of the hours-timelog v1 plan:

    init (demo)
       -> write hour inputs into 3 person/WP/month cells
       -> edit a grade rate in Config
       -> rebuild_workbook (preserves hours; recomputes derived sheets)
       -> scripts/recalc.py via subprocess (LO recalculates formulas)
       -> reopen with data_only=True

Assertions:
    1. Hour inputs survive the rebuild (raw blue inputs preserved).
    2. Each forecast-EUR leaf == ROUND(hours / hours_per_fte_month * NEW_RATE, 0).
    3. Plan-Forecast (WP, month) cell sums the leaves correctly (here: a
       single leaf per write, so the cell value == that leaf).
    4. Actuals "Cum. Forecast EUR" TOTAL row equals the sum of the three
       leaves (all writes are at months <= current_month, so the
       SUMPRODUCT mask includes them).
    5. Summary "Fcst €" column for the year covering the writes includes
       the three leaves' total (writes are colocated in one calendar year).

LibreOffice is required for recalc; the test skips if recalc.py reports
the binary is unavailable, mirroring tests/test_rebuild_e2e.py.
"""
from __future__ import annotations

import json
import re
import subprocess
from datetime import date
from pathlib import Path

import openpyxl
import pytest
from dateutil.relativedelta import relativedelta
from openpyxl.cell import Cell

from budget_planner.config import (
    DEMO_CAPACITY_ASSIGNMENTS,
    GRADE_RATES,
    PROJECT_MONTHS,
    PROJECT_START,
)
from budget_planner.layout import (
    MONTH_COL_START,
    SHEET_ACTUALS,
    SHEET_ACTUALS_HOURS,
    SHEET_CONFIG,
    SHEET_PLAN_FORECAST,
    SHEET_SUMMARY,
    forecast_val_columns,
    val_columns,
)
from budget_planner.reader import _find_header_row
from budget_planner.workbook import init_workbook, rebuild_workbook


SCRIPTS_DIR = Path(__file__).parent.parent / "scripts"
HOURS_PER_FTE_MONTH = 143  # default seeded by config_sheet.py
NEW_SENIOR_RATE = 12000    # was 10000; Alice is "Senior"
TARGET_GRADE = "Senior"


# ── Helpers ─────────────────────────────────────────────────────────────────


def _set(ws, row: int, col: int, value) -> None:
    cell = ws.cell(row, col)
    assert isinstance(cell, Cell)
    cell.value = value


def _find_person_header_row(ws, person_name: str) -> int:
    """Locate the merged-A:B person banner row on Actuals-Hours."""
    merged = [
        (mr.min_row, mr.min_col, mr.max_col)
        for mr in ws.merged_cells.ranges
    ]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value != person_name:
            continue
        if any(
            mr_r == r and mr_c1 == 1 and mr_c2 >= 2
            for mr_r, mr_c1, mr_c2 in merged
        ):
            return r
    raise AssertionError(f"Person header {person_name!r} not found")


def _find_wp_row_under_person(ws, header_row: int, wp_id: str) -> int:
    """Walk rows below a person header; match column A to ``wp_id``.

    Works for both formula-loaded workbooks (col A = ``=Config!A<n>``) and
    data_only-loaded workbooks (col A = the resolved WP id literal).
    """
    config_ws = ws.parent[SHEET_CONFIG]
    for r in range(header_row + 1, ws.max_row + 1):
        b_val = ws.cell(r, 2).value
        if isinstance(b_val, str) and b_val.upper() in ("TOTAL", "EU TOTAL"):
            raise AssertionError(
                f"WP row for {wp_id!r} not found before TOTAL row"
            )
        val = ws.cell(r, 1).value
        if val is None:
            continue
        if isinstance(val, str):
            m = re.match(r"=Config!A(\d+)", val)
            if m:
                resolved = config_ws.cell(int(m.group(1)), 1).value
                if resolved and str(resolved).strip() == wp_id:
                    return r
                continue
            if val.strip() == wp_id:
                return r
    raise AssertionError(f"WP row for {wp_id!r} under row {header_row} not found")


def _find_grade_rate_row(config_ws, grade_name: str) -> int:
    """Find the RATE BY GRADE data row for a given grade name."""
    section_row = _find_header_row(config_ws, "RATE BY GRADE")
    r = section_row + 2  # skip section header + column header
    while r <= config_ws.max_row:
        val = config_ws.cell(r, 1).value
        if val is None or str(val).strip() == "":
            break
        if str(val).strip() == grade_name:
            return r
        r += 1
    raise AssertionError(f"Grade {grade_name!r} not found in RATE BY GRADE")


def _find_wp_row_in_sheet(ws, wp_id: str, *, after_row: int = 1) -> int:
    """Find the WP row in a sheet by resolving =Config!A<n> in column A."""
    config_ws = ws.parent[SHEET_CONFIG]
    for r in range(after_row, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if isinstance(val, str):
            m = re.match(r"=Config!A(\d+)", val)
            if m:
                resolved = config_ws.cell(int(m.group(1)), 1).value
                if resolved and str(resolved).strip() == wp_id:
                    return r
    raise AssertionError(f"WP row for {wp_id!r} not found in sheet")


def _find_wp_row_data_only(ws, wp_id: str) -> int:
    """Find the WP row in a data_only-loaded sheet (col A holds resolved id)."""
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val is not None and str(val).strip() == wp_id:
            return r
    raise AssertionError(f"WP row for {wp_id!r} not found in sheet (data_only)")


def _find_actuals_total_row(ws) -> int:
    """Locate the TOTAL row in Actuals (col B == 'TOTAL')."""
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 2).value
        if isinstance(v, str) and v.strip().upper() == "TOTAL":
            return r
    raise AssertionError("TOTAL row not found on Actuals")


def _summary_first_wp_row(ws) -> int:
    """First WP data row in Summary == first row whose col A is =Config!A<n>."""
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith("=Config!A"):
            return r
    raise AssertionError("No WP data row found on Summary")


def _summary_first_wp_row_data_only(ws, wp_id: str) -> int:
    """First WP data row in data_only-loaded Summary, located by WP id."""
    # Summary has WP rows in BUDGET section and again in RP SUMMARY section.
    # We want the first occurrence (the year-by-year table).
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is not None and str(v).strip() == wp_id:
            return r
    raise AssertionError(f"WP row for {wp_id!r} not found on Summary (data_only)")


def _summary_year_columns_for(ws, year_label: int) -> tuple[int, int]:
    """Return (start_col, end_col) of the merged year-group header in Summary.

    Year header is on row 3 (annual budget table) — the row above the
    WP/Name/Plan-FTE sub-header row. Identified by merged ranges spanning
    the COLS_PER_YEAR=5 columns of that calendar year.
    """
    # Scan merged cell ranges for a cell whose value matches the year label.
    for mr in ws.merged_cells.ranges:
        if mr.min_row != mr.max_row:
            continue
        v = ws.cell(mr.min_row, mr.min_col).value
        if v is None:
            continue
        if str(v).strip() == str(year_label):
            return mr.min_col, mr.max_col
    raise AssertionError(f"Year {year_label} merged header not found on Summary")


# ── The test ────────────────────────────────────────────────────────────────


def test_hours_propagate_through_rebuild_and_recalc(tmp_path):
    """init -> write hours -> edit grade rate -> rebuild -> recalc -> assert."""
    path = str(tmp_path / "wb.xlsx")
    init_workbook(path, assignments=DEMO_CAPACITY_ASSIGNMENTS)

    # --- Pick 3 (person, WP, month) cells. Alice is grade Senior in the
    #     demo roster, assigned to WP1, WP2, WP3. We derive months from
    #     current_month (the rebuild code's own formula) so the cumulative
    #     mask catches them regardless of when the test is run. All three
    #     months must fall in the same calendar year for assertion 5. ---
    _pstart = PROJECT_START.date() if hasattr(PROJECT_START, "date") else PROJECT_START
    _delta = relativedelta(date.today(), _pstart)
    current_month = max(
        1, min(_delta.years * 12 + _delta.months + 1, PROJECT_MONTHS)
    )
    if current_month < 3:
        pytest.skip(
            f"current_month={current_month} too early in project to pick 3 "
            f"distinct logged months within the cumulative-forecast mask"
        )
    # Three consecutive months ending at current_month, all <= current_month.
    m_hi = current_month
    m_mid = current_month - 1
    m_lo = current_month - 2
    # Require all three to land in the same calendar year so the Summary
    # assertion stays a single-year aggregation.
    def _year_of(m: int) -> int:
        return (_pstart + relativedelta(months=m - 1)).year
    target_year = _year_of(m_hi)
    if _year_of(m_lo) != target_year:
        # Shift the window up into the next year's first months.
        # (Happens when current_month straddles a Dec→Jan boundary.)
        while _year_of(m_lo) != target_year and m_hi < PROJECT_MONTHS:
            m_hi += 1
            m_mid = m_hi - 1
            m_lo = m_hi - 2
            target_year = _year_of(m_hi)
        if _year_of(m_lo) != target_year or m_hi > current_month:
            pytest.skip(
                f"cannot pick 3 same-year months <= current_month={current_month}"
            )
    edits: list[tuple[str, str, int, float]] = [
        # (person, wp_id, month_num (1-based), hours)
        ("Alice", "WP1", m_lo,  100.0),
        ("Alice", "WP2", m_mid,  80.0),
        ("Alice", "WP3", m_hi,   60.0),
    ]

    # Sanity: target grade and assignments line up with demo data.
    assert GRADE_RATES[TARGET_GRADE] != NEW_SENIOR_RATE, (
        "NEW_SENIOR_RATE must differ from default to prove the rate edit took"
    )
    for person, wp_id, _m, _h in edits:
        assert any(
            a["person"] == person and a["wp"] == wp_id
            for a in DEMO_CAPACITY_ASSIGNMENTS
        ), f"demo data must assign {person} to {wp_id}"

    # --- Write hour inputs + grade rate edit ---
    wb = openpyxl.load_workbook(path)
    ah = wb[SHEET_ACTUALS_HOURS]
    for person, wp_id, month_num, hours in edits:
        hdr = _find_person_header_row(ah, person)
        wp_row = _find_wp_row_under_person(ah, hdr, wp_id)
        _set(ah, wp_row, MONTH_COL_START + (month_num - 1), hours)

    cfg = wb[SHEET_CONFIG]
    grade_row = _find_grade_rate_row(cfg, TARGET_GRADE)
    _set(cfg, grade_row, 2, NEW_SENIOR_RATE)
    wb.save(path)
    wb.close()

    # --- Rebuild ---
    rebuild_workbook(input_path=path)

    # --- Recalc via LibreOffice ---
    result = subprocess.run(
        ["python3", str(SCRIPTS_DIR / "recalc.py"), path],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        pytest.skip(f"LibreOffice recalc unavailable: {result.stderr[:200]}")
    recalc_data = json.loads(result.stdout)
    assert recalc_data["status"] == "success", recalc_data
    assert recalc_data["total_errors"] == 0, recalc_data

    # ── Assertion 1: Hour inputs preserved (raw values, formulas wb) ──
    wb = openpyxl.load_workbook(path)
    ah = wb[SHEET_ACTUALS_HOURS]
    for person, wp_id, month_num, hours in edits:
        hdr = _find_person_header_row(ah, person)
        wp_row = _find_wp_row_under_person(ah, hdr, wp_id)
        got = ah.cell(wp_row, MONTH_COL_START + (month_num - 1)).value
        assert got == hours, (
            f"hour input lost: {person}/{wp_id}/M{month_num}: "
            f"expected {hours}, got {got!r}"
        )

    # Also confirm the Config grade rate edit survived.
    cfg = wb[SHEET_CONFIG]
    grade_row = _find_grade_rate_row(cfg, TARGET_GRADE)
    assert cfg.cell(grade_row, 2).value == NEW_SENIOR_RATE
    wb.close()

    # ── Reload data_only=True for computed values ──
    wb = openpyxl.load_workbook(path, data_only=True)
    ah = wb[SHEET_ACTUALS_HOURS]

    # Forecast band starts after the Sum column + 1-col gutter.
    forecast_col_start = val_columns(PROJECT_MONTHS)["sum"] + 2

    # ── Assertion 2: Forecast-EUR leaf == ROUND(hours/143 * NEW_RATE, 0) ──
    expected_leaves: dict[tuple[str, str, int], int] = {}
    for person, wp_id, month_num, hours in edits:
        hdr = _find_person_header_row(ah, person)
        wp_row = _find_wp_row_under_person(ah, hdr, wp_id)
        leaf_col = forecast_col_start + (month_num - 1)
        leaf_val = ah.cell(wp_row, leaf_col).value
        expected = round(hours / HOURS_PER_FTE_MONTH * NEW_SENIOR_RATE)
        assert leaf_val == expected, (
            f"forecast leaf mismatch {person}/{wp_id}/M{month_num}: "
            f"expected {expected}, got {leaf_val!r}"
        )
        expected_leaves[(person, wp_id, month_num)] = expected

    sum_of_leaves = sum(expected_leaves.values())

    # ── Assertion 3: Plan-Forecast (WP, month) cell == leaf sum for that
    #    cell (here: 1 leaf per cell, so cell == single leaf). ──
    pf = wb[SHEET_PLAN_FORECAST]
    for (person, wp_id, month_num), expected in expected_leaves.items():
        del person  # Plan-Forecast aggregates across persons; only 1 leaf here
        pf_row = _find_wp_row_data_only(pf, wp_id)
        pf_col = MONTH_COL_START + (month_num - 1)
        cell_val = pf.cell(pf_row, pf_col).value
        assert cell_val == expected, (
            f"Plan-Forecast {wp_id}/M{month_num}: expected {expected}, "
            f"got {cell_val!r}"
        )

    # ── Assertion 4: Actuals Cum. Forecast EUR (TOTAL row) == sum of leaves ──
    actuals = wb[SHEET_ACTUALS]
    total_row = _find_actuals_total_row(actuals)
    cum_forecast_col = forecast_val_columns(PROJECT_MONTHS)["cum_forecast"]
    cum_forecast_total = actuals.cell(total_row, cum_forecast_col).value
    assert cum_forecast_total == sum_of_leaves, (
        f"Actuals TOTAL Cum. Forecast EUR: expected {sum_of_leaves}, "
        f"got {cum_forecast_total!r} — current_month gate may have masked "
        f"month columns (edits at months {m_lo},{m_mid},{m_hi}; "
        f"current_month={current_month})"
    )

    # ── Assertion 5: Summary 'Fcst €' for the target-year column includes
    #    the three leaves (all three writes land in the same calendar year
    #    by construction above). ──
    summary = wb[SHEET_SUMMARY]
    # Year-group header is merged across COLS_PER_YEAR cols on the year row
    # above the sub-header. "Fcst €" is the 4th sub-column (index +3).
    yr_start_col, _yr_end_col = _summary_year_columns_for(summary, target_year)
    fcst_col = yr_start_col + 3  # ["Plan FTE", "Plan €", "Act €", "Fcst €", "Var €"]

    summary_fcst_total = 0
    for _person, wp_id, _m, _h in edits:
        sm_row = _summary_first_wp_row_data_only(summary, wp_id)
        v = summary.cell(sm_row, fcst_col).value
        assert isinstance(v, (int, float)), (
            f"Summary Fcst € for {wp_id}/{target_year} should be numeric, got {v!r}"
        )
        summary_fcst_total += v

    # Each WP's target-year Fcst € sums the Plan-Forecast WP row across that
    # year's columns. With only one leaf per (WP, month), each WP's
    # contribution == its single leaf, so the per-WP sum across our 3
    # writes == sum_of_leaves.
    assert summary_fcst_total == sum_of_leaves, (
        f"Summary {target_year} Fcst € over the 3 WPs: expected {sum_of_leaves}, "
        f"got {summary_fcst_total}"
    )

    wb.close()
