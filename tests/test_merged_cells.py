"""Tests for merged cell ranges across sheets."""
import pytest
import openpyxl

from budget_planner.layout import _calendar_year_groups
from budget_planner.config import (
    PROJECT_MONTHS, PROJECT_START, MONTH_COL_START,
    VAL_COL_DIFF, VAL_COL_BREACH,
)


# ── Fixtures ──

@pytest.fixture(scope="module")
def workbook(fresh_wb_path):
    return openpyxl.load_workbook(fresh_wb_path, data_only=False)


# ── Helpers ──

def _year_merges_on_row(ws, row: int) -> list[tuple[int, int]]:
    """Return sorted (min_col, max_col) for all merges on the given row
    whose min_col >= MONTH_COL_START (i.e. year header merges, not section
    headers)."""
    result = []
    for rng in ws.merged_cells.ranges:
        if rng.min_row == row and rng.max_row == row and rng.min_col >= MONTH_COL_START:
            result.append((rng.min_col, rng.max_col))
    return sorted(result)


def _find_section_merge(ws, text: str) -> tuple[int, int, int] | None:
    """Find the merged range whose top-left cell contains *text*.

    Returns (row, min_col, max_col) or None.
    """
    for rng in ws.merged_cells.ranges:
        cell = ws.cell(row=rng.min_row, column=rng.min_col)
        if cell.value and text in str(cell.value):
            return (rng.min_row, rng.min_col, rng.max_col)
    return None


def _expected_year_spans() -> list[tuple[int, int]]:
    """Compute expected (min_col, max_col) for calendar-year header merges."""
    spans = []
    for group_start, span in _calendar_year_groups(PROJECT_START, PROJECT_MONTHS):
        year_start_col = MONTH_COL_START + group_start
        year_end_col = year_start_col + span - 1
        if span > 1:
            spans.append((year_start_col, year_end_col))
    return spans


# ── Tests ──

class TestPlanYearHeaders:
    """Plan year headers merge exactly 12 cells per year (or fewer for final year)."""

    def test_plan_year_header_merges(self, workbook):
        ws = workbook["Plan"]
        expected = _expected_year_spans()

        # The year header row is one row below the PLANNED FTE section header.
        info = _find_section_merge(ws, "PLANNED FTE")
        assert info is not None, "PLANNED FTE section header not found"
        year_row = info[0] + 1

        actual = _year_merges_on_row(ws, year_row)
        assert actual == expected, (
            f"Plan year header merges mismatch.\n"
            f"  Expected: {expected}\n"
            f"  Actual:   {actual}"
        )


class TestCapacityYearHeaders:
    """Capacity year headers merge correctly — same 12-col pattern."""

    def test_capacity_year_header_merges(self, workbook):
        ws = workbook["Capacity"]
        expected_planned = _expected_year_spans()

        # Capacity year headers are on the row right after the external
        # commitments config block + gap row.  Find by scanning for a merge
        # whose top-left value is the first calendar year (e.g. "2025").
        first_cal_year = str(PROJECT_START.year)
        year_row = None
        for rng in ws.merged_cells.ranges:
            cell = ws.cell(row=rng.min_row, column=rng.min_col)
            if cell.value and str(cell.value) == first_cal_year:
                year_row = rng.min_row
                break
        assert year_row is not None, "Capacity year header row not found"

        actual = _year_merges_on_row(ws, year_row)
        # T9 mirrors year spans onto Actual + Variance bands. Expected list
        # is the planned-side spans plus the two band mirrors.
        from budget_planner.config import PROJECT_MONTHS as _PM
        from budget_planner.layout import val_columns as _val_columns
        val_diff = _val_columns(_PM)["diff"]
        actual_start = val_diff + 2
        variance_start = actual_start + _PM + 1
        planned_start = 3  # MONTH_COL_START
        expected = list(expected_planned)
        for band_start in (actual_start, variance_start):
            for ps, pe in expected_planned:
                expected.append(
                    (ps - planned_start + band_start,
                     pe - planned_start + band_start)
                )
        assert actual == expected, (
            f"Capacity year header merges mismatch.\n"
            f"  Expected: {expected}\n"
            f"  Actual:   {actual}"
        )


class TestSummaryYearHeaders:
    """Summary year headers merge across 5 columns per year group
    (Task 10 added a Forecast EUR column between Actual EUR and Variance EUR).
    """

    def test_summary_year_header_merges(self, workbook):
        ws = workbook["Summary"]
        cal_groups = list(_calendar_year_groups(PROJECT_START, PROJECT_MONTHS))
        num_years = len(cal_groups)
        YEAR_GROUP_START = 3
        COLS_PER_YEAR = 5

        # Build expected year group merges
        expected_year = []
        for y in range(num_years):
            start_col = YEAR_GROUP_START + y * COLS_PER_YEAR
            end_col = start_col + COLS_PER_YEAR - 1
            expected_year.append((start_col, end_col))

        # PROJECT TOTAL merge
        total_start = YEAR_GROUP_START + num_years * COLS_PER_YEAR
        total_end = total_start + COLS_PER_YEAR - 1  # same width as year groups
        expected_total = (total_start, total_end)

        # Find the year header row — look for the first calendar year label
        first_cal_year = str(PROJECT_START.year)
        year_row = None
        for rng in ws.merged_cells.ranges:
            cell = ws.cell(row=rng.min_row, column=rng.min_col)
            if cell.value and str(cell.value) == first_cal_year:
                year_row = rng.min_row
                break
        assert year_row is not None, "Summary year header row not found"

        # Collect all merges on that row with col >= YEAR_GROUP_START
        actual = []
        for rng in ws.merged_cells.ranges:
            if rng.min_row == year_row and rng.max_row == year_row and rng.min_col >= YEAR_GROUP_START:
                actual.append((rng.min_col, rng.max_col))
        actual.sort()

        # Expected: all year merges + PROJECT TOTAL merge
        all_expected = sorted(expected_year + [expected_total])
        assert actual == all_expected, (
            f"Summary year header merges mismatch.\n"
            f"  Expected: {all_expected}\n"
            f"  Actual:   {actual}"
        )


class TestSectionHeaders:
    """Section headers span the correct column range."""

    @pytest.mark.parametrize("sheet_name,header_text,expected_cols", [
        ("Plan", "PLANNED FTE", (1, VAL_COL_DIFF)),
        ("Config", "EXTERNAL COMMITMENTS", (1, 6)),
        ("Actuals", "ACTUAL EUR", (1, VAL_COL_BREACH)),
    ])
    def test_section_header_span(self, workbook, sheet_name, header_text, expected_cols):
        ws = workbook[sheet_name]
        info = _find_section_merge(ws, header_text)
        assert info is not None, f"Section header '{header_text}' not found in {sheet_name}"
        _, min_col, max_col = info
        assert (min_col, max_col) == expected_cols, (
            f"{sheet_name} '{header_text}' merge spans cols {min_col}-{max_col}, "
            f"expected {expected_cols[0]}-{expected_cols[1]}"
        )
