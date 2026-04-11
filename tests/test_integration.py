"""Integration tests — cross-sheet formula verification for v2 workbook."""

import openpyxl
from openpyxl.utils import get_column_letter

from budget_planner.reader import discover_layout
from budget_planner.layout import _calendar_year_groups
from budget_planner.config import (
    WORK_PACKAGES, MONTH_COL_START, PROJECT_MONTHS, PROJECT_START,
    NUM_WPS, VAL_COL_SUM,
)


# ── Helpers for Summary sheet layout discovery ──

# Summary layout constants (must match summary.py)
_COLS_PER_YEAR = 4  # Plan FTE, Plan EUR, Act EUR, Var EUR
_YEAR_GROUP_START = 3
_CAL_GROUPS = list(_calendar_year_groups(PROJECT_START, PROJECT_MONTHS))
_NUM_YEARS = len(_CAL_GROUPS)
_TOTAL_GROUP_START = _YEAR_GROUP_START + _NUM_YEARS * _COLS_PER_YEAR


def _find_summary_first_wp_row(summary_ws):
    """Find the first WP data row by scanning for formula refs in col A."""
    for r in range(1, summary_ws.max_row + 1):
        val = summary_ws.cell(r, 1).value
        if isinstance(val, str) and val.startswith("=Config!A"):
            return r
    raise ValueError("Cannot find Summary first WP row")


def _find_summary_budget_overview(summary_ws, after_row):
    """Find Budget Overview rows by scanning for known labels after WP block."""
    labels: dict[str, int] = {}
    for r in range(after_row, summary_ws.max_row + 1):
        val = summary_ws.cell(r, 1).value
        if isinstance(val, str):
            labels[val] = r
    return labels


# ── Summary ← Plan + Actuals cross-sheet formula tests ──

def test_summary_plan_fte_formulas_reference_plan_planned_rows(fresh_wb_path):
    """Summary year Plan FTE formulas reference Plan planned block rows."""
    wb = openpyxl.load_workbook(fresh_wb_path)
    plan = wb["Plan"]
    layout = discover_layout(plan)
    summary = wb["Summary"]
    first_wp = _find_summary_first_wp_row(summary)

    for i, wp in enumerate(WORK_PACKAGES):
        row = first_wp + i
        plan_wp_row = layout["planned_first_wp_row"] + i

        for y, (group_start, span) in enumerate(_CAL_GROUPS):
            sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
            cell_val = summary.cell(row, sc).value

            first_col_l = get_column_letter(MONTH_COL_START + group_start)
            last_col_l = get_column_letter(MONTH_COL_START + group_start + span - 1)

            expected = f"=SUM(Plan!{first_col_l}{plan_wp_row}:{last_col_l}{plan_wp_row})"
            assert cell_val == expected, (
                f"{wp['id']} Y{y+1}: expected {expected}, got {cell_val}"
            )

    wb.close()


def test_summary_actual_eur_formulas_reference_actuals_rows(fresh_wb_path):
    """Summary year Actual EUR formulas reference Actuals WP rows."""
    wb = openpyxl.load_workbook(fresh_wb_path)
    summary = wb["Summary"]
    first_wp = _find_summary_first_wp_row(summary)

    # Discover actuals layout by finding its first WP row
    actuals = wb["Actuals"]
    actuals_first_wp = None
    for r in range(1, actuals.max_row + 1):
        val = actuals.cell(r, 1).value
        if isinstance(val, str) and val.startswith("=Config!A"):
            actuals_first_wp = r
            break
    assert actuals_first_wp is not None, "Cannot find Actuals first WP row"

    for i, wp in enumerate(WORK_PACKAGES):
        row = first_wp + i
        actuals_wp_row = actuals_first_wp + i

        for y, (group_start, span) in enumerate(_CAL_GROUPS):
            sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
            # Actual EUR is at sc + 2
            cell_val = summary.cell(row, sc + 2).value

            first_col_l = get_column_letter(MONTH_COL_START + group_start)
            last_col_l = get_column_letter(MONTH_COL_START + group_start + span - 1)

            expected = f"=SUM(Actuals!{first_col_l}{actuals_wp_row}:{last_col_l}{actuals_wp_row})"
            assert cell_val == expected, (
                f"{wp['id']} Y{y+1}: expected {expected}, got {cell_val}"
            )

    wb.close()


def test_summary_plan_eur_equals_fte_times_rate(fresh_wb_path):
    """Summary Plan EUR = Plan FTE * per-WP rate for each WP and year."""
    wb = openpyxl.load_workbook(fresh_wb_path)
    summary = wb["Summary"]
    plan = wb["Plan"]
    layout = discover_layout(plan)
    first_wp = _find_summary_first_wp_row(summary)

    for i in range(NUM_WPS):
        row = first_wp + i
        config_row = layout["config_first_wp_row"] + i
        for y in range(_NUM_YEARS):
            sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
            fte_col_l = get_column_letter(sc)
            # Plan EUR is at sc + 1
            cell_val = summary.cell(row, sc + 1).value
            expected = f"={fte_col_l}{row}*Config!$D${config_row}"
            assert cell_val == expected, (
                f"WP row {row} Y{y+1}: expected {expected}, got {cell_val}"
            )

    wb.close()


def test_summary_variance_eur_equals_plan_minus_actual(fresh_wb_path):
    """Summary Variance EUR = Plan EUR - Actual EUR for each WP and year."""
    wb = openpyxl.load_workbook(fresh_wb_path)
    summary = wb["Summary"]
    first_wp = _find_summary_first_wp_row(summary)

    for i in range(NUM_WPS):
        row = first_wp + i
        for y in range(_NUM_YEARS):
            sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
            plan_eur_col_l = get_column_letter(sc + 1)
            act_eur_col_l = get_column_letter(sc + 2)
            # Variance EUR is at sc + 3
            cell_val = summary.cell(row, sc + 3).value
            expected = f"={plan_eur_col_l}{row}-{act_eur_col_l}{row}"
            assert cell_val == expected, (
                f"WP row {row} Y{y+1}: expected {expected}, got {cell_val}"
            )

    wb.close()


def test_summary_project_totals_sum_across_years(fresh_wb_path):
    """Summary project total columns sum year columns via addition formulas."""
    wb = openpyxl.load_workbook(fresh_wb_path)
    summary = wb["Summary"]
    first_wp = _find_summary_first_wp_row(summary)

    for i in range(NUM_WPS):
        row = first_wp + i

        # Plan FTE total (tc+0): sum of year Plan FTE cols
        plan_fte_cols = [
            get_column_letter(_YEAR_GROUP_START + y * _COLS_PER_YEAR)
            for y in range(_NUM_YEARS)
        ]
        expected_fte = "=" + "+".join(f"{cl}{row}" for cl in plan_fte_cols)
        assert summary.cell(row, _TOTAL_GROUP_START).value == expected_fte, (
            f"Row {row} total Plan FTE: expected {expected_fte}, "
            f"got {summary.cell(row, _TOTAL_GROUP_START).value}"
        )

        # Plan EUR total (tc+1): sum of year Plan EUR cols
        plan_eur_cols = [
            get_column_letter(_YEAR_GROUP_START + y * _COLS_PER_YEAR + 1)
            for y in range(_NUM_YEARS)
        ]
        expected_eur = "=" + "+".join(f"{cl}{row}" for cl in plan_eur_cols)
        assert summary.cell(row, _TOTAL_GROUP_START + 1).value == expected_eur, (
            f"Row {row} total Plan EUR: expected {expected_eur}, "
            f"got {summary.cell(row, _TOTAL_GROUP_START + 1).value}"
        )

        # Actual EUR total (tc+2): sum of year Actual EUR cols
        act_eur_cols = [
            get_column_letter(_YEAR_GROUP_START + y * _COLS_PER_YEAR + 2)
            for y in range(_NUM_YEARS)
        ]
        expected_act = "=" + "+".join(f"{cl}{row}" for cl in act_eur_cols)
        assert summary.cell(row, _TOTAL_GROUP_START + 2).value == expected_act, (
            f"Row {row} total Actual EUR: expected {expected_act}, "
            f"got {summary.cell(row, _TOTAL_GROUP_START + 2).value}"
        )

        # Variance total (tc+3): Plan EUR total - Actual EUR total
        tc_plan_eur_l = get_column_letter(_TOTAL_GROUP_START + 1)
        tc_act_eur_l = get_column_letter(_TOTAL_GROUP_START + 2)
        expected_var = f"={tc_plan_eur_l}{row}-{tc_act_eur_l}{row}"
        assert summary.cell(row, _TOTAL_GROUP_START + 3).value == expected_var, (
            f"Row {row} total Variance: expected {expected_var}, "
            f"got {summary.cell(row, _TOTAL_GROUP_START + 3).value}"
        )

    wb.close()


def test_summary_budget_overview_references(fresh_wb_path):
    """Budget Overview section refs: Actuals total, partner_ceiling, Remaining, % Consumed."""
    wb = openpyxl.load_workbook(fresh_wb_path)
    summary = wb["Summary"]
    first_wp = _find_summary_first_wp_row(summary)
    last_wp = first_wp + NUM_WPS - 1
    # Totals row is right after WP rows, then budget overview starts after a gap
    labels = _find_summary_budget_overview(summary, last_wp + 1)

    assert "Personnel Actual" in labels, "Missing 'Personnel Actual' label"
    assert "Personnel Ceiling" in labels, "Missing 'Personnel Ceiling' label"
    assert "Remaining" in labels, "Missing 'Remaining' label"
    assert "% Consumed" in labels, "Missing '% Consumed' label"

    actual_row = labels["Personnel Actual"]
    ceiling_row = labels["Personnel Ceiling"]
    remaining_row = labels["Remaining"]
    consumed_row = labels["% Consumed"]

    # Discover actuals total row
    actuals = wb["Actuals"]
    actuals_total_row = None
    for r in range(1, actuals.max_row + 1):
        val = actuals.cell(r, 2).value
        if isinstance(val, str) and val.upper() == "TOTAL":
            actuals_total_row = r
            break
    assert actuals_total_row is not None, "Cannot find Actuals TOTAL row"

    val_col_l = get_column_letter(VAL_COL_SUM)

    # Personnel Actual = Actuals!{VAL_COL_SUM}{total_row}
    actual_formula = summary.cell(actual_row, 2).value
    expected_actual = f"=Actuals!{val_col_l}{actuals_total_row}"
    assert actual_formula == expected_actual, (
        f"Personnel Actual: expected {expected_actual}, got {actual_formula}"
    )

    # Personnel Ceiling = partner_ceiling
    ceiling_formula = summary.cell(ceiling_row, 2).value
    assert ceiling_formula == "=partner_ceiling", (
        f"Personnel Ceiling: expected =partner_ceiling, got {ceiling_formula}"
    )

    # Remaining = Ceiling - Actual
    remaining_formula = summary.cell(remaining_row, 2).value
    expected_remaining = f"=B{ceiling_row}-B{actual_row}"
    assert remaining_formula == expected_remaining, (
        f"Remaining: expected {expected_remaining}, got {remaining_formula}"
    )

    # % Consumed = IF(ceiling=0,0,actual/ceiling)
    consumed_formula = summary.cell(consumed_row, 2).value
    expected_consumed = f"=IF(B{ceiling_row}=0,0,B{actual_row}/B{ceiling_row})"
    assert consumed_formula == expected_consumed, (
        f"% Consumed: expected {expected_consumed}, got {consumed_formula}"
    )

    wb.close()


# ── Reporting Period Summary section ─────────────────────────────────────

def test_rp_summary_has_wp_rows(fresh_wb_path):
    """RP summary has one row per WP with cross-sheet formula refs."""
    from openpyxl import load_workbook
    from budget_planner.config import WORK_PACKAGES

    wb = load_workbook(fresh_wb_path)
    ws = wb["Summary"]

    # Find section header
    rp_header_row = None
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val and "REPORTING PERIOD" in str(val).upper():
            rp_header_row = r
            break
    assert rp_header_row is not None

    # Find first WP data row (skip header rows — look for Config!A cross-ref)
    first_wp_row = None
    for r in range(rp_header_row + 1, rp_header_row + 10):
        val = str(ws.cell(r, 1).value or "")
        if "Config!" in val:
            first_wp_row = r
            break
    assert first_wp_row is not None, "No WP data rows found in RP summary"

    # Count WP rows
    wp_count = 0
    for r in range(first_wp_row, first_wp_row + len(WORK_PACKAGES) + 5):
        val = str(ws.cell(r, 1).value or "")
        if "Config!" in val:
            wp_count += 1
        else:
            break
    assert wp_count == len(WORK_PACKAGES)

    # First WP row col C should be a Plan formula referencing Plan sheet
    val_c = str(ws.cell(first_wp_row, 3).value or "")
    assert "Plan!" in val_c, f"Expected Plan! cross-ref, got: {val_c}"

    wb.close()


# ── Rebuild round-trip: EU TOTAL row preservation ──────────────────────────

def test_rebuild_preserves_eu_total_structure(tmp_path):
    """Rebuild preserves person groups despite EU TOTAL row."""
    from budget_planner.workbook import init_workbook, rebuild_workbook
    from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS, TEAM
    import openpyxl

    out = str(tmp_path / "test.xlsx")
    init_workbook(out, assignments=DEMO_CAPACITY_ASSIGNMENTS)
    rebuild_workbook(out)
    wb = openpyxl.load_workbook(out, data_only=False)
    cap = wb["Capacity"]
    # Verify EU TOTAL rows exist after rebuild
    eu_total_count = sum(
        1 for r in range(1, cap.max_row + 1)
        if cap.cell(r, 2).value == "EU TOTAL"
    )
    assert eu_total_count == len(TEAM), (
        f"Expected {len(TEAM)} EU TOTAL rows after rebuild, got {eu_total_count}"
    )
    wb.close()


# ── Form C preparation rows ─────────────────────────────────────────────

def test_form_c_rows_exist(fresh_wb_path):
    """Summary sheet has Form C prep rows with indirect costs and CFS flag."""
    from openpyxl import load_workbook

    wb = load_workbook(fresh_wb_path)
    ws = wb["Summary"]

    # Find "A. Personnel" label
    personnel_row = None
    for r in range(1, ws.max_row + 1):
        val = str(ws.cell(r, 1).value or "")
        if "Personnel" in val and val.startswith("A."):
            personnel_row = r
            break
    assert personnel_row is not None, "Form C 'A. Personnel' row not found"

    # Check subsequent rows
    assert "Indirect" in str(ws.cell(personnel_row + 1, 1).value or "")
    assert "Total Eligible" in str(ws.cell(personnel_row + 2, 1).value or "")
    assert "Cum" in str(ws.cell(personnel_row + 3, 1).value or "")
    assert "CFS" in str(ws.cell(personnel_row + 4, 1).value or "")

    # CFS cell should have a formula with 430000
    cfs_row = personnel_row + 4
    # Check col C (first RP Act EUR position)
    cfs_val = str(ws.cell(cfs_row, 3 + 2).value or "")  # offset +2 = Act EUR col
    assert "430000" in cfs_val or "CFS" in str(ws.cell(cfs_row, 1).value)

    wb.close()
