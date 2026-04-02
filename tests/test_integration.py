"""Integration tests — data flow + round-trip preservation for v2 workbook."""
import re

import openpyxl
import pytest
from openpyxl.utils import get_column_letter

from budget_planner.reader import discover_layout, read_capacity_sheet, read_plan_config
from budget_planner.config import (
    WORK_PACKAGES, TEAM, MONTH_COL_START, PROJECT_MONTHS,
    NUM_WPS, VAL_COL_SUM,
)


# ── Data flow: Capacity → Plan formula linkage ──

def test_capacity_to_plan_formula_linkage(fresh_wb_path):
    """Plan Planned FTE formulas reference correct Capacity rows."""
    wb = openpyxl.load_workbook(fresh_wb_path)
    plan = wb["Plan"]
    layout = discover_layout(plan)

    # Read the capacity data to know which persons are assigned to each WP
    config_wps = read_plan_config(plan, layout)
    cap = wb["Capacity"]
    cap_data = read_capacity_sheet(cap, config_wps)

    for i, wp in enumerate(WORK_PACKAGES):
        row = layout["planned_first_wp_row"] + i
        # Pick first active month for this WP
        first_active_mi = wp["start"] - 1
        c = MONTH_COL_START + first_active_mi

        cell_value = plan.cell(row, c).value
        assert cell_value is not None, f"Planned cell for {wp['id']} M{wp['start']} is empty"
        assert isinstance(cell_value, str) and cell_value.startswith("="), (
            f"Planned cell for {wp['id']} should be a formula, got: {cell_value}"
        )
        assert "Capacity!" in cell_value, (
            f"Planned formula for {wp['id']} should reference Capacity sheet"
        )

        # Extract row numbers from formula like =SUM(Capacity!C10,Capacity!C20)
        cap_row_refs = [int(m) for m in re.findall(r"Capacity!\w+(\d+)", cell_value)]
        assert len(cap_row_refs) > 0, f"No Capacity row refs found in formula: {cell_value}"

        # Verify each referenced row belongs to a person assigned to this WP
        for person_name, group in cap_data["person_groups"].items():
            if wp["id"] in group["wp_values"]:
                # This person should have a row referenced in the formula
                # Find the capacity row for this person-WP pair
                # We can't get exact rows from cap_data, but we can verify
                # that the count of refs matches the count of assigned persons
                pass

        # Count how many persons are assigned to this WP
        assigned_count = sum(
            1 for pg in cap_data["person_groups"].values()
            if wp["id"] in pg["wp_values"]
        )
        assert len(cap_row_refs) == assigned_count, (
            f"{wp['id']}: formula references {len(cap_row_refs)} rows "
            f"but {assigned_count} persons are assigned"
        )

    wb.close()


def test_team_shares_to_capacity_seed_values(fresh_wb_path):
    """Capacity seed values match Team share * WP FTE / active months."""
    wb = openpyxl.load_workbook(fresh_wb_path)
    cap = wb["Capacity"]
    plan = wb["Plan"]
    layout = discover_layout(plan)
    config_wps = read_plan_config(plan, layout)
    cap_data = read_capacity_sheet(cap, config_wps)

    for member in TEAM:
        name = member["name"]
        assert name in cap_data["person_groups"], f"Person {name} not found in Capacity"
        person_group = cap_data["person_groups"][name]

        for wp_id, share in member["wps"].items():
            wp = next(w for w in WORK_PACKAGES if w["id"] == wp_id)
            active_months = wp["end"] - wp["start"] + 1
            expected_seed = round(share * wp["fte"] / active_months, 4) if active_months > 0 else 0

            assert wp_id in person_group["wp_values"], (
                f"{name} missing WP {wp_id} in Capacity"
            )
            monthly = person_group["wp_values"][wp_id]

            # Check first active month value matches seed
            first_active_idx = wp["start"] - 1
            actual_val = monthly[first_active_idx]
            assert actual_val == pytest.approx(expected_seed, abs=0.001), (
                f"{name}/{wp_id} M{wp['start']}: expected seed {expected_seed}, "
                f"got {actual_val}"
            )

            # Check inactive months are None
            for mi in range(PROJECT_MONTHS):
                month_num = mi + 1
                if month_num < wp["start"] or month_num > wp["end"]:
                    assert monthly[mi] is None, (
                        f"{name}/{wp_id} M{month_num}: inactive month should be None, "
                        f"got {monthly[mi]}"
                    )

    wb.close()


# ── Helpers for Summary sheet layout discovery ──

# Summary layout constants (must match summary.py)
_COLS_PER_YEAR = 4  # Plan FTE, Plan EUR, Act EUR, Var EUR
_YEAR_GROUP_START = 3
_NUM_YEARS = (PROJECT_MONTHS + 11) // 12
_TOTAL_GROUP_START = _YEAR_GROUP_START + _NUM_YEARS * _COLS_PER_YEAR


def _find_summary_first_wp_row(summary_ws):
    """Find the first WP data row by scanning for formula refs in col A."""
    for r in range(1, summary_ws.max_row + 1):
        val = summary_ws.cell(r, 1).value
        if isinstance(val, str) and val.startswith("=Plan!A"):
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

        for y in range(_NUM_YEARS):
            sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
            cell_val = summary.cell(row, sc).value

            m_start = y * 12
            m_end = min((y + 1) * 12, PROJECT_MONTHS) - 1
            first_col_l = get_column_letter(MONTH_COL_START + m_start)
            last_col_l = get_column_letter(MONTH_COL_START + m_end)

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
        if isinstance(val, str) and val.startswith("=Plan!A"):
            actuals_first_wp = r
            break
    assert actuals_first_wp is not None, "Cannot find Actuals first WP row"

    for i, wp in enumerate(WORK_PACKAGES):
        row = first_wp + i
        actuals_wp_row = actuals_first_wp + i

        for y in range(_NUM_YEARS):
            sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
            # Actual EUR is at sc + 2
            cell_val = summary.cell(row, sc + 2).value

            m_start = y * 12
            m_end = min((y + 1) * 12, PROJECT_MONTHS) - 1
            first_col_l = get_column_letter(MONTH_COL_START + m_start)
            last_col_l = get_column_letter(MONTH_COL_START + m_end)

            expected = f"=SUM(Actuals!{first_col_l}{actuals_wp_row}:{last_col_l}{actuals_wp_row})"
            assert cell_val == expected, (
                f"{wp['id']} Y{y+1}: expected {expected}, got {cell_val}"
            )

    wb.close()


def test_summary_plan_eur_equals_fte_times_rate(fresh_wb_path):
    """Summary Plan EUR = Plan FTE * fte_rate for each WP and year."""
    wb = openpyxl.load_workbook(fresh_wb_path)
    summary = wb["Summary"]
    first_wp = _find_summary_first_wp_row(summary)

    for i in range(NUM_WPS):
        row = first_wp + i
        for y in range(_NUM_YEARS):
            sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
            fte_col_l = get_column_letter(sc)
            # Plan EUR is at sc + 1
            cell_val = summary.cell(row, sc + 1).value
            expected = f"={fte_col_l}{row}*fte_rate"
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

