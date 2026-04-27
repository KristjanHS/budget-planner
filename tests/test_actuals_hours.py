"""Tests for the Actuals-Hours sheet builder (T3: hour inputs only)."""

import pytest
from openpyxl import load_workbook

from budget_planner.config import (
    DEMO_CAPACITY_ASSIGNMENTS, EXTERNAL_COMMITMENTS, TEAM,
    PROJECT_MONTHS,
)
from budget_planner.layout import MONTH_COL_START, val_columns
from budget_planner.styles import GREY_FILL, col_letter


FORECAST_COL_START = val_columns(PROJECT_MONTHS)["sum"] + 2
FORECAST_COL_END = FORECAST_COL_START + PROJECT_MONTHS - 1
# Forecast EUR +i rollup col sits adjacent to (outside) the -i band.
FORECAST_PLUS_I_COL = FORECAST_COL_END + 1


@pytest.fixture(scope="module")
def wb(assigned_wb_path):
    return load_workbook(assigned_wb_path, data_only=False)


@pytest.fixture(scope="module")
def hrs(wb):
    return wb["Actuals-Hours"]


def _person_assigned_wps(person_name):
    return {a["wp"] for a in DEMO_CAPACITY_ASSIGNMENTS
            if a["person"] == person_name}


def _find_person_header_row(ws, person_name):
    """Find row of the person header (merged A:B) for ``person_name``.

    Person headers now carry the literal person name in col A (changed
    from the previous ``=Config!A<roster_row>`` formula so the forecast-
    EUR VLOOKUP against $A$<header_row> works under data_only reads).
    We locate the row by matching col-A value == person_name AND
    verifying the A:B merge.
    """
    merged_ranges = [(mr.min_row, mr.min_col, mr.max_col)
                     for mr in ws.merged_cells.ranges]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value != person_name:
            continue
        is_person = any(
            mr_min_r == r and mr_min_c == 1 and mr_max_c >= 2
            for mr_min_r, mr_min_c, mr_max_c in merged_ranges
        )
        if is_person:
            return r
    raise AssertionError(
        f"Person header for {person_name!r} not found on {ws.title}"
    )


# ── Existence ──
# Tab position covered by tests/test_tab_order.py (Actuals-Hours = position 3).


def test_actuals_hours_sheet_exists(wb):
    assert "Actuals-Hours" in wb.sheetnames


# ── Structure: person groups match team size ──

def test_actuals_hours_has_person_group_per_team_member(hrs):
    """One header row per team member."""
    team_names = [m["name"] for m in TEAM]
    rows = [_find_person_header_row(hrs, n) for n in team_names]
    assert len(set(rows)) == len(team_names)
    # Rows monotonically increase (sections stack top-down)
    assert rows == sorted(rows)


def test_actuals_hours_person_has_assigned_wp_rows(hrs):
    """Each person has a WP-ID row for each assigned WP, followed by TOTAL,
    EU TOTAL. Externals appear between WPs and TOTAL."""
    for member in TEAM:
        name = member["name"]
        header_row = _find_person_header_row(hrs, name)
        expected_wps = _person_assigned_wps(name)
        expected_externals = [e for e in EXTERNAL_COMMITMENTS
                              if e["person"] == name]

        # WP rows immediately follow header
        wp_row_count = 0
        r = header_row + 1
        seen_wp_ids: set[str] = set()
        while True:
            v = hrs.cell(r, 1).value
            if isinstance(v, str) and v.startswith("=Config!A"):
                # Could be a WP row (points at wp_row) — accept as WP row
                wp_row_count += 1
                # Derive which WP by looking at B formula is harder; we
                # just count rows here and check count.
                seen_wp_ids.add(f"wp-{r}")
                r += 1
                continue
            # External project row: col A = literal project name (str)
            if isinstance(v, str) and v in {e["project"]
                                             for e in expected_externals}:
                r += 1
                continue
            break
        assert wp_row_count == len(expected_wps), (
            f"{name}: expected {len(expected_wps)} WP rows, got {wp_row_count}"
        )

        # Next two rows are the hours-to-date summary (col A blank;
        # col B carries the label).
        assert hrs.cell(r, 2).value == "Cumulative hours to date"
        assert hrs.cell(r + 1, 2).value == "Actual PM to date"
        r += 2

        # Next row should be TOTAL label
        total_label = hrs.cell(r, 2).value
        assert total_label == "TOTAL", (
            f"{name}: expected TOTAL at row {r}, got {total_label!r}"
        )
        # Then EU TOTAL
        eu_label = hrs.cell(r + 1, 2).value
        assert eu_label == "EU TOTAL", (
            f"{name}: expected EU TOTAL at row {r + 1}, got {eu_label!r}"
        )


# ── Inactive months: grey + empty ──

def test_actuals_hours_inactive_month_is_empty_and_grey(hrs):
    """Bob's WP4 row is inactive at M1 (WP4 schedule starts M2)."""
    bob_header = _find_person_header_row(hrs, "Bob")
    wp4_row = bob_header + 1  # first WP row for Bob is WP4
    m1_col = MONTH_COL_START  # M1
    cell = hrs.cell(wp4_row, m1_col)
    assert cell.value in (None, ""), (
        f"Bob WP4 @ M1 should be empty (inactive); got {cell.value!r}"
    )
    # Grey fill
    fill = cell.fill
    assert fill is not None
    # PatternFill equality is stricter — compare by fgColor rgb
    assert str(getattr(fill.fgColor, "rgb", "")).upper().endswith(
        str(GREY_FILL.fgColor.rgb).upper()
    ), f"Bob WP4 @ M1 fill should be GREY_FILL, got {fill.fgColor.rgb!r}"


# ── Active months: blue input + empty (no demo hours pre-filled) ──

def test_actuals_hours_active_month_is_blue_and_empty(hrs):
    """Alice WP1 is active M1-M14. Her WP1 row at M1 must be blue + empty."""
    alice_header = _find_person_header_row(hrs, "Alice")
    wp1_row = alice_header + 1
    m1_col = MONTH_COL_START
    cell = hrs.cell(wp1_row, m1_col)
    # Empty (no demo pre-fill on hours)
    assert cell.value in (None, ""), (
        f"Alice WP1 @ M1 should be empty; got {cell.value!r}"
    )
    # Blue font (0000FF)
    color = getattr(cell.font.color, "rgb", None)
    assert color is not None
    assert str(color).upper().endswith("0000FF"), (
        f"Alice WP1 @ M1 font should be blue, got {color!r}"
    )
    # Integer number format — no decimals
    assert "0" in cell.number_format
    assert "." not in cell.number_format


# ── TOTAL row formula ──

def test_actuals_hours_total_row_sums_data_rows(hrs):
    """TOTAL row M1 formula = SUM of the data rows above for that person."""
    alice_header = _find_person_header_row(hrs, "Alice")
    # Derive Alice's expected data-row count from DEMO_CAPACITY_ASSIGNMENTS +
    # EXTERNAL_COMMITMENTS so this test tracks demo edits per the project
    # testing rule (no hardcoded seed values).
    alice_wps = sum(1 for a in DEMO_CAPACITY_ASSIGNMENTS if a["person"] == "Alice")
    alice_externals = sum(1 for e in EXTERNAL_COMMITMENTS if e["person"] == "Alice")
    expected_data_rows = alice_wps + alice_externals
    data_row_start = alice_header + 1
    data_row_end = alice_header + expected_data_rows
    # +2 for the hours-to-date summary rows between externals and TOTAL.
    total_row = data_row_end + 3
    assert hrs.cell(total_row, 2).value == "TOTAL"

    m1_col = MONTH_COL_START
    cl = col_letter(m1_col)
    formula = hrs.cell(total_row, m1_col).value
    expected = f"=SUM({cl}{data_row_start}:{cl}{data_row_end})"
    assert formula == expected, (
        f"TOTAL @ M1 expected {expected!r}, got {formula!r}"
    )


# ── EU TOTAL row formula ──

def test_actuals_hours_eu_total_references_grant_wps_and_eu_externals(hrs):
    """EU TOTAL must reference Bob's WP rows + EU external (HorizonEU-Alpha)
    only, not the non-EU external (ProjectX). Row layout per person:
    WP rows → externals (in DEMO order) → 2 hours-to-date summary rows →
    TOTAL → EU TOTAL.
    """
    bob_header = _find_person_header_row(hrs, "Bob")
    # Derive Bob's row layout from DEMO_CAPACITY_ASSIGNMENTS + EXTERNAL_COMMITMENTS
    # rather than hardcoding WP4/WP5/WP6 (per project testing rule).
    bob_wps = [a["wp"] for a in DEMO_CAPACITY_ASSIGNMENTS if a["person"] == "Bob"]
    bob_externals = [e for e in EXTERNAL_COMMITMENTS if e["person"] == "Bob"]
    num_wp = len(bob_wps)
    wp_rows = list(range(bob_header + 1, bob_header + 1 + num_wp))
    # External rows follow WP rows; pick the first non-EU and first EU rows
    # for the (don't-include) / (must-include) assertions below.
    ext_first_row = bob_header + 1 + num_wp
    ext_projectx_row = next(
        ext_first_row + i for i, e in enumerate(bob_externals) if not e["is_eu"]
    )
    ext_horizon_row = next(
        ext_first_row + i for i, e in enumerate(bob_externals) if e["is_eu"]
    )
    # +2 hours-to-date summary rows between externals and TOTAL.
    total_row = bob_header + num_wp + len(bob_externals) + 3
    eu_total_row = total_row + 1

    assert hrs.cell(total_row, 2).value == "TOTAL"
    assert hrs.cell(eu_total_row, 2).value == "EU TOTAL"

    m1_col = MONTH_COL_START
    cl = col_letter(m1_col)
    formula = hrs.cell(eu_total_row, m1_col).value
    # Must include all 3 WP rows and the EU external row
    for row in wp_rows + [ext_horizon_row]:
        assert f"{cl}{row}" in (formula or ""), (
            f"EU TOTAL @ M1 missing {cl}{row}: {formula!r}"
        )
    # Must NOT include the non-EU ProjectX external row
    assert f"{cl}{ext_projectx_row}" not in (formula or ""), (
        f"EU TOTAL @ M1 should not include non-EU ProjectX row "
        f"{cl}{ext_projectx_row}: {formula!r}"
    )


# ── EU ceiling CF rules ──

def test_actuals_hours_eu_ceiling_cf_per_year(hrs):
    """One red + one amber CF rule on EU TOTAL per calendar year."""
    # Count CF rules that reference eu_annual_ceiling_pm. The EU TOTAL row
    # has 2 rules per calendar year (red + amber). Project spans 14 months
    # starting Nov 2026 → years 2026, 2027 = 2 years per
    # person; 4 persons → 16 rules total expected.
    ceiling_rules = 0
    for cf_range in hrs.conditional_formatting:
        rules = hrs.conditional_formatting[cf_range]
        for rule in rules:
            formulas = getattr(rule, "formula", []) or []
            for f in formulas:
                if "eu_annual_ceiling_pm" in f and "hours_per_fte_month" in f:
                    ceiling_rules += 1
    # 2 calendar years × 2 severities × 4 people = 16
    assert ceiling_rules == 2 * 2 * len(TEAM), (
        f"Expected {5 * 2 * len(TEAM)} EU-ceiling CF rules, got {ceiling_rules}"
    )


# ── Sum column only (no Budget/Variance) ──

# ── Forecast-EUR band (Task 4) ──

def test_forecast_eur_band_columns_are_outline_grouped(hrs):
    """Each forecast-EUR month column is outline_level 1 so the band can
    be collapsed via Excel's outline toggle."""
    for col in range(FORECAST_COL_START, FORECAST_COL_END + 1):
        cl = col_letter(col)
        level = hrs.column_dimensions[cl].outline_level
        assert level == 1, (
            f"Col {cl} expected outline_level=1, got {level}"
        )
    # Outline summary button on the right.
    assert hrs.sheet_properties.outlinePr.summaryRight is True


def test_forecast_eur_header_exists(hrs):
    """A merged band-label cell contains 'FORECAST EUR' above the forecast
    month columns. Row 2 (row 1 now carries the through-month strip)."""
    label = hrs.cell(2, FORECAST_COL_START).value
    assert isinstance(label, str) and "FORECAST EUR" in label.upper(), (
        f"Expected FORECAST EUR band label, got {label!r}"
    )
    # Merged across the band.
    top_left = f"{col_letter(FORECAST_COL_START)}2"
    bottom_right = f"{col_letter(FORECAST_COL_END)}2"
    merged = [str(mr) for mr in hrs.merged_cells.ranges]
    expected = f"{top_left}:{bottom_right}"
    assert expected in merged, (
        f"Expected merge {expected} in {merged}"
    )


def test_forecast_eur_leaf_formula_structure(hrs):
    """Pick one active WP-person-month cell and assert its forecast-EUR
    leaf formula has the IF-guarded ROUND(VLOOKUP(.)/hours_per_fte_month)
    shape. The VLOOKUP lookup value references the person header cell
    (now a literal name, not a formula) via $A$<header_row>."""
    # Alice WP1 M1 is active — hour cell is empty input; leaf formula
    # references it and the roster_lookup rate via hours_per_fte_month.
    alice_header = _find_person_header_row(hrs, "Alice")
    wp1_row = alice_header + 1
    leaf = hrs.cell(wp1_row, FORECAST_COL_START).value
    assert isinstance(leaf, str), f"Expected str formula, got {leaf!r}"
    assert leaf.startswith("=IF("), f"Formula must start with =IF(: {leaf!r}"
    assert "hours_per_fte_month" in leaf, f"Missing hours_per_fte_month: {leaf!r}"
    assert "roster_lookup" in leaf, f"Missing roster_lookup: {leaf!r}"
    # Column 4 is the Rate column in the widened roster_lookup.
    assert ",4," in leaf.replace(" ", ""), (
        f"Formula should VLOOKUP col 4 (Rate): {leaf!r}"
    )
    # ROUND(..., 0) wrapper ensures each leaf is a whole EUR — avoids
    # sub-cent drift when Plan-Forecast SUMs hundreds of these leaves.
    assert "ROUND(" in leaf, (
        f"Formula should wrap ROUND(...,0) for whole-EUR leaves: {leaf!r}"
    )
    # Person header is literal (not =Config!A<roster_row>) so data_only
    # reads of the workbook can evaluate the VLOOKUP key correctly — the
    # leaf must NOT resolve the lookup-value via a Config cross-ref.
    assert "=Config!A" not in leaf, (
        f"Leaf must not resolve lookup-value via Config formula: {leaf!r}"
    )


def test_person_header_is_literal_name(hrs):
    """Person header cell (col A of the merged A:B banner) must carry the
    literal person name, not a formula. The forecast-EUR VLOOKUP uses
    this cell as the lookup value; a formula there would evaluate to
    None under data_only reads and silently break the EUR column."""
    for member in TEAM:
        name = member["name"]
        header_row = _find_person_header_row(hrs, name)
        value = hrs.cell(header_row, 1).value
        assert isinstance(value, str), (
            f"Person header for {name} must be a string, got {value!r}"
        )
        assert not value.startswith("="), (
            f"Person header for {name} must be literal, got formula {value!r}"
        )
        assert value == name, (
            f"Person header at row {header_row} expected {name!r}, "
            f"got {value!r}"
        )


def test_forecast_eur_external_row_empty(hrs):
    """External-project rows in the forecast band are empty — externals are
    non-programme projects and don't consume programme EUR budget."""
    # Alice has 3 WP rows (WP1-WP3) + 1 external (ProjectX).
    alice_header = _find_person_header_row(hrs, "Alice")
    alice_ext_row = alice_header + 4  # 3 WP rows + 1st external
    # The row label in col A should be the literal project name.
    assert hrs.cell(alice_ext_row, 1).value == "ProjectX"
    # All forecast-band cells on that row must be empty.
    for mi in range(PROJECT_MONTHS):
        cell = hrs.cell(alice_ext_row, FORECAST_COL_START + mi)
        assert cell.value in (None, ""), (
            f"External row @ band col {mi} should be empty, got {cell.value!r}"
        )


def test_forecast_eur_inactive_month_empty_and_grey(hrs):
    """Inactive-month forecast-band cells carry GREY_FILL and are empty —
    mirror the hour side so the band visually aligns."""
    # Bob's WP4 row is inactive at M1 (WP4 starts M2).
    bob_header = _find_person_header_row(hrs, "Bob")
    wp4_row = bob_header + 1
    cell = hrs.cell(wp4_row, FORECAST_COL_START)  # M1 on forecast side
    assert cell.value in (None, "")
    assert str(getattr(cell.fill.fgColor, "rgb", "")).upper().endswith(
        str(GREY_FILL.fgColor.rgb).upper()
    ), f"Inactive forecast cell should be grey, got {cell.fill.fgColor.rgb!r}"


def test_forecast_eur_total_row_sums_wp_rows(hrs):
    """Per-person TOTAL row in the forecast band = SUM of WP-row forecast
    cells for that month. Externals are blank so they aren't in the SUM."""
    alice_header = _find_person_header_row(hrs, "Alice")
    # Alice: 3 WP rows (header+1..+3) + 1 external (+4) + 2 hours-to-date
    # summary rows (+5, +6) → TOTAL at +7.
    wp_rows = [alice_header + 1, alice_header + 2, alice_header + 3]
    total_row = alice_header + 7
    assert hrs.cell(total_row, 2).value == "TOTAL"
    fcl = col_letter(FORECAST_COL_START)
    formula = hrs.cell(total_row, FORECAST_COL_START).value
    assert isinstance(formula, str)
    assert formula.startswith("=SUM("), f"Expected SUM(), got {formula!r}"
    for wp_r in wp_rows:
        assert f"{fcl}{wp_r}" in formula, (
            f"TOTAL @ band M1 missing {fcl}{wp_r}: {formula!r}"
        )


def test_forecast_plus_i_col_outside_band_outline(hrs):
    """The +i rollup col is NOT part of the -i column outline group, so it
    stays visible when the band collapses."""
    cl = col_letter(FORECAST_PLUS_I_COL)
    assert hrs.column_dimensions[cl].outline_level in (0, None), (
        f"+i col {cl} should not be in the -i outline group, "
        f"got outline_level={hrs.column_dimensions[cl].outline_level}"
    )


def test_forecast_plus_i_header_label(hrs):
    """Band-label row (row 2) carries 'Forecast EUR +i' on the +i rollup col."""
    label = hrs.cell(2, FORECAST_PLUS_I_COL).value
    assert label == "Forecast EUR +i", (
        f"Expected 'Forecast EUR +i' at +i col band-label row, got {label!r}"
    )


def test_forecast_plus_i_wp_row_formula(hrs):
    """WP rows carry `=ROUND(SUM(fc_start{r}:fc_end{r})*1.25,0)` on the +i col."""
    alice_header = _find_person_header_row(hrs, "Alice")
    wp1_row = alice_header + 1
    formula = hrs.cell(wp1_row, FORECAST_PLUS_I_COL).value
    fc_start_l = col_letter(FORECAST_COL_START)
    fc_end_l = col_letter(FORECAST_COL_END)
    expected = f"=ROUND(SUM({fc_start_l}{wp1_row}:{fc_end_l}{wp1_row})*1.25,0)"
    assert formula == expected, (
        f"+i WP row formula expected {expected!r}, got {formula!r}"
    )


def test_forecast_plus_i_external_row_empty(hrs):
    """External rows' +i cell is empty (externals don't contribute programme EUR)."""
    alice_header = _find_person_header_row(hrs, "Alice")
    alice_ext_row = alice_header + 4
    assert hrs.cell(alice_ext_row, 1).value == "ProjectX"
    cell = hrs.cell(alice_ext_row, FORECAST_PLUS_I_COL)
    assert cell.value in (None, ""), (
        f"External +i cell should be empty, got {cell.value!r}"
    )


def test_actuals_hours_has_only_sum_validation_column(hrs):
    """Year header row has 'Sum' at val_col_sum and no Budget/Variance
    label at val_col_budget — the hours log has no budget-vs-actual view.

    (val_col_diff is now reused by the forecast-EUR band; Task 4 added the
    collapsible parallel EUR columns starting at val_col_sum + 2.)
    """
    vcols = val_columns(PROJECT_MONTHS)
    # Year header row = 3 (row 1 = through-month strip, row 2 = band label,
    # row 3 = year, row 4 = month).
    year_row = 3
    sum_label = hrs.cell(year_row, vcols["sum"]).value
    budget_label = hrs.cell(year_row, vcols["budget"]).value
    assert sum_label == "Sum"
    assert budget_label in (None, "")


# ── Cum-hours helper column (feeds Capacity to-date Actual PM) ──

CUM_HOURS_COL = FORECAST_PLUS_I_COL + 2


def test_cum_hours_header_present(hrs):
    """Header row carries the 'Cum hours (to date)' label."""
    # Month header row = 4 (strip=1, band=2, year=3, month=4).
    assert hrs.cell(4, CUM_HOURS_COL).value == "Cum hours (to date)"


def test_cum_hours_wp_row_formula_shape(hrs):
    """WP rows carry a windowed SUM+INDEX formula keyed on
    ``hours_actuals_through_month``."""
    alice_row = _find_person_header_row(hrs, "Alice") + 1  # first WP row
    formula = hrs.cell(alice_row, CUM_HOURS_COL).value
    assert isinstance(formula, str)
    assert formula.startswith("=IF(hours_actuals_through_month<=0,0,")
    assert "INDEX" in formula
    start_cl = col_letter(MONTH_COL_START)
    end_cl = col_letter(MONTH_COL_START + PROJECT_MONTHS - 1)
    assert f"${start_cl}${alice_row}:${end_cl}${alice_row}" in formula


def test_cum_hours_total_and_eu_total_formulas(hrs):
    """TOTAL and EU TOTAL rows also carry the windowed SUM — Capacity's
    aggregate path doesn't need it today, but the helper col is uniform
    per row so cross-section reads remain consistent."""
    alice_header = _find_person_header_row(hrs, "Alice")
    total_row = None
    eu_total_row = None
    for r in range(alice_header + 1, alice_header + 30):
        if hrs.cell(r, 2).value == "TOTAL":
            total_row = r
        elif hrs.cell(r, 2).value == "EU TOTAL":
            eu_total_row = r
            break
    assert total_row is not None
    assert eu_total_row is not None
    for row in (total_row, eu_total_row):
        formula = hrs.cell(row, CUM_HOURS_COL).value
        assert isinstance(formula, str)
        assert "hours_actuals_through_month" in formula
        assert "INDEX" in formula
