"""Tests for the Actuals-Hours sheet builder (T3: hour inputs only)."""

import pytest
from openpyxl import load_workbook

from budget_planner.workbook import init_workbook
from budget_planner.config import (
    DEMO_CAPACITY_ASSIGNMENTS, EXTERNAL_COMMITMENTS, TEAM,
    PROJECT_MONTHS,
)
from budget_planner.layout import MONTH_COL_START, val_columns
from budget_planner.styles import GREY_FILL, col_letter


FORECAST_COL_START = val_columns(PROJECT_MONTHS)["sum"] + 2
FORECAST_COL_END = FORECAST_COL_START + PROJECT_MONTHS - 1


@pytest.fixture(scope="module")
def wb_path(tmp_path_factory):
    out = tmp_path_factory.mktemp("ah") / "wb.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    return str(out)


@pytest.fixture(scope="module")
def wb(wb_path):
    return load_workbook(wb_path, data_only=False)


@pytest.fixture(scope="module")
def hrs(wb):
    return wb["Actuals-Hours"]


def _person_assigned_wps(person_name):
    return {a["wp"] for a in DEMO_CAPACITY_ASSIGNMENTS
            if a["person"] == person_name}


def _find_person_header_row(ws, person_name):
    """Find row of the person header (merged A:B) for ``person_name``.

    Works on unrecalculated workbooks where col A holds a Config cross-ref
    formula rather than the literal name — we locate it by finding the
    SECTION_FILL merged header whose col-A formula points at the Config
    roster row derived from TEAM order.
    """
    # Person header rows have col A formula =Config!A<roster_row>; the
    # roster_row for TEAM[i] is predictable (roster first row + i), but to
    # avoid coupling to Config internals we scan for the merged header
    # whose col-A formula contains "Config!A" and whose col-B "TOTAL" row
    # two-or-more rows below matches. Simpler: walk rows, match by index.
    team_names = [m["name"] for m in TEAM]
    idx = team_names.index(person_name)
    # We need the idx-th section header. Sections start after month header
    # row. Find all header rows by scanning for "SECTION_FILL merged" via
    # col-A value starting with "=Config!A".
    headers = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith("=Config!A"):
            headers.append(r)
    # Filter to person-header rows: each person has exactly one such row
    # before its data rows. WP rows also use =Config!A<wp_row> but are not
    # merged A:B. Cheap filter: check the merge.
    merged_ranges = {str(mr): (mr.min_row, mr.max_row, mr.min_col, mr.max_col)
                     for mr in ws.merged_cells.ranges}
    person_headers = []
    for r in headers:
        is_person = any(
            mr_min_r == r and mr_min_c == 1 and mr_max_c >= 2
            for _k, (mr_min_r, _mx, mr_min_c, mr_max_c) in merged_ranges.items()
        )
        if is_person:
            person_headers.append(r)
    return person_headers[idx]


# ── Existence + tab order ──

def test_actuals_hours_sheet_exists(wb):
    assert "Actuals-Hours" in wb.sheetnames


def test_actuals_hours_tab_order_after_actuals(wb):
    names = wb.sheetnames
    assert names.index("Actuals-Hours") == names.index("Actuals") + 1


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
    # Alice has 3 WP rows (WP1, WP2, WP3) + 1 external (ProjectX)
    expected_data_rows = 3 + 1
    data_row_start = alice_header + 1
    data_row_end = alice_header + expected_data_rows
    total_row = data_row_end + 1
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
    """Bob has WP4/WP5/WP6 + ProjectX (non-EU) + HorizonEU-Alpha (EU).
    EU TOTAL must reference WP rows + HorizonEU-Alpha only, not ProjectX."""
    bob_header = _find_person_header_row(hrs, "Bob")
    # Bob: 3 WP rows (WP4, WP5, WP6) + 2 externals (ProjectX, HorizonEU-Alpha)
    wp_rows = [bob_header + 1, bob_header + 2, bob_header + 3]
    ext_projectx_row = bob_header + 4  # first external = ProjectX (not EU)
    ext_horizon_row = bob_header + 5   # second external = HorizonEU-Alpha (EU)
    total_row = bob_header + 6
    eu_total_row = bob_header + 7

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
    month columns."""
    # The band label sits at row 1 starting at FORECAST_COL_START.
    label = hrs.cell(1, FORECAST_COL_START).value
    assert isinstance(label, str) and "FORECAST EUR" in label.upper(), (
        f"Expected FORECAST EUR band label, got {label!r}"
    )
    # Merged across the band.
    top_left = f"{col_letter(FORECAST_COL_START)}1"
    bottom_right = f"{col_letter(FORECAST_COL_END)}1"
    merged = [str(mr) for mr in hrs.merged_cells.ranges]
    expected = f"{top_left}:{bottom_right}"
    assert expected in merged, (
        f"Expected merge {expected} in {merged}"
    )


def test_forecast_eur_leaf_formula_structure(hrs):
    """Pick one active WP-person-month cell and assert its forecast-EUR
    leaf formula has the IF-guarded VLOOKUP/hours_per_fte_month shape."""
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
    # Alice: 3 WP rows (header+1, +2, +3) + 1 external at +4 → TOTAL at +5.
    wp_rows = [alice_header + 1, alice_header + 2, alice_header + 3]
    total_row = alice_header + 5
    assert hrs.cell(total_row, 2).value == "TOTAL"
    fcl = col_letter(FORECAST_COL_START)
    formula = hrs.cell(total_row, FORECAST_COL_START).value
    assert isinstance(formula, str)
    assert formula.startswith("=SUM("), f"Expected SUM(), got {formula!r}"
    for wp_r in wp_rows:
        assert f"{fcl}{wp_r}" in formula, (
            f"TOTAL @ band M1 missing {fcl}{wp_r}: {formula!r}"
        )


def test_actuals_hours_has_only_sum_validation_column(hrs):
    """Year header row has 'Sum' at val_col_sum and no Budget/Variance
    label at val_col_budget — the hours log has no budget-vs-actual view.

    (val_col_diff is now reused by the forecast-EUR band; Task 4 added the
    collapsible parallel EUR columns starting at val_col_sum + 2.)
    """
    vcols = val_columns(PROJECT_MONTHS)
    # Year header row = 2 (row 1 now carries the FORECAST EUR band label).
    year_row = 2
    sum_label = hrs.cell(year_row, vcols["sum"]).value
    budget_label = hrs.cell(year_row, vcols["budget"]).value
    assert sum_label == "Sum"
    assert budget_label in (None, "")
