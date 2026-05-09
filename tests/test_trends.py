"""Tests for the Trends sheet builder."""
from typing import TYPE_CHECKING

import pytest
import openpyxl

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet

from budget_planner.config import (
    NUM_WPS, TEAM,
    VAL_COL_BUDGET, PROJECT_START,
)
from budget_planner.layout import MONTH_COL_START
from budget_planner.styles import col_letter


# ── Fixtures ──

@pytest.fixture(scope="module")
def trends_workbook(demo_wb_path):
    return openpyxl.load_workbook(demo_wb_path, data_only=False)


@pytest.fixture(scope="module")
def trends_ws(trends_workbook):
    return trends_workbook["Trends"]


def _find_row(ws, col, value):
    """Find the first row where ws.cell(row, col).value == value."""
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if v == value:
            return r
    return None


def _find_row_containing(ws, col, substring):
    """Find the first row where substring appears in str(ws.cell(row, col).value)."""
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if v is not None and substring in str(v):
            return r
    return None


# ================================================================
# Sheet basics
# ================================================================

def test_current_month_named_range(trends_workbook):
    defined = {}
    for name in trends_workbook.defined_names:
        dn = trends_workbook.defined_names[name]
        defined[dn.name] = dn.attr_text
    assert "current_month" in defined
    assert defined["current_month"] == "Trends!$B$1"


def test_current_month_default_value(trends_ws):
    """B1 has auto-calculating formula referencing project_start and TODAY()."""
    val = trends_ws["B1"].value
    assert isinstance(val, str) and val.startswith("="), (
        f"Expected formula in Trends!B1, got {val}"
    )
    assert "TODAY()" in val
    assert "project_start" in val


# ================================================================
# WP Burn Rate table
# ================================================================

def test_burn_rate_section_header(trends_ws):
    """Row 3 has the WP BURN RATE & EXHAUSTION section header."""
    val = trends_ws.cell(3, 1).value
    assert val is not None
    assert "WP BURN RATE" in str(val)


def test_burn_rate_wp_count(trends_ws):
    """8 WP rows between header row and TOTAL row."""
    total_row = _find_row(trends_ws, 2, "TOTAL")
    assert total_row is not None
    # Header row is 4, first WP is 5
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    wp_count = total_row - first_wp_row
    assert wp_count == NUM_WPS


def test_burn_rate_wp_id_refs_plan(trends_ws):
    """Col A formulas reference Config WP rows."""
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    for i in range(NUM_WPS):
        cell = trends_ws.cell(first_wp_row + i, 1)
        formula = str(cell.value)
        assert formula.startswith("=Config!A"), (
            f"WP row {i}: expected Config!A ref, got {formula}"
        )


def test_burn_rate_time_elapsed_formula(trends_ws):
    """Col C has formula referencing current_month for % Time Elapsed.

    Under the single-range schema the formula uses direct Config!I{row} /
    Config!J{row} refs (WP BUDGET Start/End cols post-indirect-view restructure;
    previously cols H/I before Budget +i was inserted at col F).
    """
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    cell = trends_ws.cell(first_wp_row, 3)
    formula = str(cell.value)
    assert "current_month" in formula
    assert "Config!I" in formula  # WP start col ref
    assert "Config!J" in formula  # WP end col ref


def test_burn_rate_budget_eur_formula(trends_ws):
    """Col D references Actuals cumulative columns for % Budget EUR."""
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    cell = trends_ws.cell(first_wp_row, 4)
    formula = str(cell.value)
    assert "Actuals!" in formula
    # Should reference both the cumulative column and the ceiling column
    ceil_letter = col_letter(VAL_COL_BUDGET)
    assert f"Actuals!{ceil_letter}" in formula


def test_burn_rate_delta_formula(trends_ws):
    """Col E = D - C (Burn Delta), short-circuited on posted_idx=0 like Col D
    (otherwise '—' - C propagates as #VALUE! once a WP starts but has no
    postings yet)."""
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    cell = trends_ws.cell(first_wp_row, 5)
    formula = str(cell.value)
    row = first_wp_row
    assert f"D{row}-C{row}" in formula or f"D{row} - C{row}" in formula
    assert "Actuals!" in formula and "=0" in formula, (
        f"Col E must guard on posted_idx=0 to avoid #VALUE! propagation: {formula!r}"
    )


def test_burn_rate_status_formula(trends_ws):
    """Col F has IF/ABS formula for Status; mirrors Col D's posted_idx=0
    short-circuit because ABS(E) propagates the upstream #VALUE! otherwise."""
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    cell = trends_ws.cell(first_wp_row, 6)
    formula = str(cell.value)
    assert "IF" in formula
    assert "ABS" in formula
    assert "Actuals!" in formula and "=0" in formula, (
        f"Col F must guard on posted_idx=0 to avoid #VALUE! propagation: {formula!r}"
    )


def test_burn_rate_exhaustion_formula(trends_ws):
    """Col G uses Actuals slope (per design §3) wrapped in ROUND, guarded by
    posted_idx=0 -> "No data". Phase 2.4 replaced the COUNTIF/cum heuristic
    with the actuals-tab `slope` column (avg-burn = posted_value / posted_idx).
    """
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    cell = trends_ws.cell(first_wp_row, 7)
    formula = str(cell.value)
    assert "ROUND" in formula
    assert "Actuals!" in formula
    assert '"No data"' in formula
    # COUNTIF heuristic is gone after Phase 2.4 (slope is now the source).
    assert "COUNTIF" not in formula, (
        f"Col G must not consult COUNTIF after Phase 2.4 rewire: {formula!r}"
    )


def test_burn_rate_total_row(trends_ws):
    """TOTAL row has 'TOTAL' in col B and weighted avg formula in col E."""
    total_row = _find_row(trends_ws, 2, "TOTAL")
    assert total_row is not None
    cell = trends_ws.cell(total_row, 5)
    formula = str(cell.value)
    assert "ISNUMBER" in formula
    assert "partner_ceiling" in formula


# ================================================================
# Person Load table
# ================================================================

def test_person_load_section_header(trends_ws):
    """Trends has a PERSON LOAD section header."""
    row = _find_row_containing(trends_ws, 1, "PERSON LOAD")
    assert row is not None, "PERSON LOAD section header not found in column A"
    assert "PERSON LOAD" in str(trends_ws.cell(row, 1).value)


def test_person_load_person_count(trends_ws):
    """4 person rows (from TEAM config)."""
    # Person header row is 33, data starts at 34
    header_row = _find_row_containing(trends_ws, 1, "Person")
    assert header_row is not None
    first_person_row = header_row + 1
    count = 0
    for i in range(20):  # generous upper bound
        cell = trends_ws.cell(first_person_row + i, 1)
        if cell.value is None or str(cell.value).strip() == "":
            break
        count += 1
    assert count == len(TEAM)


def test_person_load_headers_are_four(trends_ws):
    """Person Load header row has exactly 4 headers (Person, Avg FTE to date,
    Plan Avg FTE to date, Variance to-date %) — cols A-D; col E onward empty."""
    header_row = _find_row_containing(trends_ws, 1, "Person")
    assert header_row is not None
    expected = [
        "Person",
        "Avg FTE to date",
        "Plan Avg FTE to date",
        "Variance to-date %",
    ]
    for i, h in enumerate(expected, 1):
        got = trends_ws.cell(header_row, i).value
        assert got == h, f"Col {i} header: expected {h!r}, got {got!r}"
    # Col E (5) must be empty to prove no residual column was kept
    assert trends_ws.cell(header_row, 5).value is None, (
        f"Col E header should be empty, got {trends_ws.cell(header_row, 5).value!r}"
    )


def test_person_load_avg_to_date_formula(trends_ws):
    """Col B = Capacity to-date Actual PM / MAX(1, hours_actuals_through_month), IFERROR-wrapped."""
    header_row = _find_row_containing(trends_ws, 1, "Person")
    assert header_row is not None
    first_person_row = header_row + 1
    cell = trends_ws.cell(first_person_row, 2)
    formula = str(cell.value)
    assert "Capacity!" in formula, f"Col B must reference Capacity sheet: {formula!r}"
    assert "MAX(1,hours_actuals_through_month)" in formula, (
        f"Col B must divide by MAX(1,hours_actuals_through_month): {formula!r}"
    )
    assert "IFERROR" in formula, (
        f"Col B must wrap division in IFERROR to guard DIV/0: {formula!r}"
    )


def test_person_load_plan_avg_to_date_formula(trends_ws):
    """Col C = (Capacity Actual + Variance) / MAX(1, hours_actuals_through_month), IFERROR-wrapped."""
    header_row = _find_row_containing(trends_ws, 1, "Person")
    assert header_row is not None
    first_person_row = header_row + 1
    cell = trends_ws.cell(first_person_row, 3)
    formula = str(cell.value)
    assert "Capacity!" in formula, f"Col C must reference Capacity: {formula!r}"
    assert "MAX(1,hours_actuals_through_month)" in formula, (
        f"Col C must divide by MAX(1,hours_actuals_through_month): {formula!r}"
    )
    # Plan = Actual + Variance — both Capacity refs must appear, and a '+' between
    assert formula.count("Capacity!") >= 2, (
        f"Col C must sum two Capacity refs (Actual + Variance): {formula!r}"
    )
    assert "IFERROR" in formula


def test_person_load_variance_to_date_pct_formula(trends_ws):
    """Col D = Variance / (Actual + Variance), IFERROR-wrapped, 0.0% format."""
    header_row = _find_row_containing(trends_ws, 1, "Person")
    assert header_row is not None
    first_person_row = header_row + 1
    cell = trends_ws.cell(first_person_row, 4)
    formula = str(cell.value)
    assert "Capacity!" in formula
    assert "IFERROR" in formula, (
        f"Col D must wrap in IFERROR to guard zero plan: {formula!r}"
    )
    # Denominator is a sum of two Capacity refs (Actual + Variance)
    assert formula.count("Capacity!") >= 3, (
        f"Col D must reference Capacity thrice (1 numerator, 2 denominator): "
        f"{formula!r}"
    )
    # Percentage format
    assert cell.number_format == "0.0%", (
        f"Col D number_format must be '0.0%', got {cell.number_format!r}"
    )


def test_person_load_value_cols_are_green(trends_ws):
    """Cols B, C, D are all green cross-sheet refs (GREEN_XREF)."""
    header_row = _find_row_containing(trends_ws, 1, "Person")
    assert header_row is not None
    first_person_row = header_row + 1

    def _font_rgb(cell):
        return str(getattr(cell.font.color, "rgb", "")).upper()

    for col in (2, 3, 4):
        rgb = _font_rgb(trends_ws.cell(first_person_row, col))
        assert rgb.endswith("008000"), (
            f"Col {col} font should be GREEN_XREF (008000), got {rgb}"
        )


def test_person_load_variance_cf_rules(trends_ws):
    """Col D has red (>0.20) and amber (>0.10) CF rules; red priority beats amber."""
    header_row = _find_row_containing(trends_ws, 1, "Person")
    assert header_row is not None
    first_person_row = header_row + 1
    last_person_row = first_person_row + len(TEAM) - 1

    target_range = f"D{first_person_row}:D{last_person_row}"

    # Collect all CF rules applied to that exact range.
    # openpyxl's ConditionalFormatting keys stringify as "<ConditionalFormatting R>" —
    # probe the range by substring match rather than exact str() equality.
    rules = []
    for cf_range in trends_ws.conditional_formatting:
        if target_range not in str(cf_range):
            continue
        for rule in trends_ws.conditional_formatting[cf_range]:
            rules.append(rule)

    # Match by formula content (CF fill color identity is DXF-context-dependent).
    # Red threshold = 0.2 (20%), amber threshold = 0.1 (10%) — match on the
    # exact ">N" suffix to avoid 0.1 substring-matching inside 0.10 etc.
    red_rules = [
        r for r in rules
        if r.formula and "ABS" in r.formula[0] and ">0.2" in r.formula[0]
    ]
    amber_rules = [
        r for r in rules
        if r.formula and "ABS" in r.formula[0] and ">0.1" in r.formula[0]
        and ">0.2" not in r.formula[0]
    ]

    assert len(red_rules) == 1, (
        f"Expected one red (>0.20) CF rule on {target_range}, "
        f"got rules={[r.formula for r in rules]}"
    )
    assert len(amber_rules) == 1, (
        f"Expected one amber (>0.10) CF rule on {target_range}, "
        f"got rules={[r.formula for r in rules]}"
    )

    # Red must come first → lower priority number → higher priority
    red_priority = red_rules[0].priority
    amber_priority = amber_rules[0].priority
    assert red_priority is not None and amber_priority is not None
    assert red_priority < amber_priority, (
        f"Red rule must have higher CF priority (lower number) than amber "
        f"so red wins at >=20%. red_priority={red_priority}, "
        f"amber_priority={amber_priority}"
    )


# ================================================================
# Chart
# ================================================================

def test_cumulative_chart_exists(trends_ws):
    """Trends has at least one chart."""
    assert len(trends_ws._charts) >= 1


def test_chart_has_five_source_rows(trends_ws):
    """Chart data spans 5 helper rows (planned, actual, forecast, ceiling, rp).

    Note: openpyxl expands multi-column References into individual series
    on save/load roundtrip, so we check chart title + that series count
    is a multiple of 5 (one per month column per source row).
    """
    chart = trends_ws._charts[0]
    # After roundtrip the title is a rich Title object; extract text runs
    title_text = "".join(
        run.t for p in chart.title.tx.rich.paragraphs for run in p.r
    )
    assert "Planned vs Actual" in title_text
    # 60 months * 5 rows = 300 series after roundtrip expansion
    assert len(chart.series) >= 5
    assert len(chart.series) % 5 == 0


def test_helper_rows_hidden(trends_ws):
    """Chart helper rows (between person load section and chart) are hidden."""
    # Discover helper rows dynamically: they are the hidden rows in the sheet
    hidden_rows = [
        r for r in range(1, trends_ws.max_row + 1)
        if trends_ws.row_dimensions[r].hidden
    ]
    assert len(hidden_rows) == 8, (
        f"Expected 8 hidden helper rows "
        f"(planned, actual, forecast, ceiling, labels, rp, gap, zero), "
        f"found {len(hidden_rows)}: {hidden_rows}"
    )
    # Verify they are contiguous
    for i in range(1, len(hidden_rows)):
        assert hidden_rows[i] == hidden_rows[i - 1] + 1, (
            f"Hidden helper rows should be contiguous, got {hidden_rows}"
        )


# ================================================================
# Conditional formatting
# ================================================================

# ================================================================
# Dynamic row layout — parametric tests for row collision safety
# ================================================================

def _make_wps(n):
    """Generate n synthetic work packages."""
    return [
        {"id": f"WP{i+1}", "name": f"Work Package {i+1}",
         "fte": 10.0, "start": 1, "end": 60}
        for i in range(n)
    ]


def _make_team(n, work_packages):
    """Generate n synthetic team members, each assigned to WP1."""
    wp_id = work_packages[0]["id"]
    return [
        {"name": f"Person{i+1}", "max_fte": 1.0, "grade": "Mid",
         "wps": {wp_id: 1.0}}
        for i in range(n)
    ]


def _build_trends_only(num_wps: int, num_team: int) -> "Worksheet":
    """Build a workbook with given WP/team counts and return the Trends ws + layout."""
    from openpyxl import Workbook as WB
    from openpyxl.worksheet.worksheet import Worksheet
    from budget_planner.workbook import _build_all_sheets
    from budget_planner.layout import SHEET_TRENDS

    from budget_planner import config as _cfg

    work_packages = _make_wps(num_wps)
    team = _make_team(num_team, work_packages)
    external_commitments: list[dict[str, object]] = []
    project_params = {
        "title": "Test", "start": PROJECT_START,
        "months": 24, "rate": 7500,
    }

    wb = WB()
    _build_all_sheets(wb, project_params, work_packages, team,
                      external_commitments, grade_rates=_cfg.GRADE_RATES)

    trends_ws = wb[SHEET_TRENDS]
    assert isinstance(trends_ws, Worksheet)
    return trends_ws


@pytest.mark.parametrize("num_wps,num_team", [
    (3, 2),    # small — fewer WPs than default
    (9, 4),    # current default config size
    (15, 8),   # large — more WPs and team members
])
def test_trends_no_row_collisions(num_wps, num_team):
    """Sections must not overlap regardless of WP / team count.

    Verifies that:
    - WP data rows don't overlap with chart anchor
    - Chart anchor + 15 (chart height) doesn't overlap person section
    - Person rows don't overlap with hidden helper rows
    - All hidden rows are contiguous and below person data
    """
    ws = _build_trends_only(num_wps, num_team)

    # Discover sections dynamically
    total_row = _find_row(ws, 2, "TOTAL")
    assert total_row is not None, "TOTAL row not found"

    wp_header = _find_row_containing(ws, 1, "WP ID")
    assert wp_header is not None
    first_wp = wp_header + 1
    last_wp = total_row - 1
    assert last_wp - first_wp + 1 == num_wps, (
        f"Expected {num_wps} WP rows, got {last_wp - first_wp + 1}"
    )

    person_section = _find_row_containing(ws, 1, "PERSON LOAD")
    assert person_section is not None, "PERSON LOAD section not found"

    person_header = _find_row_containing(ws, 1, "Person")
    assert person_header is not None
    first_person = person_header + 1

    # Count person data rows
    person_count = 0
    for i in range(50):
        cell = ws.cell(first_person + i, 1)
        if cell.value is None or str(cell.value).strip() == "":
            break
        person_count += 1
    assert person_count == num_team
    last_person = first_person + person_count - 1

    # Find hidden (helper) rows
    hidden_rows = sorted(
        r for r in range(1, ws.max_row + 1)
        if ws.row_dimensions[r].hidden
    )
    assert len(hidden_rows) == 8, (
        f"Expected 8 hidden helper rows, found {len(hidden_rows)}: {hidden_rows}"
    )

    # Verify no overlaps
    # Chart anchor should be after total row
    chart_anchor = total_row + 2  # matches the code
    assert chart_anchor > total_row, "Chart anchor must be after TOTAL row"

    # Person section must start after chart (chart height 7.5 cm ≈ 14 rows)
    assert person_section >= chart_anchor + 14, (
        f"Person section row {person_section} too close to chart at {chart_anchor}. "
        f"Chart needs ~14 rows of height."
    )

    # Gap chart (T13) sits between the cumulative chart and the person
    # section. Pin the gap-chart → person-section margin directly so the
    # reconciliation block can't silently grow into Person Load.
    from openpyxl.utils.cell import coordinate_from_string
    assert len(ws._charts) == 2, (  # type: ignore[attr-defined]
        f"Expected 2 charts (cumulative + gap), got {len(ws._charts)}"  # type: ignore[attr-defined]
    )
    gap_anchor = ws._charts[1].anchor  # type: ignore[attr-defined]
    # openpyxl anchors a chart either as a string like "A123" or as a
    # OneCellAnchor/TwoCellAnchor with `_from.row` (0-indexed).
    if isinstance(gap_anchor, str):
        _, gap_chart_anchor_row = coordinate_from_string(gap_anchor)
    else:
        gap_chart_anchor_row = gap_anchor._from.row + 1
    assert person_section >= gap_chart_anchor_row + 14, (
        f"Person section {person_section} overlaps gap chart at {gap_chart_anchor_row}"
    )

    # Helper rows must be after person data
    assert hidden_rows[0] > last_person, (
        f"First helper row {hidden_rows[0]} overlaps person data ending at {last_person}"
    )

    # Helper rows must be contiguous
    for i in range(1, len(hidden_rows)):
        assert hidden_rows[i] == hidden_rows[i - 1] + 1, (
            f"Hidden helper rows not contiguous: {hidden_rows}"
        )


def test_chart_has_rp_boundary_series():
    """Trends cumulative chart has an RP boundary marker series."""
    ws = _build_trends_only(num_wps=3, num_team=2)

    # Find the cumulative chart (first chart; second is the gap chart)
    assert len(ws._charts) == 2  # type: ignore[attr-defined]
    chart = ws._charts[0]  # type: ignore[attr-defined]

    # Should have 5 series: Planned, Actual, Forecast, Ceiling, RP End
    assert len(chart.series) == 5, (
        f"Expected 5 chart series, got {len(chart.series)}"
    )

    # Last series should be "RP End"
    rp_series = chart.series[4]
    assert rp_series.tx is not None
    # SeriesLabel value should contain "RP"
    label_val = getattr(rp_series.tx, 'v', None)
    assert label_val is not None and "RP" in label_val


# ================================================================
# Cumulative Forecast series (Task 11)
# ================================================================

def test_chart_has_cumulative_forecast_series():
    """LineChart has a 'Cumulative Forecast' series between Actual and Ceiling.

    Order: Planned (idx 0), Actual (idx 1), Forecast (idx 2),
           Ceiling (idx 3), RP End (idx 4).
    """
    ws = _build_trends_only(num_wps=3, num_team=2)
    assert len(ws._charts) == 2  # type: ignore[attr-defined]
    chart = ws._charts[0]  # type: ignore[attr-defined]

    # Forecast is at index 2 (after Actual, before Ceiling)
    forecast = chart.series[2]
    assert forecast.tx is not None
    label_val = getattr(forecast.tx, 'v', None)
    assert label_val == "Cumulative Forecast", (
        f"Expected 'Cumulative Forecast' at series idx 2, got {label_val!r}"
    )


def test_cumulative_forecast_series_styling():
    """Cumulative Forecast series is amber (FFB300) and dashed."""
    ws = _build_trends_only(num_wps=3, num_team=2)
    chart = ws._charts[0]  # type: ignore[attr-defined]
    forecast = chart.series[2]

    line = forecast.graphicalProperties.line
    # solidFill is a ColorChoice after openpyxl normalises the assignment
    color_str = getattr(line.solidFill, "srgbClr", line.solidFill)
    assert color_str == "FFB300", (
        f"Expected amber (FFB300) line color, got {color_str!r}"
    )
    # dashStyle setter writes through to the underlying prstDash element
    dash = line.dashStyle or getattr(line, "prstDash", None)
    assert dash == "dash", (
        f"Expected dashed line, got dashStyle={dash!r}"
    )


def test_cumulative_forecast_helper_row_formulas(trends_ws):
    """Forecast helper row uses the project-aggregated 2-anchor pattern
    (Phase 2.4 / design §9): every cell tests two month-index anchors and
    emits NA() everywhere else, so LineChart draws a single straight
    segment between (eur_actuals_through_month, sum_posted_now) and
    (project_months, sum_forecast_end).

    Each cell shape:
        =IF(<m>=eur_actuals_through_month,<sum_posted_now>,
            IF(<m>=project_months,<sum_forecast_end>,
               NA()))
    """
    # Discover the forecast helper row dynamically: it's the hidden row
    # whose cells reference project_months and Actuals posted_idx/value/
    # slope cells (the 2-anchor formula pattern).
    hidden_rows = sorted(
        r for r in range(1, trends_ws.max_row + 1)
        if trends_ws.row_dimensions[r].hidden
    )

    forecast_row = None
    for r in hidden_rows:
        cell_val = trends_ws.cell(r, MONTH_COL_START).value
        if cell_val is None:
            continue
        s = str(cell_val)
        # Forecast row: anchors on eur_actuals_through_month + project_months,
        # references Actuals (slope-projection sum), and emits NA() on the
        # off-anchor branch.
        if (
            "eur_actuals_through_month" in s
            and "project_months" in s
            and "NA()" in s
            and "Actuals!" in s
        ):
            forecast_row = r
            break
    assert forecast_row is not None, (
        f"No hidden helper row matches the 2-anchor forecast pattern "
        f"(hidden_rows={hidden_rows})"
    )

    # First-month cell: full 2-anchor IF shape with NA() fallthrough.
    first_cell = trends_ws.cell(forecast_row, MONTH_COL_START).value
    assert isinstance(first_cell, str)
    assert "1=eur_actuals_through_month" in first_cell, (
        f"First-month forecast helper should test '1=eur_actuals_through_month', "
        f"got {first_cell!r}"
    )
    assert "1=project_months" in first_cell, (
        f"First-month forecast helper should test '1=project_months' "
        f"(future-anchor branch), got {first_cell!r}"
    )
    assert "NA()" in first_cell, (
        f"First-month forecast helper should emit NA() on off-anchor cells, "
        f"got {first_cell!r}"
    )
    # Future anchor must reference the slope projection.
    assert "MAX(0," in first_cell, (
        f"Future-anchor sum must clamp at 0 via MAX(0, ...): {first_cell!r}"
    )
    # No more Plan-Forecast refs after Phase 2.4.
    assert "Plan-Forecast" not in first_cell, (
        f"Forecast helper must not reference Plan-Forecast: {first_cell!r}"
    )

    # Mid-month cell at m_idx=4 (literal m=5): same 2-anchor shape with
    # the literal updated.
    mid_cell = trends_ws.cell(forecast_row, MONTH_COL_START + 4).value
    assert isinstance(mid_cell, str)
    assert "5=eur_actuals_through_month" in mid_cell, (
        f"Mid-month (m=5) forecast helper should test "
        f"'5=eur_actuals_through_month', got {mid_cell!r}"
    )
    assert "5=project_months" in mid_cell
    assert "NA()" in mid_cell


# ================================================================
# v2 T10: Actual / Forecast helper-row clipping past through-month
# ================================================================

def _find_helper_row_by_probe(ws, substring):
    """Return first hidden row whose first-month cell contains `substring`."""
    hidden_rows = sorted(
        r for r in range(1, ws.max_row + 1)
        if ws.row_dimensions[r].hidden
    )
    for r in hidden_rows:
        v = ws.cell(r, MONTH_COL_START).value
        if v is not None and substring in str(v):
            return r
    return None


def test_actual_helper_row_first_month_clipped(trends_ws):
    """First-month Actual helper formula references eur_actuals_through_month
    and NA() — the v2 T10 clipping wrapper.
    """
    # The Actual helper row is the hidden row whose first-month formula
    # contains 'Actuals!' but not 'Plan-Forecast' (distinguishes from forecast
    # row) and not 'partner_ceiling' (distinguishes from ceiling row).
    hidden_rows = sorted(
        r for r in range(1, trends_ws.max_row + 1)
        if trends_ws.row_dimensions[r].hidden
    )
    actual_row = None
    for r in hidden_rows:
        v = trends_ws.cell(r, MONTH_COL_START).value
        if v is None:
            continue
        s = str(v)
        if ("Actuals!" in s and "'Plan-Forecast'" not in s
                and "partner_ceiling" not in s
                and "SUMPRODUCT" not in s):
            actual_row = r
            break
    assert actual_row is not None, (
        f"No hidden helper row with an Actuals! ref found (hidden={hidden_rows})"
    )

    first_cell = trends_ws.cell(actual_row, MONTH_COL_START).value
    assert isinstance(first_cell, str)
    assert "eur_actuals_through_month" in first_cell, (
        f"First-month Actual formula must reference eur_actuals_through_month, "
        f"got {first_cell!r}"
    )
    assert "NA()" in first_cell, (
        f"First-month Actual formula must clip with NA(), got {first_cell!r}"
    )
    # Literal month index 1 for m_idx=0
    assert "1>eur_actuals_through_month" in first_cell, (
        f"First-month Actual formula should test '1>eur_actuals_through_month', "
        f"got {first_cell!r}"
    )


def test_actual_helper_row_mid_month_clipped(trends_ws):
    """Actual helper-row formula at m_idx=5 uses literal '6>eur_actuals_through_month'."""
    hidden_rows = sorted(
        r for r in range(1, trends_ws.max_row + 1)
        if trends_ws.row_dimensions[r].hidden
    )
    actual_row = None
    for r in hidden_rows:
        v = trends_ws.cell(r, MONTH_COL_START).value
        if v is None:
            continue
        s = str(v)
        if ("Actuals!" in s and "'Plan-Forecast'" not in s
                and "partner_ceiling" not in s
                and "SUMPRODUCT" not in s):
            actual_row = r
            break
    assert actual_row is not None

    # m_idx = 5 -> column = MONTH_COL_START + 5, literal month idx = 6.
    mid_cell = trends_ws.cell(actual_row, MONTH_COL_START + 5).value
    assert isinstance(mid_cell, str)
    assert "6>eur_actuals_through_month" in mid_cell, (
        f"Mid-month (m_idx=5) Actual formula should test "
        f"'6>eur_actuals_through_month', got {mid_cell!r}"
    )
    assert "NA()" in mid_cell
    # Cumulative shape preserved inside the IF's false-branch.
    assert "+Actuals!" in mid_cell, (
        f"Mid-month Actual formula should retain cumulative "
        f"'<prev>+Actuals!...' shape, got {mid_cell!r}"
    )


def test_forecast_helper_row_two_anchor_pattern(trends_ws):
    """Forecast helper row no longer clips by hours_actuals_through_month
    after Phase 2.4. The new 2-anchor pattern (design §9) anchors on
    eur_actuals_through_month (now-anchor) and project_months
    (future-anchor); off-anchor cells emit NA(). The hours-side cadence
    is no longer consulted because the forecast is sourced from Actuals
    EUR posted-cumulative cells (not Plan-Forecast hours derivations).
    """
    # Locate the forecast helper row by its 2-anchor shape (project_months
    # + eur_actuals_through_month + NA() + Actuals!).
    forecast_row = _find_helper_row_by_probe(trends_ws, "project_months")
    assert forecast_row is not None

    # m_idx=5 -> literal m=6: both anchors must use the literal '6=...' form.
    mid_cell = trends_ws.cell(forecast_row, MONTH_COL_START + 5).value
    assert isinstance(mid_cell, str)
    assert "6=eur_actuals_through_month" in mid_cell, (
        f"Mid-month (m_idx=5) forecast formula should test "
        f"'6=eur_actuals_through_month', got {mid_cell!r}"
    )
    assert "6=project_months" in mid_cell, (
        f"Mid-month forecast formula should test '6=project_months' "
        f"(future-anchor branch), got {mid_cell!r}"
    )
    assert "NA()" in mid_cell
    # hours_actuals_through_month is no longer the forecast trust boundary.
    assert "hours_actuals_through_month" not in mid_cell, (
        f"Forecast helper must not reference hours_actuals_through_month "
        f"after Phase 2.4 (forecast sources from Actuals EUR slope, not "
        f"hours-derived Plan-Forecast): {mid_cell!r}"
    )


def test_planned_and_ceiling_helper_rows_unchanged(trends_ws):
    """Planned + Ceiling helper rows must NOT contain NA() — only Actual and
    Forecast are clipped by v2 T10.
    """
    hidden_rows = sorted(
        r for r in range(1, trends_ws.max_row + 1)
        if trends_ws.row_dimensions[r].hidden
    )
    # Planned: contains SUMPRODUCT (and not Plan-Forecast/Actuals/partner_ceiling alone)
    planned_row = None
    ceiling_row = None
    for r in hidden_rows:
        v = trends_ws.cell(r, MONTH_COL_START).value
        if v is None:
            continue
        s = str(v)
        if "SUMPRODUCT" in s and planned_row is None:
            planned_row = r
        elif s.strip() == "=partner_ceiling" and ceiling_row is None:
            ceiling_row = r
    assert planned_row is not None, (
        f"Could not find planned helper row (hidden={hidden_rows})"
    )
    assert ceiling_row is not None, (
        f"Could not find ceiling helper row (hidden={hidden_rows})"
    )

    # Probe a few month columns on each.
    for m_idx in (0, 3, 7):
        c = MONTH_COL_START + m_idx
        pv = trends_ws.cell(planned_row, c).value
        cv = trends_ws.cell(ceiling_row, c).value
        assert pv is None or "NA()" not in str(pv), (
            f"Planned helper row {planned_row} col {c} unexpectedly clipped: {pv!r}"
        )
        assert cv is None or "NA()" not in str(cv), (
            f"Ceiling helper row {ceiling_row} col {c} unexpectedly clipped: {cv!r}"
        )
        assert pv is None or "eur_actuals_through_month" not in str(pv), (
            f"Planned row must not reference eur_actuals_through_month, got {pv!r}"
        )
        assert pv is None or "hours_actuals_through_month" not in str(pv), (
            f"Planned row must not reference hours_actuals_through_month, got {pv!r}"
        )
