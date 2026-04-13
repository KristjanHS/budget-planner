"""Build the _META sheet — AI-agent-readable workbook documentation."""
from openpyxl.styles import Font
from datetime import datetime

from budget_planner.layout import (
    MONTH_COL_START,
    month_col_end as _month_col_end,
    val_columns,
    CAPACITY_AMBER_LOAD,
    CAPACITY_VARIANCE_FTE_RED,
)
from budget_planner.styles import WRAP, col_letter


def build_meta_sheet(ws, layout, *, project_params, work_packages,
                     team, external_commitments, config_layout, summary_layout,
                     grade_rates, actuals_hours_layout=None,
                     plan_forecast_layout=None, rebuild_warnings=None):
    """Populate the _META worksheet with structural documentation.

    Every data parameter is required: _META must document the **live**
    workbook state, never config.py seed data. Callers must pass the same
    values they handed to the other builders.

    Parameters
    ----------
    ws : openpyxl.worksheet.worksheet.Worksheet
        The _META sheet (already created by the caller).
    layout : dict
        Row-number mapping produced by the plan-sheet builder. Keys used:
        config_first_wp_row, config_last_wp_row, planned_header_row,
        planned_first_wp_row, planned_last_wp_row, planned_total_row.
    project_params : dict
        Dict with keys title, start, months, rate.
    work_packages : list
        List of WP dicts.
    team : list
        List of team member dicts.
    external_commitments : list
        List of external commitment dicts.
    config_layout : dict
        Layout dict from build_config_sheet(). Keys used:
        grade_first_row, grade_last_row, roster_header_row, roster_first_row,
        roster_last_row.
    summary_layout : dict
        Layout dict from build_summary_sheet(). Keys used:
        first_wp_row, last_wp_row, total_row.
    grade_rates : dict
        Mapping of grade name to EUR/PM rate.
    actuals_hours_layout : dict or None
        Layout dict from build_actuals_hours_sheet(). When provided, a
        dedicated Actuals-Hours section is rendered. Partial doc in Stage 1
        (hours-timelog v1); full rewrite lands in Task 15.
    plan_forecast_layout : dict or None
        Layout dict from build_plan_forecast_sheet(). When provided, a
        dedicated Plan-Forecast section is rendered. Partial doc in Stage 1.
    rebuild_warnings : list or None
        Optional list of warning strings to append to meta text.
    """

    month_col_start = MONTH_COL_START
    month_col_end = _month_col_end(project_params["months"])
    _vcols = val_columns(project_params["months"])
    val_col_sum = _vcols["sum"]
    val_col_budget = _vcols["budget"]
    val_col_diff = _vcols["diff"]
    val_col_cum = _vcols["cum"]
    val_col_pct_ceil = _vcols["pct_ceil"]
    val_col_breach = _vcols["breach"]

    # Derive row numbers from layout dict
    plan_header = layout["planned_header_row"]
    plan_year_row = layout["planned_header_row"] + 1
    plan_month_row = layout["planned_header_row"] + 2
    plan_first_wp = layout["planned_first_wp_row"]
    plan_last_wp = layout["planned_last_wp_row"]
    plan_total = layout["planned_total_row"]

    # Team and WP counts (needed early for row derivations)
    num_team = len(team)
    num_ext = len(external_commitments)
    num_wps = len(work_packages)

    # Config sheet row positions
    cfg_first_wp = layout["config_first_wp_row"]
    cfg_last_wp = layout["config_last_wp_row"]
    cfg_total_row = cfg_last_wp + 1

    # Grade table rows (config_layout is always provided by the orchestrator
    # — T6 removed the lazy-import fallback and made it required in spirit;
    # T7 removes the last None branch by threading live layouts everywhere.)
    cfg_grade_first = config_layout["grade_first_row"]
    cfg_grade_last = config_layout["grade_last_row"]
    cfg_roster_header = config_layout["roster_header_row"]
    cfg_roster_first = config_layout["roster_first_row"]
    cfg_grade_section = cfg_grade_first - 2  # section header is 2 rows above first data
    cfg_grade_header = cfg_grade_first - 1
    cfg_wp_section = cfg_first_wp - 2  # WP section header row
    cfg_wp_header = cfg_first_wp - 1   # WP column header row
    cfg_roster_section = cfg_roster_header - 1
    cfg_roster_last = cfg_roster_first + num_team - 1
    num_grades = cfg_grade_last - cfg_grade_first + 1
    grade_names_str = ", ".join(grade_rates.keys())

    # Summary sheet references — from live summary_layout, not hardcoded.
    sum_first_wp = summary_layout["first_wp_row"]
    sum_last_wp = summary_layout["last_wp_row"]
    sum_total = summary_layout["total_row"]

    # ── Stage 1 hours-timelog layouts (optional until Task 15 rewrite) ──
    # When the orchestrator threads these through, emit live row numbers;
    # otherwise fall back to placeholder dashes so the doc still renders.
    if actuals_hours_layout is not None:
        ah_band_label = str(actuals_hours_layout["band_label_row"])
        ah_year_header = str(actuals_hours_layout["year_header_row"])
        ah_month_header = str(actuals_hours_layout["month_header_row"])
        # First person header row sits directly below the month header.
        ah_first_person_header = str(
            actuals_hours_layout["month_header_row"] + 1
        )
    else:
        ah_band_label = ah_year_header = ah_month_header = "—"
        ah_first_person_header = "—"

    if plan_forecast_layout is not None:
        pf_forecast_header = str(plan_forecast_layout["forecast_header_row"])
        pf_first_wp_row = str(plan_forecast_layout["forecast_first_wp_row"])
        pf_last_wp_row = str(plan_forecast_layout["forecast_last_wp_row"])
        pf_total_row = str(plan_forecast_layout["forecast_total_row"])
    else:
        pf_forecast_header = pf_first_wp_row = "—"
        pf_last_wp_row = pf_total_row = "—"

    # Column letters
    first_mcol = col_letter(month_col_start)
    last_mcol = col_letter(month_col_end)
    val_sum_col = col_letter(val_col_sum)
    val_budget_col = col_letter(val_col_budget)
    val_diff_col = col_letter(val_col_diff)
    val_cum_col = col_letter(val_col_cum)
    val_pct_col = col_letter(val_col_pct_ceil)
    val_breach_col = col_letter(val_col_breach)

    meta_text = """WORKBOOK: Programme Project Budget Planner
VERSION: 4.0
CREATED: {date}
PURPOSE: Track planned FTE allocation and actual EUR spend per work package per month for an EU-funded (programme) project. Per-person capacity planning with cross-project visibility. Derive planned euro amounts from FTE using a configurable rate. Record actual EUR from accountant reports on a separate sheet.

═══ SHEET INDEX ═══
1. _META          — This sheet. AI-agent-readable documentation of workbook structure, conventions, and modification instructions.
2. Config         — Single source of truth: project parameters + WP budget + person-WP assignments + external commitments + WP schedule + team roster + reporting periods + grade rate table.
3. Plan           — Compact 1-row header + Planned FTE block (formulas from Capacity).
4. Capacity       — Primary planning surface: per-person monthly FTE allocation across programme WPs and external projects (uses assignments defined on Config).
5. Actuals        — EUR actuals per WP per month (input from accountant) + parallel forecast monitoring block (reads Plan-Forecast). Reconciliation visual: monthly Actual EUR cells past `eur_actuals_through_month` are greyed (CF) — those months are not closed yet. Tab color: Blue.
6. Actuals-Hours  — Per-person monthly hour timelog + sibling forecast-EUR band (outline_level=1, collapsible). Sole source of forecast-EUR leaves. Tab color: Red (C00000). Named ranges used: hours_per_fte_month, eu_annual_ceiling_pm, roster_lookup, eur_actuals_through_month.
                    Data flow IN  ← (none — primary input surface for hours)
                    Data flow OUT → Plan-Forecast (leaves SUMmed per WP × month) → Actuals forecast block → Capacity Actual-FTE band → Summary Forecast EUR/year → Trends Cumulative Forecast.
7. Plan-Forecast  — WP × month forecast-EUR grid aggregating the leaf formulas on Actuals-Hours. Mirrors Plan shape. Tab color: Green (00B050).
                    Data flow IN  ← Actuals-Hours forecast-EUR leaves (SUM per WP × month).
                    Data flow OUT → Actuals forecast monitoring block, Summary Forecast EUR/year, Trends Cumulative Forecast helper row.
8. Summary        — Calendar-year aggregation. Plan FTE, Plan EUR, Actual EUR, Forecast EUR, Variance EUR per WP per year (5 cols/year) + project total group. Budget overview + Form C reporting-period block.
9. Trends         — WP burn rate & exhaustion, cumulative EUR chart (Planned + Actual + Cumulative Forecast + Ceiling), person load (planned + actual-side cols). Tab color: Gold (FFD700).

═══ CONFIG SHEET STRUCTURE ═══

SHEET: CONFIG (single source of truth for all project parameters)
  Tab color: Blue (4472C4)
  Freeze panes: A{cfg_wp_section}

  SECTION A: PROJECT PARAMETERS (rows 1-5)
    Row 1: A1="Project Title",            B1=<title>           (blue input)
    Row 2: A2="Project Start",            B2=<date>            (blue input, MMM YYYY)
    Row 3: A3="Project Months",           B3=<integer>         (blue input)
    Row 4: A4="FTE Rate (EUR)",           B4=<formula>         (black formula, €#,##0)
           Weighted average: =SUMPRODUCT(IFERROR(VLOOKUP(roster_grades,tbl_GradeRates,2,FALSE),0)*roster_ftes)/SUM(roster_ftes)
           Derives rate from per-grade rates weighted by team FTE. Not directly editable.
    Row 5: A5="EU Annual Ceiling (PM/FTE/year)", B5=<ceiling> (blue input, 0.00; default 12.0 ≈ 215 days/year)
    Row 6: (blank separator)

  SECTION B: WP BUDGET TABLE (rows {cfg_wp_section}-{cfg_total_row})
    Row {cfg_wp_section}: Section header "WP BUDGET" (merged A:G, section fill)
    Row {cfg_wp_header}: Column headers: WP ID | WP Name | FTE Budget | Rate (€/PM) | Ceiling (€) | Assigned PM | Variance
    Rows {cfg_first_wp}-{cfg_last_wp}: WP data rows
      Cols A-C: blue input (WP ID, Name, FTE Budget)
      Col D: Rate (€/PM) — blue input, per-WP rate (default 6500)
      Col E: Ceiling (€) — black formula =C*D (FTE × Rate)
      Col F: Assigned PM — black formula =SUMIFS(assign_ftes,assign_wps,A<row>)
      Col G: Variance — black formula =C<row>-F<row> (CF yellow when |G| > 0.01)
    Row {cfg_total_row}: TOTAL row (SUM on cols C, E, F, G; light blue fill)
    Excel Table "tbl_WorkPackages" covers A{cfg_wp_header}:G{cfg_last_wp} (header + data, excludes TOTAL)

  SECTION B2: PERSON-WP ASSIGNMENTS (below WP BUDGET)
    Section header "PERSON-WP ASSIGNMENTS" (merged A:H, section fill)
    Column headers: Person | WP | Total FTE | Start Month | End Month | Sig Assigned | Sig External | Balance
    One row per (person, WP) pair allocated on the project.
      Cols A-E: blue input (Person, WP, Total FTE, Start Month, End Month)
      Col F: Sig Assigned — black formula =SUMIFS(assign_ftes,assign_persons,A<row>)
      Col G: Sig External — black formula =SUMIFS(ext_ftes,ext_persons,A<row>)
      Col H: Balance — black formula
        =VLOOKUP(A<row>,roster_lookup,2,FALSE)*project_months - F<row> - G<row>
    Cols F-H repeat per row (intentional — in-place context per assignment).
    Named ranges: assign_persons, assign_wps, assign_ftes.
    Conditional formatting on Balance: red if < 0 (over-committed),
    yellow if < 10% of available PM (low remaining capacity).

  SECTION B3: EXTERNAL COMMITMENTS (below PERSON-WP ASSIGNMENTS)
    Section header "EXTERNAL COMMITMENTS" (merged A:F, section fill)
    Column headers: Person | Project | Total FTE | Start Month | End Month | EU Grant
    All cells blue input. {num_ext} entries (from config on init, preserved on rebuild).
    Named ranges: ext_persons, ext_ftes.

  SECTION C: WP SCHEDULE TABLE (below EXTERNAL COMMITMENTS)
    Section header "WP SCHEDULE" (merged A:D, section fill)
    Column headers: WP ID | Start | End | Type
    Multiple rows per WP allowed (for multi-range active periods).
    A month is active if it falls within ANY schedule range for that WP.
      Col A: WP ID (blue input, can repeat for multiple ranges)
      Col B: Start month (blue input, format 0)
      Col C: End month (blue input, format 0)
      Col D: Type — distribution curve (blue input, dropdown: flat/front-loaded/bell/back-loaded)
    Named ranges: sched_wp_ids, sched_starts, sched_ends (for MINIFS/MAXIFS in Trends formulas)
    Conditional formatting: yellow fill if Start > End

  SECTION D: TEAM ROSTER (below WP SCHEDULE)
    Row {cfg_roster_section}: Section header "TEAM ROSTER" (merged A:C, section fill)
    Row {cfg_roster_header}: Column headers: Person | Max FTE | Grade
    Rows {cfg_roster_first}+: Team member data rows (blue input, {num_team} people)
      Col A: Person name (blue input)
      Col B: Max FTE (blue input, 0.00)
      Col C: Grade (blue input, data validation dropdown from grade names)
    Named range roster_ftes covers B{cfg_roster_first}:B{cfg_roster_last}
    Named range roster_grades covers C{cfg_roster_first}:C{cfg_roster_last}

  SECTION E: REPORTING PERIODS (below Team Roster)
    Col A: Period ID (e.g. RP1) — blue input
    Col B: Start Month — blue input
    Col C: End Month — blue input
    User-editable: add/remove/modify periods for Form C reporting boundaries.

  SECTION F: RATE BY GRADE (bottom of sheet, rows {cfg_grade_section}-{cfg_grade_last})
    Row {cfg_grade_section}: Section header "RATE BY GRADE" (merged A:B, section fill)
    Row {cfg_grade_header}: Column headers: Grade | Rate (EUR/PM)
    Rows {cfg_grade_first}-{cfg_grade_last}: Grade data rows ({num_grades} grades: {grade_names_str})
      Col A: Grade name (blue input)
      Col B: Rate EUR/PM (blue input, €#,##0)
    Named range tbl_GradeRates covers A{cfg_grade_first}:B{cfg_grade_last}
    Purpose: per-grade personnel rates used to compute weighted-average fte_rate formula.

═══ CAPACITY SHEET STRUCTURE ═══

SHEET: CAPACITY (primary planning surface)
  Tab color: Purple

  Assignments and external commitments live on Config (sections B2/B3);
  Capacity is purely the monthly planning grid. Per-row Budget formulas
  reference the Config row via =Config!C<assign_row>.

  ROW 1: YEAR HEADER
  ROW 2: MONTH HEADER (calendar year grouping, same layout as Plan/Actuals)

  PER-PERSON GROUPS: one group per team member ({num_team} people)
    Person Header: bold, section fill, formula ref to Config roster (=Config!A<roster_row>)
    Programme WP Rows: only ASSIGNED WPs for this person (from Config PERSON-WP ASSIGNMENTS, not all WPs)
      WP ID/Name formula refs to Config WP table
      Monthly FTE: static values auto-filled by distribution curves (blue input)
      Distribution type per WP: flat, front-loaded, bell, back-loaded (from Config WP SCHEDULE Type col)
      Inactive months: grey fill, empty — determined by intersection of the assignment
        window [Start, End] with WP SCHEDULE active months. WPs with multiple SCHEDULE
        entries (gaps) have all gap months greyed out even if inside the assignment window.
        Total FTE is distributed across the valid months only, never leaked into gaps.
      Users can override individual month cells post-generation for fine-tuning
    External Project Rows: project name (input), monthly FTE (flat distribution, blue input)
      Config EXTERNAL COMMITMENTS section has "EU Grant" column (Yes/No) marking EU-funded externals
    TOTAL Row: SUM of all rows above (formula)
    REMAINING Row: Config Max FTE - TOTAL (formula, refs Config!B<roster_row>)
    EU TOTAL Row: SUM of programme WP rows + EU-flagged external rows (light blue fill)
      Tracks total EU commitment per person per month across all EU grants

  CONDITIONAL FORMATTING:
    TOTAL > Max FTE: red fill (overcommitted)
    TOTAL > 0.8 x Max FTE: amber fill (approaching limit)
    REMAINING < 0: red fill
    EU TOTAL per calendar year > EU Annual Ceiling: red fill (exceeds 215-day limit)
    EU TOTAL per calendar year > EU Annual Ceiling - 1: amber fill (approaching limit)

  VALIDATION COLUMNS: {val_sum_col} = Sum, {val_budget_col} = Budget (from Config PERSON-WP ASSIGNMENTS Total FTE via =Config!C<asgn_row>; falls back to =Config!C<wp_row> FTE Budget if no matching assignment row exists), {val_diff_col} = Variance (per row)

  RIGHT-SIDE GROUPED BAND — ACTUAL FTE / VARIANCE FTE / UTILISATION % (T9):
    Layout (offsets relative to {val_diff_col}):
      Actual-FTE band:   cols [{val_diff_col}+2 .. {val_diff_col}+1+months]   (1-col gutter after Variance)
      Variance-FTE band: cols [{val_diff_col}+months+3 .. {val_diff_col}+2+2*months] (1-col gutter after Actual)
    Both bands are outline_level=1 (collapsible via summaryRight outline button).
    Row 1: merged band labels "ACTUAL FTE (from hours)" and "VARIANCE FTE (Planned − Actual)".
    Year + month headers mirrored from the planned band.
    Per WP/external row, Actual FTE cell = green cross-ref to Actuals-Hours
      sibling hour cell, divided by hours_per_fte_month:
        =IF(<hours-cell>="","",<hours-cell>/hours_per_fte_month)
      External rows: Actual FTE band stays empty (non-programme projects).
    Per WP/external row, Variance FTE cell = Planned − Actual (black formula).
    TOTAL row: Actual band sums the WP+external rows; Variance band shows
      Planned TOTAL − Actual TOTAL.
    CONDITIONAL FORMATTING:
      - Variance FTE: red fill if |variance| > {variance_red_fte} (mismatch alert)
      - Utilisation on Actual TOTAL row vs Config Max FTE:
          red ≥ Max FTE (over-utilised)
          amber ≥ {capacity_amber_load:.0%} × Max FTE (approaching limit)
      - EU annual ceiling check on Actual band: red if PM-equivalent
        annual sum > eu_annual_ceiling_pm (mirrors planned-band ceiling check).

═══ PLAN SHEET STRUCTURE ═══

ROW 1: COMPACT HEADER
  A1 = "Title:", B1 = Project title (green cross-ref to Config!B1, bold 14pt)
  D1 = "Start:", E1 = Project start date (green cross-ref to Config!B2, MMM YYYY)
  F1 = "Ceiling:", G1 = Personnel Ceiling (formula: =SUM(Config!E{cfg_first_wp}:E{cfg_last_wp}), €#,##0)
       G1 also serves as the partner_ceiling named range target.

ROW 2: (blank separator)

PLANNED FTE (Rows {plan_header}-{plan_total})
  Row {plan_header} = Section header "PLANNED FTE (PERSON-MONTHS)" (merged, section fill)
  Row {plan_year_row} = Calendar year labels (merged per calendar year: "2026", "2027", ...)
  Row {plan_month_row} = Month labels (M1 Nov 26, M2 Dec 26, ..., M{pm}) in columns {first_mcol}..{last_mcol}
  Rows {plan_first_wp}-{plan_last_wp} = Monthly FTE per WP ({num_wps} WPs)
    Col A = WP ID (green cross-ref: =Config!A<config_row>)
    Col B = WP Name (green cross-ref: =Config!B<config_row>)
    Cols {first_mcol}-{last_mcol} = SUM formulas referencing Capacity sheet (black text, formula)
      - Each cell: =SUM(Capacity!<col><person1_wp_row>, ..., Capacity!<col><personN_wp_row>)
      - NOT directly editable — edit Capacity sheet instead
      - Inactive months: empty + grey fill
    Col {val_sum_col} = Sum of monthly values (formula: SUM)
    Col {val_budget_col} = Budget FTE from Config col C (green cross-ref: =Config!C<config_row>)
    Col {val_diff_col} = Variance: Budget - Sum (formula). Zero = OK, nonzero = WARNING.
  Row {plan_total} = "MONTHLY TOTAL" — column totals (SUM formula per month + per validation col)
  Conditional formatting: yellow fill on variance if ABS > 0.01
  Freeze panes: C{plan_first_wp}

═══ ACTUALS SHEET STRUCTURE ═══

SHEET: ACTUALS (EUR actuals from accountant)
  Tab color: Blue (0000FF)

  Row 1 = Section header "ACTUAL EUR" (merged, section fill)
  Row 2 = Calendar year labels (merged per calendar year, same layout as Plan)
  Row 3 = Month headers (M1..M{pm}) in columns {first_mcol}..{last_mcol}
         Also has "WP ID"/"WP Name" labels in cols A/B and validation col labels.
  Rows 4-{act_last_wp} = One row per WP ({num_wps} WPs)
    Col A = WP ID (green cross-ref: =Config!A<config_row>)
    Col B = WP Name (green cross-ref: =Config!B<config_row>)
    Cols {first_mcol}-{last_mcol} = EUR amount per month (input, blue text, €#,##0)
      - Active months initialized to 0 (or preserved from rebuild)
      - Inactive months greyed out and empty (same WP active period as Config)
    Col {val_sum_col} = Sum of monthly EUR (formula: SUM)
    Col {val_budget_col} = Ceiling EUR from Config col E (green cross-ref: =Config!E<config_row>)
    Col {val_diff_col} = Variance: Ceiling - Sum (formula). Negative = OVERRUN.
    Col {val_cum_col} = Cum. EUR: SUMPRODUCT cumulative up to current_month (formula)
    Col {val_pct_col} = % Ceiling: Cum. EUR / Ceiling (formula, 0.0% format)
    Col {val_breach_col} = Breach In: months until ceiling exhausted at current rate (formula)
  Row {act_total} = TOTAL row (light green fill):
    Monthly columns = SUM per column across WP rows
    {val_sum_col} = SUM of WP sums
    {val_budget_col} = partner_ceiling named range (total project ceiling)
    {val_diff_col} = Ceiling - Sum
    {val_cum_col}, {val_pct_col}, {val_breach_col} = same formulas on TOTAL row

  PARALLEL FORECAST MONITORING BLOCK (to the right of Actual-EUR monitoring cols):
    4 columns — Cum. Forecast EUR, Forecast Variance, Forecast % Ceiling,
    Forecast Breach In — separated from the Actual-EUR block by a 1-col
    gutter. The block is outline-grouped (column outline_level = 1) so users
    can collapse it. Values are green cross-refs reading Plan-Forecast rows
    (matching WP order). CF rules mirror the Actual-EUR block:
      - Red fill on negative Forecast Variance
      - Forecast % Ceiling: amber >{actuals_amber_pct}%, red >100%
      - Forecast Breach In: amber <{actuals_amber_breach} months, red <{actuals_red_breach} months
    Source of truth: Plan-Forecast totals → Actuals forecast block → Summary/Trends.

  RECONCILIATION VISUAL (T8 — eur_actuals_through_month):
    The named cell `eur_actuals_through_month` (Config!$B$<row>, blue input,
    integer) records the last month for which the accountant has closed Actual
    EUR. Months past that index are not yet reconciled. CF rule on the Actual-
    EUR monthly grid (rows 4-{act_total}, cols {first_mcol}-{last_mcol}):
        =COLUMN(<cell>)-COLUMN($<first_month_col>$1)+1 > eur_actuals_through_month
        → grey fill (CF_GREY_FILL).
    Mirror rule on Actuals-Hours forecast band greys the OPPOSITE side
    (months ≤ eur_actuals_through_month) so users see at a glance which side
    is authoritative for each month.

  Conditional formatting:
    - Red fill on negative variance
    - % Ceiling: amber >{actuals_amber_pct}%, red >100%
    - Breach In: amber <{actuals_amber_breach} months, red <{actuals_red_breach} months
    - Grey fill on Actual-EUR cells past `eur_actuals_through_month` (T8)
  Freeze panes: C4

═══ ACTUALS-HOURS SHEET STRUCTURE ═══

SHEET: ACTUALS-HOURS (per-person monthly hour timelog — primary hours input surface)
  Tab color: Red (C00000) — distinct to flag this as the hours input sheet.
  Freeze panes: C{ah_first_person_header}

  PURPOSE: Log actual hours worked per person × WP (and per person ×
  external project) per month. A sibling forecast-EUR band (one col per
  month, offset right of the hours grid by a gutter) derives EUR from hours
  via the rate lookup. These forecast-EUR leaves are the only authoritative
  source for Plan-Forecast and Actuals forecast monitoring.

  HEADER ROWS:
    Band label row (row {ah_band_label}): Merged "FORECAST EUR (from hours)"
      over the forecast band columns.
    Year header row (row {ah_year_header}): Calendar year labels (same layout
      as Plan/Actuals), mirrored across the forecast band.
    Month header row (row {ah_month_header}): M1..M{pm} labels, mirrored across
      the forecast band.
    Validation header: "Sum" column only (no Budget / no Variance here —
    no row-level ceiling, by design).

  PER-PERSON GROUPS: one group per team member ({num_team} people).
    Person header row: bold, section fill, literal person name in col A
      (used as VLOOKUP key by the forecast-EUR leaves; must be a literal
      so data_only workbook reads evaluate the EUR chain correctly).
    WP rows: one per (person, WP) assignment from Config PERSON-WP ASSIGNMENTS
      (intersection — NOT all WPs for all people).
    External project rows: one per external commitment for this person.
    TOTAL row: SUM of all rows above (hours + forecast EUR band).
    EU TOTAL row: SUM of programme WP rows + EU-flagged external rows.
      NO REMAINING row (by design — no per-row ceiling). The annual hours
      ceiling is enforced by CF against `eu_annual_ceiling_pm * hours_per_fte_month`.

  HOUR INPUT CONVENTION:
    - Blue input font, number format #,##0 (whole hours).
    - Active months only: inactive months (outside WP schedule or outside
      assignment window) are greyed + empty — same rule as Capacity.

  FORECAST-EUR BAND (outline level 1, collapsible):
    One column per project month, offset by a 1-col gutter right of the Sum
    column. Each leaf cell:
      =IF(hour="", "", ROUND(hour / hours_per_fte_month * VLOOKUP(person, roster_lookup, 4, FALSE), 0))
    ROUND wraps each leaf to whole EUR — avoids sub-cent drift when
    Plan-Forecast SUMs aggregate hundreds of leaves.
    Green cross-ref font. Per-person TOTAL / EU TOTAL rows SUM across these
    leaves to give per-person forecast EUR.

  CONDITIONAL FORMATTING (annual EU hours ceiling check):
    Red:   annual hour sum > eu_annual_ceiling_pm * hours_per_fte_month
    Amber: within one hours_per_fte_month unit of the ceiling

  NAMED RANGES USED: hours_per_fte_month, eu_annual_ceiling_pm, roster_lookup.

═══ PLAN-FORECAST SHEET STRUCTURE ═══

SHEET: PLAN-FORECAST (WP × month forecast-EUR grid — aggregation only)
  Tab color: Green (00B050)
  Freeze panes: C{pf_first_wp_row}

  PURPOSE: Aggregate the per-person forecast-EUR leaves on Actuals-Hours
  into the same WP × month shape as Plan. Read-only — all cells are formulas.

  Row 1: Compact header (Title / Start / Ceiling) — same layout as Plan row 1.
  Row {pf_forecast_header}: Section header "FORECAST EUR (from hours)".
  Year / month header rows: calendar years + M1..M{pm} (same layout as Plan).

  WP ROWS (rows {pf_first_wp_row}-{pf_last_wp_row}): one per WP ({num_wps} WPs).
    Col A: WP ID    = green cross-ref to Config.
    Col B: WP Name  = green cross-ref to Config.
    Monthly cells  =SUM(<Actuals-Hours forecast-EUR leaves for this WP × month>)
                     across every (person, WP) assignment. Green font.
    Inactive months: greyed + empty (same active rule as Actuals/Plan).
    Validation block (Sum / Ceiling / Variance):
      Sum     = SUM of monthly cells on this row.
      Ceiling = green cross-ref to Config col E (per-WP ceiling).
      Variance = Ceiling - Sum. Red fill when negative (forecast exceeds ceiling).

  Row {pf_total_row} = TOTAL (light green fill): SUM across WP rows per month
    plus totals of the three validation columns.

═══ ACTUALS-HOURS / PLAN-FORECAST DATA FLOW ═══

  Hours input (blue cells on Actuals-Hours)
    → per-cell forecast-EUR leaves (=hour / hours_per_fte_month * rate)
    → Plan-Forecast WP × month aggregation (=SUM of relevant leaves)
    → Actuals forecast monitoring block (4 right-side cols, green cross-refs
       into Plan-Forecast)
    → Capacity Actual-FTE / Variance band (T9 — per-row hours / hours_per_fte_month)
    → Summary Forecast EUR column per calendar year (T10)
    → Trends Cumulative Forecast helper row + chart series (T11)
    → Trends Person Load actual-side cols F (Actual Avg FTE), G (Actual Peak),
       H (Planned − Actual Variance) (T12).

  See the consolidated diagram under "═══ DATA FLOW ═══" below for the full
  chain including reconciliation via `eur_actuals_through_month`.

═══ SUMMARY SHEET STRUCTURE ═══

SHEET: SUMMARY (calendar-year budget aggregation)
  Tab color: Green (008000)

  Row 1 = "ANNUAL BUDGET SUMMARY" title (merged A:B)
  Row 3 = Calendar year group headers ("2025", "2026", ... ), each spanning 5 sub-columns (T10) + "PROJECT TOTAL" group
  Row 4 = Sub-headers per year: Plan FTE, Plan €, Act €, Forecast €, Var €
  Rows {sum_first_wp}-{sum_last_wp} = One row per WP (all formulas referencing Plan, Actuals, and Plan-Forecast sheets)
    Each calendar year group sums the months belonging to that year (partial years for first/last)
    Plan FTE = SUM of Plan sheet month range for that year
    Plan EUR = Plan FTE * per-WP rate (Config col D)
    Actual EUR = SUM of Actuals sheet month range for that year
    Forecast EUR (T10) = SUM of Plan-Forecast sheet month range for that year (green cross-ref)
    Variance EUR = Plan EUR - Actual EUR
  Row {sum_total} = Totals
  Project Total group: aggregates across all years (Plan FTE, Plan €, Act €, Forecast €, Var €)
  Conditional formatting:
    - red fill on negative Variance EUR (per year and project total)
    - red fill on Forecast EUR < 0.8 × Plan EUR (per year and project total — early-warning when timelog-derived forecast undershoots plan)

  BUDGET OVERVIEW (below totals row):
    Personnel Actual = Actuals TOTAL sum (green cross-ref)
    Total Claimable (+25%) = Personnel Actual * 1.25
    Personnel Ceiling = partner_ceiling (green cross-ref)
    Remaining = Ceiling - Personnel Actual
    % Consumed = Personnel Actual / Ceiling
    Conditional formatting: amber if Remaining < 10% of Ceiling, red if negative

  REPORTING PERIOD SUMMARY section (below Budget Overview):
    Same 4-column-per-group layout as the annual summary above,
    but grouped by user-defined reporting periods from Config.
    Columns per RP: Plan FTE, Plan EUR, Act EUR, Var EUR
    PROJECT TOTAL group at the end.

    FORM C PREPARATION rows below RP TOTAL:
      A. Personnel EUR — actual personnel costs per reporting period
      E. Indirect (25%) — 25% flat rate on personnel (minimum; excludes travel/equipment)
      Total Eligible — A + E
      Cum. Eligible — running cumulative across reporting periods
      CFS Required? — "YES" if cumulative eligible > EUR 430,000 (programme audit threshold)

    Note: Indirect costs shown are a lower bound (personnel-only base).
    Actual indirect should include travel, equipment, and other direct costs.

  Freeze panes: C{sum_first_wp}

═══ TRENDS SHEET STRUCTURE ═══

SHEET: TRENDS (PM dashboard — burn rate, exhaustion, person load)
  Tab color: Gold (FFD700)

  Row 1: Current Month label (A1) + input value (B1, blue, named range current_month)
  Row 2: (blank)

  BLOCK 1: WP BURN RATE & EXHAUSTION
    Section header row = Section header "WP BURN RATE & EXHAUSTION" (merged A-G)
    Header row = Column headers: WP ID, WP Name, % Time Elapsed, % Budget EUR, Burn Delta, Status, Exhaustion Mo.
    WP rows ({num_wps} rows) = formulas referencing Config WP table + Actuals cumulative columns
      - All cells show "—" for WPs not yet started (current_month < wp_start from Config)
      - % Time Elapsed = MIN(1, (current_month - MINIFS(sched_starts,...) + 1) / (MAXIFS(sched_ends,...) - MINIFS(...) + 1))
      - % Budget EUR = Actuals cumulative / Actuals ceiling
      - Burn Delta = % Budget EUR - % Time Elapsed (positive = overspending)
      - Status = text RAG: OK (|delta|<=0.05), WATCH (delta>0.05), OVER (delta>0.15), SLOW (delta<-0.10)
      - Exhaustion Mo. = projected month when ceiling is hit at current rate
    TOTAL row = weighted average burn delta using individual IF(ISNUMBER()) terms per WP
      Formula: =(IF(ISNUMBER(E5),E5*Actuals!ceiling5,0)+...+IF(ISNUMBER(EN),EN*Actuals!ceilingN,0))/partner_ceiling
    Conditional formatting: green/amber/red on Burn Delta, Status, Exhaustion Mo.

  BLOCK 2: PERSON LOAD (below chart) — extended in T12 with actual-side cols
    Section header "PERSON LOAD" (merged A-H)
    Column headers (8 cols):
      A=Person, B=Avg FTE (to date), C=Avg FTE (remaining),
      D=Peak Month FTE, E=Status,
      F=Actual Avg FTE (to date),  G=Actual Peak FTE,  (T12, green cross-ref)
      H=Planned − Actual Variance
    {num_team} person rows:
      - Person name: green cross-ref to Config roster (=Config!A<roster_row>)
      - Avg FTE (to date): SUMPRODUCT over Capacity planned-band TOTAL row up to current_month
      - Avg FTE (remaining): SUMPRODUCT over Capacity planned-band TOTAL row after current_month
      - Peak Month FTE: MAX of Capacity planned-band TOTAL row
      - Status: OVER/WATCH/OK based on remaining avg vs Config Max FTE
      - Actual Avg FTE (to date): IFERROR(AVERAGE(<Capacity Actual-FTE band slice 1..current_month>),0)
        — green cross-ref into Capacity right-side Actual-FTE band (T9).
      - Actual Peak FTE: IFERROR(MAX(<Capacity Actual-FTE band row>),0)
      - Planned − Actual Variance: planned avg − actual avg (black formula).
    Conditional formatting: green/amber/red on Status and Peak Month (per-person rules)

  BLOCK 3: Cumulative EUR LineChart (anchored 2 rows below TOTAL, ~15 rows tall)
    LineChart series:
      Planned (blue solid)
      Actual (red solid)
      Cumulative Forecast (amber dashed) — T11, sums Plan-Forecast monthly TOTAL
      Ceiling (grey dashed)
      RP End (green dotted, diamond markers)
    Data from hidden helper rows below person load section.

  HIDDEN HELPER ROWS (below person load):
    Helper row 1: Cumulative planned EUR per month (running sum of SUMPRODUCT(Plan FTE, per-WP rates))
    Helper row 2: Cumulative actual EUR per month (running sum of Actuals TOTAL row)
    Helper row 3 (T11): Cumulative forecast EUR per month — running sum of
      Plan-Forecast MONTHLY TOTAL row (`helper_row_forecast` in trends layout).
      Empty Plan-Forecast cells contribute 0; forecast curve only diverges from
      plan once hours start being logged.
    Ceiling line (partner_ceiling repeated per month)
    Month-year labels (e.g. "Nov 25") for chart X-axis
    RP boundary markers (partner_ceiling value at each RP end month, sparse — only RP end months populated)

═══ CONVENTIONS ═══
  FONT COLORS (financial model standard):
    Blue (0,0,255)   = Hardcoded input — user edits these
    Black (0,0,0)    = Formula — do not overwrite
    Green (0,128,0)  = Cross-sheet reference formula
  FILLS:
    Grey (D9D9D9)    = Inactive / protected area
    Yellow (FFFF00)   = Attention needed
    Light blue (D6E4F0) = Header area
    Red (FFC7CE)     = Error / overrun
    Amber (FFEB9C)   = Approaching limit (Capacity TOTAL > 0.8 x Max FTE)
  NUMBER FORMATS:
    FTE values: 0.00
    EUR values: €#,##0
    Percentages: 0.0%

═══ DATA FLOW ═══
  Config → Capacity → Plan ← Actuals → Summary (no circular refs)
                                  ↘ Trends ↙

  Hours-timelog forecast chain (v1 — fully wired in Stage 2 T8-T12):

    Config (hours_per_fte_month, roster_lookup col 4 = per-person €/PM rate,
            eu_annual_ceiling_pm, eur_actuals_through_month)
      ↓
    Actuals-Hours per-person × WP × month HOUR cells (blue input)
      ↓  =IF(h="","",ROUND(h/hours_per_fte_month*VLOOKUP(person,roster_lookup,4,FALSE),0))
    Actuals-Hours forecast-EUR leaves (collapsible band)
      ↓  =SUM(<all leaves for this WP × month>)
    Plan-Forecast WP × month grid + MONTHLY TOTAL row
      ↓
      ├──→ Actuals forecast monitoring block (Cum / Variance / % Ceiling / Breach In)
      │       — green cross-refs to Plan-Forecast WP rows
      ├──→ Capacity Actual-FTE band (per row: hour-cell / hours_per_fte_month)
      │       — same cells feed Trends Person Load actual-side cols F/G/H
      ├──→ Summary "Forecast €" column per calendar year + project total (T10)
      └──→ Trends helper_row_forecast (running sum) → "Cumulative Forecast"
              chart series (amber dashed)

    Reconciliation (T8): named cell `eur_actuals_through_month` separates
      authoritative finance-closed months (Actuals EUR) from forecast-only
      months. CF on both sheets greys the non-authoritative side per month.

  Config WP Budget table (IDs, names, FTE budgets, per-WP rates, ceilings) --> all sheets via cross-refs
  Config WP Schedule table (IDs, start/end months, type) --> active month determination + Capacity distribution curves
  Config grade rates (tbl_GradeRates) + roster grades/FTEs --> fte_rate weighted average formula
  Config team roster (names, Max FTE, Grade) --> Capacity person headers + REMAINING formulas
  Config Budget col E (Ceiling = FTE × per-WP Rate) --> Actuals ceiling validation
  Config PERSON-WP ASSIGNMENTS (Person, WP, Total FTE, Start, End) --> per-person filtered WP rows on Capacity
  Config EXTERNAL COMMITMENTS (Person, Project, Total FTE, Start, End, EU?) --> external project rows on Capacity
  Capacity monthly FTE (static values from distribution) --> Plan Planned FTE (SUM formulas across people)
  Actuals monthly EUR (input) --> Summary actual EUR aggregation
  Plan Planned FTE --> Summary planned FTE/EUR aggregation
  Actuals cumulative columns + Config WP table --> Trends burn rate & exhaustion
  Capacity TOTAL rows + Config roster --> Trends person load
  Plan Planned FTE + Actuals TOTAL --> Trends cumulative chart

  PRIMARY PLANNING SURFACE: Config sheet PERSON-WP ASSIGNMENTS (B2) and EXTERNAL COMMITMENTS (B3)
  sections. Edit person-WP assignments and external commitments there. Monthly FTE values on the
  Capacity sheet are auto-filled by distribution curves on rebuild. Users can override individual
  month cells on Capacity post-generation for fine-tuning.
  Plan Planned FTE is formula-driven — do not edit directly.
  Actuals sheet is the input surface for EUR spend — enter accountant figures there.
  Trends sheet is read-only formulas — set current_month (Trends!B1) to update.

═══ CLI COMMANDS ═══
  python -m budget_planner init                    — Create fresh empty workbook from config.py seed data
  python -m budget_planner init --demo             — Create workbook with demo Capacity + Actuals data
  python -m budget_planner rebuild [path]          — Rebuild from existing workbook (preserves Capacity FTE + Actuals EUR)
  python -m budget_planner --snapshot <path>       — Save snapshot for scenario comparison
  python -m budget_planner --restore <path>        — Restore from snapshot

═══ BUILD ORDER ═══
  Build order (dependency DAG, enforced by test_build_order_is_stable):
    Config → Capacity → Plan → patch Plan → Trends stub
      → Actuals-Hours → Plan-Forecast → Actuals → Summary → Trends → _META

  Tab order (user-facing, set after build via wb.move_sheet):
    _META, Config, Capacity, Plan, Actuals, Actuals-Hours, Plan-Forecast, Trends, Summary

  Why this order:
    Config sheet created first so named ranges are available to every
    downstream builder: project_start, project_months, fte_rate,
    eu_annual_ceiling_pm, hours_per_fte_month, eur_actuals_through_month,
    tbl_GradeRates, roster_ftes, roster_grades, roster_lookup,
    sched_wp_ids/starts/ends.
    Capacity sheet built next; reads config_first_wp_row from config_layout.
    Plan sheet built, then its FTE cells overwritten with SUM formulas referencing Capacity.
    Trends stub (B1 + current_month named range) created before Actuals so
    current_month is available to Actuals' cumulative formulas.
    Actuals-Hours built before Plan-Forecast so the forecast-EUR leaves exist
    for Plan-Forecast to SUM.
    Plan-Forecast built before Actuals so the forecast monitoring block on
    Actuals can green-cross-ref Plan-Forecast WP rows.
    Summary built after Plan-Forecast so its Forecast EUR column (T10) can
    pull from Plan-Forecast.
    Trends content populated after Actuals/Summary/Plan-Forecast so their
    layouts (and helper_row_forecast for the Cumulative Forecast series)
    are known.
    _META is built last so it can document live row/column positions from
    every sibling builder's returned layout dict.

═══ NAMED RANGES ═══
  project_start        = Config!$B$2   (project start date)
  project_months       = Config!$B$3   (integer project duration)
  fte_rate              = Config!$B$4   (weighted-average EUR/PM rate — formula, not static input)
  eu_annual_ceiling_pm = Config!$B$5   (max PM per calendar year across all EU grants; ≈215 days)
  hours_per_fte_month  = Config!$B$<row>  (hours per FTE-month; default 143 ≈ 215 days/year / 12 × 8h; divisor for Actuals-Hours forecast-EUR leaves and Capacity Actual-FTE band)
  eur_actuals_through_month = Config!$B$<row>  (last month index with finance-closed Actual EUR; blue input integer. Drives the T8 reconciliation CF on Actuals + Actuals-Hours: months past this index grey on Actuals; months ≤ this index grey on Actuals-Hours forecast band.)
  tbl_GradeRates       = Config!$A${cfg_grade_first}:$B${cfg_grade_last}  (grade name + rate lookup table)
  roster_ftes          = Config!$B${cfg_roster_first}:$B${cfg_roster_last}  (team Max FTE column)
  roster_grades        = Config!$C${cfg_roster_first}:$C${cfg_roster_last}  (team Grade column)
  roster_lookup        = Config!$A${cfg_roster_first}:$D${cfg_roster_last}  (4-col lookup: Name | Max FTE | Grade | Rate; col 4 used by Actuals-Hours forecast-EUR leaves)
  sched_wp_ids         = Config!$A$<sched_first>:$A$<sched_last>  (WP IDs in schedule table)
  sched_starts         = Config!$B$<sched_first>:$B$<sched_last>  (start months in schedule table)
  sched_ends           = Config!$C$<sched_first>:$C$<sched_last>  (end months in schedule table)
  partner_ceiling      = Plan!$G$1     (total personnel ceiling in EUR, =SUM of Config!E WP column)
  current_month        = Trends!$B$1   (last month with completed actuals, blue input)

FORMULAS USE ABSOLUTE REFERENCES for config lookups and mixed references for month iterations.

═══ MODIFICATION GUIDE (for AI agents and PMs) ═══

TO ADD A WORK PACKAGE:
  1. Add a row to the WP BUDGET table on Config (WP ID, Name, FTE Budget, Rate)
  2. Add one or more schedule rows to the WP SCHEDULE table on Config (WP ID, Start, End, Type)
  3. Add person-WP assignment rows in the Config PERSON-WP ASSIGNMENTS section (B2)
  4. Run: python -m budget_planner rebuild
  5. Rebuild regenerates all sheets. Existing Config data and Actuals EUR are preserved.

TO ADD A PERSON-WP ASSIGNMENT:
  1. Add a row to the PERSON-WP ASSIGNMENTS section on the Config sheet (B2)
     (Person, WP, Total FTE, Start Month, End Month, Type)
  2. Run: python -m budget_planner rebuild
  3. The person's grid on Capacity will show the new WP with auto-filled monthly FTE values.

TO ADD A TEAM MEMBER:
  1. Add a row to the Config sheet TEAM ROSTER section (Name, Max FTE, Grade)
     Grade must be one of the values in the RATE BY GRADE table.
  2. Add person-WP assignment rows in the Config PERSON-WP ASSIGNMENTS section (B2)
  3. Run: python -m budget_planner rebuild

TO CHANGE PROJECT DURATION:
  1. Edit Project Months on the Config sheet (cell B3)
  2. Run: python -m budget_planner rebuild
  3. Adds/removes month columns. Existing data for months within the new range is preserved.

TO CHANGE FTE RATE:
  The FTE rate is a weighted-average formula — not directly editable.
  To change it, edit the per-grade rates in the RATE BY GRADE section or
  change team member grades/FTEs in the TEAM ROSTER. The fte_rate formula
  recalculates automatically.

TO REPLACE EXAMPLE DATA:
  Edit Config sheet: WP Budget + WP Schedule tables, PERSON-WP ASSIGNMENTS (B2),
  EXTERNAL COMMITMENTS (B3), TEAM ROSTER, RATE BY GRADE. Then run: python -m budget_planner rebuild

SCENARIO PLANNING:
  Use --snapshot to save current state, edit Capacity/Actuals freely, --restore to revert.

TO CREATE A DEMO WORKBOOK:
  Run: python -m budget_planner init --demo
  Seeds Capacity with time-based FTE assignments and Actuals with multiplied planned EUR.

═══ CELL PROTECTION ═══
  Protection is currently DISABLED (commented out in workbook.py).
  When re-enabled: all sheets password-protected (password: "grant").
  Only blue input cells are unlocked for editing.
  Formula cells (black/green font) are locked — do not overwrite.

═══ KNOWN LIMITATIONS ═══

1. MANUAL ROW INSERTION BREAKS FORMULAS.
   If you manually insert or delete rows in the Capacity sheet (e.g., add an
   external project), Plan Planned FTE formulas will NOT auto-adjust (openpyxl
   does not update formulas like Excel does). Use rebuild to regenerate.

2. RENAMED WP IDs LOSE PRESERVED DATA.
   If you rename a WP ID on the Config sheet, rebuild cannot match old Capacity
   and Actuals data to the new ID. Orphaned data is logged in the LAST REBUILD
   LOG section below.

3. MONTHLY OVERRIDES ARE LOST ON REBUILD.
   Monthly FTE values in Capacity are recalculated from the Config PERSON-WP ASSIGNMENTS
   section on every rebuild. If you manually override individual month cells for
   fine-tuning, those overrides will be replaced by distribution-curve values on the
   next rebuild.
""".format(
        date=datetime.now().strftime("%Y-%m-%d"),
        pm=project_params["months"],
        plan_header=plan_header,
        plan_year_row=plan_year_row,
        plan_month_row=plan_month_row,
        plan_first_wp=plan_first_wp,
        plan_last_wp=plan_last_wp,
        plan_total=plan_total,
        cfg_first_wp=cfg_first_wp,
        cfg_last_wp=cfg_last_wp,
        cfg_total_row=cfg_total_row,
        cfg_grade_section=cfg_grade_section,
        cfg_grade_header=cfg_grade_header,
        cfg_grade_first=cfg_grade_first,
        cfg_grade_last=cfg_grade_last,
        num_grades=num_grades,
        grade_names_str=grade_names_str,
        cfg_wp_section=cfg_wp_section,
        cfg_wp_header=cfg_wp_header,
        cfg_roster_section=cfg_roster_section,
        cfg_roster_header=cfg_roster_header,
        cfg_roster_first=cfg_roster_first,
        cfg_roster_last=cfg_roster_last,
        act_last_wp=4 + num_wps - 1,
        act_total=4 + num_wps,
        actuals_amber_pct=int(0.85 * 100),
        actuals_amber_breach=12,
        actuals_red_breach=6,
        capacity_amber_load=CAPACITY_AMBER_LOAD,
        variance_red_fte=CAPACITY_VARIANCE_FTE_RED,
        first_mcol=first_mcol,
        last_mcol=last_mcol,
        val_sum_col=val_sum_col,
        val_budget_col=val_budget_col,
        val_diff_col=val_diff_col,
        val_cum_col=val_cum_col,
        val_pct_col=val_pct_col,
        val_breach_col=val_breach_col,
        sum_first_wp=sum_first_wp,
        sum_last_wp=sum_last_wp,
        sum_total=sum_total,
        num_wps=num_wps,
        num_team=num_team,
        num_ext=num_ext,
        ah_band_label=ah_band_label,
        ah_year_header=ah_year_header,
        ah_month_header=ah_month_header,
        ah_first_person_header=ah_first_person_header,
        pf_forecast_header=pf_forecast_header,
        pf_first_wp_row=pf_first_wp_row,
        pf_last_wp_row=pf_last_wp_row,
        pf_total_row=pf_total_row,
    )

    if rebuild_warnings:
        meta_text += "\n═══ LAST REBUILD LOG ═══\n"
        meta_text += f"Rebuild date: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n"
        for warning in rebuild_warnings:
            meta_text += f"  WARNING: {warning}\n"
        meta_text += "\n"

    # Excel cells cap at 32 767 chars and silently truncate longer values
    # (openpyxl does not warn). The full v1 _META text exceeds that, so split
    # at section-header boundaries (lines starting with "═══ ") into one
    # cell per chunk, preserving readability and search-ability.
    _CELL_CHAR_LIMIT = 32_767
    chunks = _split_meta_for_cells(meta_text, _CELL_CHAR_LIMIT)
    for i, chunk in enumerate(chunks, start=1):
        cell = ws.cell(row=i, column=1, value=chunk)
        cell.font = Font(name="Consolas", size=10)
        cell.alignment = WRAP
    ws.column_dimensions["A"].width = 120
    ws.sheet_properties.tabColor = "FF0000"


def _split_meta_for_cells(text: str, limit: int) -> list[str]:
    """Split ``text`` into chunks ≤ ``limit`` chars, preferring to break at
    blank lines preceding "═══ " section headers.

    Excel's per-cell character limit is 32 767. The meta builder writes a
    single long string; chunks are written to consecutive cells in column A
    so the rendered _META sheet reads top-to-bottom as one continuous doc.

    Falls back to a hard slice at ``limit`` if no section-header break fits
    within a chunk (defensive — every real meta_text has plenty of headers).
    """
    if len(text) <= limit:
        return [text]
    chunks: list[str] = []
    remaining = text
    while len(remaining) > limit:
        # Find the last "═══ " header within the limit window so the cut
        # lands on a clean section boundary.
        window = remaining[:limit]
        cut = window.rfind("\n═══ ")
        if cut < 0:
            cut = limit  # hard fallback (no header in range)
        chunks.append(remaining[:cut])
        remaining = remaining[cut:].lstrip("\n")
    if remaining:
        chunks.append(remaining)
    return chunks
