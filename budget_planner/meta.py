"""Build the _META sheet — AI-agent-readable workbook documentation."""
from openpyxl.styles import Font
from datetime import datetime

from budget_planner.styles import WRAP, col_letter


def build_meta_sheet(ws, layout, *, project_params=None, work_packages=None,
                     team=None, external_commitments=None, rebuild_warnings=None):
    """Populate the _META worksheet with structural documentation.

    Parameters
    ----------
    ws : openpyxl.worksheet.worksheet.Worksheet
        The _META sheet (already created by the caller).
    layout : dict
        Row-number mapping produced by the plan-sheet builder.  Keys used:
        config_total_row, config_last_wp_row, planned_header_row,
        planned_first_wp_row, planned_last_wp_row, planned_total_row.
    project_params : dict or None
        Dict with keys title, start, months, rate. Falls back to config.
    work_packages : list or None
        List of WP dicts. Falls back to config.
    team : list or None
        List of team member dicts. Falls back to config.
    external_commitments : list or None
        List of external commitment dicts. Falls back to config.
    rebuild_warnings : list or None
        List of warning strings to append to meta text.
    """
    if any(v is None for v in [project_params, work_packages, team, external_commitments]):
        from budget_planner import config
        if project_params is None:
            project_params = {
                "title": config.PROJECT_TITLE, "start": config.PROJECT_START,
                "months": config.PROJECT_MONTHS, "rate": config.DEFAULT_RATE,
            }
        if work_packages is None:
            work_packages = config.WORK_PACKAGES
        if team is None:
            team = config.TEAM
        if external_commitments is None:
            external_commitments = config.EXTERNAL_COMMITMENTS

    assert project_params is not None
    assert work_packages is not None
    assert team is not None
    assert external_commitments is not None

    month_col_start = 3
    num_wps = len(work_packages)
    month_col_end = month_col_start + project_params["months"] - 1
    val_col_sum = month_col_end + 1
    val_col_budget = month_col_end + 2
    val_col_diff = month_col_end + 3
    val_col_cum = val_col_diff + 1
    val_col_pct_ceil = val_col_diff + 2
    val_col_breach = val_col_diff + 3

    # Derive row numbers from layout dict
    config_end = layout["config_total_row"]
    wp_last_config = layout["config_last_wp_row"]
    config_total = layout["config_total_row"]
    plan_header = layout["planned_header_row"]
    plan_year_row = layout["planned_header_row"] + 1
    plan_month_row = layout["planned_header_row"] + 2
    plan_first_wp = layout["planned_first_wp_row"]
    plan_last_wp = layout["planned_last_wp_row"]
    plan_total = layout["planned_total_row"]

    # Summary sheet references
    sum_last_wp = 4 + num_wps
    sum_total = 5 + num_wps

    # Team sheet references
    num_team = len(team)
    roster_last = 2 + num_team  # header rows 1-2, data starts row 3
    num_ext = len(external_commitments)

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
VERSION: 3.0
CREATED: {date}
PURPOSE: Track planned FTE allocation and actual EUR spend per work package per month for an EU-funded (programme) project. Per-person capacity planning with cross-project visibility. Derive planned euro amounts from FTE using a configurable rate. Record actual EUR from accountant reports on a separate sheet.

═══ SHEET INDEX ═══
1. _META        — This sheet. AI-agent-readable documentation of workbook structure, conventions, and modification instructions.
2. Plan         — Global parameters + config table + Planned FTE (formulas from Capacity).
3. Team         — Person roster + WP share seeds for initial allocation.
4. Capacity     — Primary planning surface: per-person monthly FTE allocation across programme WPs and external projects.
5. Actuals      — EUR actuals per WP per month (input from accountant). Tab color: Blue.
6. Summary      — Annual aggregation. Plan FTE, Plan EUR, Actual EUR, Variance EUR per WP per year. Budget overview.
7. Trends       — WP burn rate & exhaustion, cumulative EUR chart, person load. Tab color: Gold (FFD700).

═══ TEAM SHEET STRUCTURE ═══

SHEET: TEAM
  Tab color: Orange

  BLOCK 1: TEAM ROSTER (rows 1-{roster_last})
    Row 1: Section header "TEAM ROSTER"
    Row 2: Column headers (Person, Max FTE)
    Rows 3-{roster_last}: One row per team member (input, blue text)

  BLOCK 2: WP SHARE SEEDS (below roster)
    Section header + column headers (Person, then one col per WP ID)
    Data rows: one per team member, share fraction per WP (input, blue text)
    Validation row: SUM per column (should equal 1.0)
    NOTE: Shares are used only for initial seeding. Capacity is source of truth.

═══ CAPACITY SHEET STRUCTURE ═══

SHEET: CAPACITY (primary planning surface)
  Tab color: Purple

  CONFIG BLOCK: EXTERNAL COMMITMENTS (top of sheet)
    Columns: Person, Project, Total FTE, Start Month, End Month
    {num_ext} entries from seed data (all input, blue text)

  YEAR/MONTH HEADERS: same layout as Plan sheet

  PER-PERSON GROUPS: one group per team member ({num_team} people)
    Person Header: bold, section fill, formula ref to Team roster
    Programme WP Rows: WP ID/Name formula refs to Plan config, monthly FTE (input, blue)
    External Project Rows: project name (input), monthly FTE (input, blue)
    TOTAL Row: SUM of all rows above (formula)
    REMAINING Row: Team Max FTE - TOTAL (formula)

  CONDITIONAL FORMATTING:
    TOTAL > Max FTE: red fill (overcommitted)
    TOTAL > 0.8 x Max FTE: amber fill (approaching limit)
    REMAINING < 0: red fill

  VALIDATION COLUMNS: {val_sum_col} = Sum, {val_budget_col} = Budget, {val_diff_col} = Variance (per row)

═══ PLAN SHEET STRUCTURE ═══

BLOCK 1: GLOBAL PARAMETERS (Rows 1-4, columns A-B only)
  B1 = Project title (text, input)
  B2 = Project start date (date, input). All month columns derive calendar dates from this.
  B3 = Monthly FTE rate in EUR (number, input). Used to derive euro amounts.
  B4 = Personnel Ceiling (formula: SUM of config column G). Total budget in EUR.

BLOCK 2: CONFIGURATION TABLE (Rows 6-{config_end})
  Row 6 = Section header "CONFIGURATION"
  Row 7 = Column headers
  Rows 8-{wp_last_config} = One row per WP
  Columns:
    A = WP ID (text, input)
    B = WP Name (text, input)
    C = Total FTE budget in person-months (number, input)
    D = Start month number 1-{pm} (integer, input)
    E = End month number 1-{pm} (integer, input)
    F = Active months (formula: E-D+1)
    G = Personnel Ceiling (€) (formula: C * fte_rate)
  Row {config_total} = Totals row

BLOCK 3: PLANNED FTE (Rows {plan_header}-{plan_total})
  Row {plan_header} = Section header "PLANNED FTE (PERSON-MONTHS)"
  Row {plan_year_row} = Year labels (merged per 12 months)
  Row {plan_month_row} = Month numbers (M1..M{pm}) in columns {first_mcol}..{last_mcol}
  Rows {plan_first_wp}-{plan_last_wp} = Monthly FTE per WP
    Col A = WP ID (formula, references config)
    Col B = WP Name (formula, references config)
    Cols {first_mcol}-{last_mcol} = SUM formulas referencing Capacity sheet (black text, formula)
      - Each cell sums that WP's allocation across all persons in Capacity
      - NOT directly editable — edit Capacity sheet instead
    Col {val_sum_col} = Sum of monthly values (formula: SUM)
    Col {val_budget_col} = Budget FTE from config (formula, green cross-ref)
    Col {val_diff_col} = Variance: Budget - Sum (formula). Zero = OK, nonzero = WARNING.
  Row {plan_total} = Column totals (SUM formula per month)

═══ ACTUALS SHEET STRUCTURE ═══

SHEET: ACTUALS (EUR actuals from accountant)
  Tab color: Blue (0000FF)

  Row 1 = Section header "ACTUAL EUR"
  Row 2 = Year labels (merged per 12 months, same layout as Plan)
  Row 3 = Month headers (M1..M{pm}) in columns {first_mcol}..{last_mcol}
  Rows 4-{act_last_wp} = One row per WP
    Col A = WP ID (formula, cross-ref to Plan config via =Plan!A{{config_row}})
    Col B = WP Name (formula, cross-ref to Plan config via =Plan!B{{config_row}})
    Cols {first_mcol}-{last_mcol} = EUR amount per month (input, blue text, €#,##0 format)
      - Active months initialized to 0
      - Inactive months greyed out and empty (same WP active period as Plan)
    Col {val_sum_col} = Sum of monthly EUR (formula: SUM)
    Col {val_budget_col} = Ceiling EUR from Plan config column G (formula, green cross-ref)
    Col {val_diff_col} = Variance: Ceiling - Sum (formula). Negative = OVERRUN.
    Col {val_cum_col} = Cum. EUR: SUMPRODUCT cumulative up to current_month (formula)
    Col {val_pct_col} = % Ceiling: Cum. EUR / Ceiling (formula, 0.0% format)
    Col {val_breach_col} = Breach In: months until ceiling exhausted at current rate (formula)
  TOTAL row:
    Monthly columns = SUM per column across WP rows
    {val_sum_col} = SUM of WP sums
    {val_budget_col} = partner_ceiling named range (total project ceiling)
    {val_diff_col} = Ceiling - Sum
    {val_cum_col}, {val_pct_col}, {val_breach_col} = same formulas on TOTAL row
  Conditional formatting:
    - Red fill on negative variance
    - % Ceiling: amber >85%, red >100%
    - Breach In: amber <12 months, red <6 months

═══ SUMMARY SHEET STRUCTURE ═══
  Row 1 = Title
  Row 3 = Year group headers (Year 1..Year N), each spanning 4 sub-columns + Project Total group
  Row 4 = Sub-headers per year: Plan FTE, Plan €, Act €, Var €
  Rows 5-{sum_last_wp} = One row per WP (all formulas referencing Plan and Actuals sheets)
  Row {sum_total} = Totals
  Right-most group (Project Total): aggregates across all years (Plan FTE, Plan €, Act €, Var €)
  Budget overview: per-WP Planned EUR, Actual EUR, Remaining EUR derived from Plan config and Actuals

═══ TRENDS SHEET STRUCTURE ═══

SHEET: TRENDS (PM dashboard — burn rate, exhaustion, person load)
  Tab color: Gold (FFD700)

  Row 1: Current Month label (A1) + input value (B1, blue, named range current_month)
  Row 2: (blank)

  BLOCK 1: WP BURN RATE & EXHAUSTION (rows 3-13)
    Row 3 = Section header "WP BURN RATE & EXHAUSTION" (merged A-G)
    Row 4 = Column headers: WP ID, WP Name, % Time Elapsed, % Budget EUR, Burn Delta, Status, Exhaustion Mo.
    Rows 5-12 = One row per WP (formulas referencing Plan config + Actuals cumulative columns)
      - All cells show "—" for WPs not yet started (current_month < wp_start)
      - Burn Delta = % Budget EUR - % Time Elapsed (positive = overspending)
      - Status = text RAG: OK, WATCH, OVER, SLOW
      - Exhaustion Mo. = projected month when ceiling is hit at current rate
    Row 13 = TOTAL (weighted average burn delta by WP ceiling share)
    Conditional formatting: green/amber/red on Burn Delta, Status, Exhaustion Mo.

  BLOCK 2: Cumulative EUR Chart (anchored at A15, ~15 rows tall)
    LineChart: Planned (blue) vs Actual (red) vs Ceiling (grey dashed)
    Data from hidden helper rows 40-43

  BLOCK 3: PERSON LOAD (rows 32-37)
    Row 32 = Section header "PERSON LOAD" (merged A-E)
    Row 33 = Column headers: Person, Avg FTE (to date), Avg FTE (remaining), Peak Month FTE, Status
    Rows 34-37 = One row per team member ({num_team} people)
      - Formulas reference Capacity TOTAL rows and Team roster
      - Status = OVER/WATCH/OK based on remaining avg vs max_fte
    Conditional formatting: green/amber/red on Status and Peak Month

  HIDDEN HELPER ROWS (40-43):
    Row 40 = Cumulative planned EUR per month (running sum of Plan FTE * fte_rate)
    Row 41 = Cumulative actual EUR per month (running sum of Actuals TOTAL row)
    Row 42 = Ceiling line (partner_ceiling repeated)
    Row 43 = Month labels M1-M{pm} for chart X-axis

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
  Team → Capacity → Plan ← Actuals → Summary (no circular refs)
                                  ↘ Trends ↙

  Team roster (names, Max FTE) --> Capacity person headers + REMAINING formulas
  Team WP shares --> Capacity seed values (generation-time only)
  Capacity monthly FTE (input) --> Plan Planned FTE (SUM formulas across people)
  Plan config (WP budgets) --> Actuals ceiling validation
  Actuals monthly EUR (input) --> Summary actual EUR aggregation
  Plan Planned FTE --> Summary planned FTE/EUR aggregation
  Actuals cumulative columns + Plan config --> Trends burn rate & exhaustion
  Capacity TOTAL rows + Team roster --> Trends person load
  Plan Planned FTE + Actuals TOTAL --> Trends cumulative chart

  PRIMARY PLANNING SURFACE: Capacity sheet. Edit monthly FTE per person there.
  Plan Planned FTE is formula-driven — do not edit directly.
  Actuals sheet is the input surface for EUR spend — enter accountant figures there.
  Trends sheet is read-only formulas — set current_month (Trends!B1) to update.

═══ BUILD ORDER ═══
  Config → Plan → Team → Capacity → patch Plan → Trends skeleton → Actuals → Summary → Trends → _META
  Config sheet created first so named ranges are available for Plan formulas.
  Trends skeleton (B1 + named range) created before Actuals so current_month is available.
  Trends content populated after Actuals/Summary so their layouts are known.

═══ NAMED RANGES ═══
  project_start   = Config!$B$2   (project start date)
  fte_rate         = Config!$B$4   (monthly EUR rate per person-month)
  partner_ceiling  = Plan!$B$4     (total personnel ceiling in EUR)
  current_month    = Trends!$B$1   (last month with completed actuals, blue input)

FORMULAS USE ABSOLUTE REFERENCES for config lookups and mixed references for month iterations.

═══ MODIFICATION GUIDE (for AI agents and PMs) ═══

TO ADD A WORK PACKAGE:
  1. Add a row to tbl_WorkPackages on the Config sheet (WP ID, Name, FTE, Start, End)
  2. Run: python -m budget_planner rebuild
  3. Rebuild regenerates all sheets. Existing Capacity FTE and Actuals EUR are preserved.

TO ADD A TEAM MEMBER:
  1. Add a row to the Team sheet roster (Name, Max FTE)
  2. Add share seeds in the WP Share Seeds matrix
  3. Run: python -m budget_planner rebuild

TO CHANGE PROJECT DURATION:
  1. Edit Project Months on the Config sheet (cell B3)
  2. Run: python -m budget_planner rebuild
  3. Adds/removes month columns. Existing data for months within the new range is preserved.

TO CHANGE FTE RATE:
  Edit cell B4 on Config sheet. All EUR calculations update via the fte_rate named range.

TO REPLACE EXAMPLE DATA:
  Edit Config sheet (WPs), Team sheet (roster + shares), Capacity config block
  (external commitments). Then run: python -m budget_planner rebuild

SCENARIO PLANNING:
  Use --snapshot to save current state, edit Capacity/Actuals freely, --restore to revert.

═══ CELL PROTECTION ═══
  All sheets are password-protected (password: "grant").
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
""".format(
        date=datetime.now().strftime("%Y-%m-%d"),
        config_end=config_end,
        wp_last_config=wp_last_config,
        config_total=config_total,
        pm=project_params["months"],
        plan_header=plan_header,
        plan_year_row=plan_year_row,
        plan_month_row=plan_month_row,
        plan_first_wp=plan_first_wp,
        plan_last_wp=plan_last_wp,
        plan_total=plan_total,
        act_last_wp=4 + num_wps - 1,
        first_mcol=first_mcol,
        last_mcol=last_mcol,
        val_sum_col=val_sum_col,
        val_budget_col=val_budget_col,
        val_diff_col=val_diff_col,
        val_cum_col=val_cum_col,
        val_pct_col=val_pct_col,
        val_breach_col=val_breach_col,
        sum_last_wp=sum_last_wp,
        sum_total=sum_total,
        roster_last=roster_last,
        num_team=num_team,
        num_ext=num_ext,
    )

    if rebuild_warnings:
        meta_text += "\n═══ LAST REBUILD LOG ═══\n"
        meta_text += f"Rebuild date: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n"
        for warning in rebuild_warnings:
            meta_text += f"  WARNING: {warning}\n"
        meta_text += "\n"

    ws["A1"] = meta_text
    ws["A1"].font = Font(name="Consolas", size=10)
    ws["A1"].alignment = WRAP
    ws.column_dimensions["A"].width = 120
    ws.sheet_properties.tabColor = "FF0000"
