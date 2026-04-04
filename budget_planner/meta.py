"""Build the _META sheet — AI-agent-readable workbook documentation."""
from openpyxl.styles import Font
from datetime import datetime

from budget_planner.layout import MONTH_COL_START, month_col_end as _month_col_end, val_columns
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
        config_first_wp_row, config_last_wp_row, planned_header_row,
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

    month_col_start = MONTH_COL_START
    num_wps = len(work_packages)
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

    # Config sheet row positions
    cfg_first_wp = layout["config_first_wp_row"]
    cfg_last_wp = layout["config_last_wp_row"]
    cfg_total_row = cfg_last_wp + 1
    cfg_roster_section = cfg_total_row + 2
    cfg_roster_header = cfg_roster_section + 1
    cfg_roster_first = cfg_roster_header + 1

    # Summary sheet references (row 3=year headers, row 4=sub-headers, row 5+=WPs)
    sum_first_wp = 5
    sum_last_wp = sum_first_wp + num_wps - 1
    sum_total = sum_last_wp + 1

    # Team roster references (in Config sheet)
    num_team = len(team)
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
VERSION: 4.0
CREATED: {date}
PURPOSE: Track planned FTE allocation and actual EUR spend per work package per month for an EU-funded (programme) project. Per-person capacity planning with cross-project visibility. Derive planned euro amounts from FTE using a configurable rate. Record actual EUR from accountant reports on a separate sheet.

═══ SHEET INDEX ═══
1. _META        — This sheet. AI-agent-readable documentation of workbook structure, conventions, and modification instructions.
2. Config       — Project parameters (rows 1-4) + WP table (rows 6+) + team roster (below WP TOTAL).
3. Plan         — Compact 1-row header + Planned FTE block (formulas from Capacity).
4. Capacity     — Primary planning surface: per-person monthly FTE allocation across programme WPs and external projects.
5. Actuals      — EUR actuals per WP per month (input from accountant). Tab color: Blue.
6. Summary      — Calendar-year aggregation. Plan FTE, Plan EUR, Actual EUR, Variance EUR per WP per year. Budget overview.
7. Trends       — WP burn rate & exhaustion, cumulative EUR chart, person load. Tab color: Gold (FFD700).

═══ CONFIG SHEET STRUCTURE ═══

SHEET: CONFIG (single source of truth for all project parameters)
  Tab color: Blue (4472C4)
  Freeze panes: A6

  SECTION A: PROJECT PARAMETERS (rows 1-4)
    Row 1: A1="Project Title",    B1=<title>           (blue input)
    Row 2: A2="Project Start",    B2=<date>            (blue input, MMM YYYY)
    Row 3: A3="Project Months",   B3=<integer>         (blue input)
    Row 4: A4="FTE Rate (EUR)",   B4=<rate>            (blue input, €#,##0)
    Row 5: (blank separator)

  SECTION B: WORK PACKAGES TABLE (rows 6+)
    Row 6: Section header "WORK PACKAGES" (merged A6:G6, section fill)
    Row 7: Column headers: WP ID | WP Name | FTE Budget | Start Month | End Month | Active Months | Personnel Ceiling (€)
    Rows {cfg_first_wp}-{cfg_last_wp}: WP data rows
      Cols A-E: blue input (WP ID, Name, FTE Budget, Start Mo, End Mo)
      Col F: =E-D+1 (Active Months, black formula)
      Col G: =C*fte_rate (Personnel Ceiling EUR, black formula)
    Row {cfg_total_row}: TOTAL row (SUM on cols C and G, light blue fill)
    Excel Table "tbl_WorkPackages" covers A7:G{cfg_last_wp} (header + data, excludes TOTAL)
    Conditional formatting: yellow fill if Start > End

  SECTION C: TEAM ROSTER (below WP TOTAL)
    Row {cfg_roster_section}: Section header "TEAM ROSTER" (merged A:B, section fill)
    Row {cfg_roster_header}: Column headers: Person | Max FTE
    Rows {cfg_roster_first}+: Team member data rows (blue input, {num_team} people)

═══ CAPACITY SHEET STRUCTURE ═══

SHEET: CAPACITY (primary planning surface)
  Tab color: Purple

  CONFIG BLOCK: EXTERNAL COMMITMENTS (top of sheet)
    Columns: Person, Project, Total FTE, Start Month, End Month
    {num_ext} entries from seed data (all input, blue text)

  YEAR/MONTH HEADERS: calendar year grouping (same layout as Plan/Actuals)

  PER-PERSON GROUPS: one group per team member ({num_team} people)
    Person Header: bold, section fill, formula ref to Config roster (=Config!A<roster_row>)
    Programme WP Rows: WP ID/Name formula refs to Config WP table, monthly FTE (input, blue)
    External Project Rows: project name (input), monthly FTE (input, blue)
    TOTAL Row: SUM of all rows above (formula)
    REMAINING Row: Config Max FTE - TOTAL (formula, refs Config!B<roster_row>)

  CONDITIONAL FORMATTING:
    TOTAL > Max FTE: red fill (overcommitted)
    TOTAL > 0.8 x Max FTE: amber fill (approaching limit)
    REMAINING < 0: red fill

  VALIDATION COLUMNS: {val_sum_col} = Sum, {val_budget_col} = Budget, {val_diff_col} = Variance (per row)

═══ PLAN SHEET STRUCTURE ═══

ROW 1: COMPACT HEADER
  A1 = "Title:", B1 = Project title (green cross-ref to Config!B1, bold 14pt)
  D1 = "Start:", E1 = Project start date (green cross-ref to Config!B2, MMM YYYY)
  F1 = "Rate:",  G1 = FTE Rate (green cross-ref to Config!B4, €#,##0)
  H1 = "Ceiling:", I1 = Personnel Ceiling (formula: =SUM(Config!G{cfg_first_wp}:G{cfg_last_wp}), €#,##0)
       I1 also serves as the partner_ceiling named range target.

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
    Col {val_budget_col} = Ceiling EUR from Config col G (green cross-ref: =Config!G<config_row>)
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
  Conditional formatting:
    - Red fill on negative variance
    - % Ceiling: amber >{actuals_amber_pct}%, red >100%
    - Breach In: amber <{actuals_amber_breach} months, red <{actuals_red_breach} months
  Freeze panes: C4

═══ SUMMARY SHEET STRUCTURE ═══

SHEET: SUMMARY (calendar-year budget aggregation)
  Tab color: Green (008000)

  Row 1 = "ANNUAL BUDGET SUMMARY" title (merged A:B)
  Row 3 = Calendar year group headers ("2025", "2026", ... ), each spanning 4 sub-columns + "PROJECT TOTAL" group
  Row 4 = Sub-headers per year: Plan FTE, Plan €, Act €, Var €
  Rows {sum_first_wp}-{sum_last_wp} = One row per WP (all formulas referencing Plan and Actuals sheets)
    Each calendar year group sums the months belonging to that year (partial years for first/last)
    Plan FTE = SUM of Plan sheet month range for that year
    Plan EUR = Plan FTE * fte_rate
    Actual EUR = SUM of Actuals sheet month range for that year
    Variance EUR = Plan EUR - Actual EUR
  Row {sum_total} = Totals
  Project Total group: aggregates across all years (Plan FTE, Plan €, Act €, Var €)
  Conditional formatting: red fill on negative Variance EUR (per year and project total)

  BUDGET OVERVIEW (below totals row):
    Personnel Actual = Actuals TOTAL sum (green cross-ref)
    Total Claimable (+25%) = Personnel Actual * 1.25
    Personnel Ceiling = partner_ceiling (green cross-ref)
    Remaining = Ceiling - Personnel Actual
    % Consumed = Personnel Actual / Ceiling
    Conditional formatting: amber if Remaining < 10% of Ceiling, red if negative

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
      - % Time Elapsed = MIN(1, (current_month - Config!D<row> + 1) / (Config!E<row> - Config!D<row> + 1))
      - % Budget EUR = Actuals cumulative / Actuals ceiling
      - Burn Delta = % Budget EUR - % Time Elapsed (positive = overspending)
      - Status = text RAG: OK (|delta|<=0.05), WATCH (delta>0.05), OVER (delta>0.15), SLOW (delta<-0.10)
      - Exhaustion Mo. = projected month when ceiling is hit at current rate
    TOTAL row = weighted average burn delta using individual IF(ISNUMBER()) terms per WP
      Formula: =(IF(ISNUMBER(E5),E5*Actuals!ceiling5,0)+...+IF(ISNUMBER(EN),EN*Actuals!ceilingN,0))/partner_ceiling
    Conditional formatting: green/amber/red on Burn Delta, Status, Exhaustion Mo.

  BLOCK 2: PERSON LOAD (below chart)
    Section header "PERSON LOAD" (merged A-E)
    Column headers: Person, Avg FTE (to date), Avg FTE (remaining), Peak Month FTE, Status
    {num_team} person rows:
      - Person name: green cross-ref to Config roster (=Config!A<roster_row>)
      - Avg FTE (to date): SUMPRODUCT over Capacity TOTAL row up to current_month
      - Avg FTE (remaining): SUMPRODUCT over Capacity TOTAL row after current_month
      - Peak Month FTE: MAX of Capacity TOTAL row
      - Status: OVER/WATCH/OK based on remaining avg vs Config Max FTE
    Conditional formatting: green/amber/red on Status and Peak Month (per-person rules)

  BLOCK 3: Cumulative EUR LineChart (anchored 2 rows below TOTAL, ~15 rows tall)
    LineChart: Planned (blue solid) vs Actual (red solid) vs Ceiling (grey dashed)
    Data from hidden helper rows below person load section

  HIDDEN HELPER ROWS (below person load):
    Cumulative planned EUR per month (running sum of Plan FTE * fte_rate)
    Cumulative actual EUR per month (running sum of Actuals TOTAL row)
    Ceiling line (partner_ceiling repeated per month)
    Month-year labels (e.g. "Nov 25") for chart X-axis

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

  Config WP table (WP IDs, names, budgets) --> all sheets via cross-refs
  Config team roster (names, Max FTE) --> Capacity person headers + REMAINING formulas
  Config derived cols (Active Months, Personnel Ceiling) --> Actuals ceiling validation
  Capacity monthly FTE (input) --> Plan Planned FTE (SUM formulas across people)
  Actuals monthly EUR (input) --> Summary actual EUR aggregation
  Plan Planned FTE --> Summary planned FTE/EUR aggregation
  Actuals cumulative columns + Config WP table --> Trends burn rate & exhaustion
  Capacity TOTAL rows + Config roster --> Trends person load
  Plan Planned FTE + Actuals TOTAL --> Trends cumulative chart

  PRIMARY PLANNING SURFACE: Capacity sheet. Edit monthly FTE per person there.
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
  Config → Plan → Capacity → patch Plan → Trends skeleton → Actuals → Summary → Trends → _META
  Config sheet created first so named ranges (project_start, fte_rate) are available.
  Plan sheet built next, provides config_first_wp_row to downstream builders.
  Capacity sheet built, then Plan cells overwritten with SUM formulas referencing Capacity.
  Trends skeleton (B1 + current_month named range) created before Actuals so current_month is available.
  Trends content populated after Actuals/Summary so their layouts are known.

═══ NAMED RANGES ═══
  project_start   = Config!$B$2   (project start date)
  fte_rate         = Config!$B$4   (monthly EUR rate per person-month)
  partner_ceiling  = Plan!$I$1     (total personnel ceiling in EUR, =SUM of Config!G WP column)
  current_month    = Trends!$B$1   (last month with completed actuals, blue input)

FORMULAS USE ABSOLUTE REFERENCES for config lookups and mixed references for month iterations.

═══ MODIFICATION GUIDE (for AI agents and PMs) ═══

TO ADD A WORK PACKAGE:
  1. Add a row to tbl_WorkPackages on the Config sheet (WP ID, Name, FTE, Start, End)
  2. Run: python -m budget_planner rebuild
  3. Rebuild regenerates all sheets. Existing Capacity FTE and Actuals EUR are preserved.

TO ADD A TEAM MEMBER:
  1. Add a row to the Config sheet TEAM ROSTER section (Name, Max FTE)
  2. Run: python -m budget_planner rebuild

TO CHANGE PROJECT DURATION:
  1. Edit Project Months on the Config sheet (cell B3)
  2. Run: python -m budget_planner rebuild
  3. Adds/removes month columns. Existing data for months within the new range is preserved.

TO CHANGE FTE RATE:
  Edit cell B4 on Config sheet. All EUR calculations update via the fte_rate named range.

TO REPLACE EXAMPLE DATA:
  Edit Config sheet (WPs, team roster), Capacity config block
  (external commitments). Then run: python -m budget_planner rebuild

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
        cfg_roster_section=cfg_roster_section,
        cfg_roster_header=cfg_roster_header,
        cfg_roster_first=cfg_roster_first,
        act_last_wp=4 + num_wps - 1,
        act_total=4 + num_wps,
        actuals_amber_pct=int(0.85 * 100),
        actuals_amber_breach=12,
        actuals_red_breach=6,
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
