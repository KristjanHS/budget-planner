"""Build the _META sheet — AI-agent-readable workbook documentation."""
from openpyxl.styles import Font
from datetime import datetime

from budget_planner.styles import WRAP, col_letter
from budget_planner.config import (
    PROJECT_MONTHS, NUM_WPS, MONTH_COL_START, MONTH_COL_END,
    VAL_COL_SUM, VAL_COL_BUDGET, VAL_COL_DIFF,
    TEAM, EXTERNAL_COMMITMENTS,
)


def build_meta_sheet(ws, layout):
    """Populate the _META worksheet with structural documentation.

    Parameters
    ----------
    ws : openpyxl.worksheet.worksheet.Worksheet
        The _META sheet (already created by the caller).
    layout : dict
        Row-number mapping produced by the plan-sheet builder.  Keys used:
        config_total_row, config_last_wp_row, planned_header_row,
        planned_first_wp_row, planned_last_wp_row, planned_total_row,
        actual_header_row, actual_first_wp_row, actual_last_wp_row,
        actual_total_row.
    """
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
    act_header = layout["actual_header_row"]
    act_first_wp = layout["actual_first_wp_row"]
    act_last_wp = layout["actual_last_wp_row"]
    act_total = layout["actual_total_row"]

    # Summary sheet references
    sum_last_wp = 4 + NUM_WPS
    sum_total = 5 + NUM_WPS

    # Team sheet references
    num_team = len(TEAM)
    roster_last = 2 + num_team  # header rows 1-2, data starts row 3
    num_ext = len(EXTERNAL_COMMITMENTS)

    # Column letters
    first_mcol = col_letter(MONTH_COL_START)
    last_mcol = col_letter(MONTH_COL_END)
    val_sum_col = col_letter(VAL_COL_SUM)
    val_budget_col = col_letter(VAL_COL_BUDGET)
    val_diff_col = col_letter(VAL_COL_DIFF)

    meta_text = """WORKBOOK: Programme Project Budget Planner
VERSION: 2.0
CREATED: {date}
PURPOSE: Track planned and actual FTE allocation per work package per month for an EU-funded (programme) project. Per-person capacity planning with cross-project visibility. Derive euro amounts from FTE using a configurable rate.

═══ SHEET INDEX ═══
1. _META        — This sheet. AI-agent-readable documentation of workbook structure, conventions, and modification instructions.
2. Plan         — Config table + Planned FTE (formulas from Capacity) + Actual FTE (input).
3. Team         — Person roster + WP share seeds for initial allocation.
4. Capacity     — Primary planning surface: per-person monthly FTE allocation across programme WPs and external projects.
5. Summary      — Annual aggregation. Planned vs Actual FTE and EUR per WP per year. Budget health indicators.

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

BLOCK 1: GLOBAL PARAMETERS (Rows 1-2)
  B1 = Project title (text, input)
  B2 = Project start date (date, input). All month columns derive calendar dates from this.
  E2 = Monthly FTE rate in EUR (number, input). Used to derive euro amounts.

BLOCK 2: CONFIGURATION TABLE (Rows 4-{config_end})
  Row 4 = Section header "CONFIGURATION"
  Row 5 = Column headers
  Rows 6-{wp_last_config} = One row per WP
  Columns:
    A = WP ID (text, input)
    B = WP Name (text, input)
    C = Total FTE budget in person-months (number, input)
    D = Start month number 1-{pm} (integer, input)
    E = End month number 1-{pm} (integer, input)
    F = Active months (formula: E-D+1)
    G = Budget EUR (formula: C * $E$2)
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

BLOCK 4: ACTUAL FTE (Rows {act_header}-{act_total})
  Same structure as PLANNED block.
  Monthly cells start at 0 (input, blue text). User enters logged FTE here.
  Validation columns:
    {val_sum_col} = Cumulative actual (SUM formula)
    {val_budget_col} = Budget FTE from config (formula, green cross-ref)
    {val_diff_col} = Remaining: Budget - Cumulative (formula). Negative = OVERRUN.
  Row {act_total} = Column totals

═══ SUMMARY SHEET STRUCTURE ═══
  Row 1 = Title
  Row 3 = Year group headers (Year 1..Year 5), each spanning 5 sub-columns + Project Total group
  Row 4 = Sub-headers per year: Planned FTE, Actual FTE, Variance FTE, Planned EUR, Actual EUR
  Rows 5-{sum_last_wp} = One row per WP (all formulas referencing Plan sheet)
  Row {sum_total} = Totals
  Right-most group (Project Total): aggregates across all years + Remaining FTE, Remaining EUR, % Consumed

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
    EUR values: #,##0
    Percentages: 0.0%

═══ DATA FLOW ═══
  Team roster (names, Max FTE) --> Capacity person headers + REMAINING formulas
  Team WP shares --> Capacity seed values (generation-time only)
  Capacity monthly FTE (input) --> Plan Planned FTE (SUM formulas across people)
  Plan Planned FTE --> Summary annual aggregation
  Plan Actual FTE (input) --> Summary actual aggregation

  PRIMARY PLANNING SURFACE: Capacity sheet. Edit monthly FTE per person there.
  Plan Planned FTE is formula-driven — do not edit directly.

═══ MODIFICATION GUIDE (for AI agents) ═══

TO ADD A WORK PACKAGE:
  1. Insert a row in the config table (rows 6-{wp_last_config}), update totals row
  2. Insert a WP row in each person's group on the Capacity sheet
  3. Plan PLANNED block rows {plan_first_wp}-{plan_last_wp} are rebuilt automatically on --update
  4. Insert a corresponding row in ACTUAL block (rows {act_first_wp}-{act_last_wp})
  5. Insert a row in Summary sheet (rows 5-{sum_last_wp})
  6. Update Team sheet share matrix columns
  7. Regeneration via --update is recommended (rebuilds formula sheets)

TO ADD A TEAM MEMBER:
  1. Add a row to Team sheet roster (rows 3-{roster_last})
  2. Add a row to Team sheet share matrix
  3. Add a person group to Capacity sheet (header + WP rows + external rows + TOTAL + REMAINING)
  4. Plan Planned FTE formulas are rebuilt on --update to include the new person

TO CHANGE PROJECT DURATION:
  1. Add/remove month columns in Capacity and Plan sheets (cols {first_mcol}-{last_mcol})
  2. Update year header merges
  3. Update validation SUM ranges
  4. Update Summary sheet year ranges
  5. Update month number headers

TO CHANGE FTE RATE:
  Edit cell E2 on Plan sheet. All EUR calculations update automatically.

TO REPLACE EXAMPLE DATA:
  Edit config rows 6-{wp_last_config}: change WP IDs, names, FTE budgets, start/end months.
  Edit Team roster and share matrix. Edit Capacity monthly values.
  Recommended: use --update mode to regenerate formula sheets from edited data.

NAMED RANGES:
  project_start = Plan!$B$2
  fte_rate = Plan!$E$2

FORMULAS USE ABSOLUTE REFERENCES for config lookups and mixed references for month iterations.
Row/column indices are documented in comments within the spreadsheet.

SCENARIO PLANNING:
  Use --snapshot to save current state, edit Capacity freely, --restore to revert.
  No dedicated draft sheet — file-level snapshots provide safe replanning.
""".format(
        date=datetime.now().strftime("%Y-%m-%d"),
        config_end=config_end,
        wp_last_config=wp_last_config,
        config_total=config_total,
        pm=PROJECT_MONTHS,
        plan_header=plan_header,
        plan_year_row=plan_year_row,
        plan_month_row=plan_month_row,
        plan_first_wp=plan_first_wp,
        plan_last_wp=plan_last_wp,
        plan_total=plan_total,
        act_header=act_header,
        act_first_wp=act_first_wp,
        act_last_wp=act_last_wp,
        act_total=act_total,
        first_mcol=first_mcol,
        last_mcol=last_mcol,
        val_sum_col=val_sum_col,
        val_budget_col=val_budget_col,
        val_diff_col=val_diff_col,
        sum_last_wp=sum_last_wp,
        sum_total=sum_total,
        roster_last=roster_last,
        num_team=num_team,
        num_ext=num_ext,
    )

    ws["A1"] = meta_text
    ws["A1"].font = Font(name="Consolas", size=10)
    ws["A1"].alignment = WRAP
    ws.column_dimensions["A"].width = 120
    ws.sheet_properties.tabColor = "FF0000"
