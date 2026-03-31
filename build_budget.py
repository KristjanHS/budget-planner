from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers, Protection
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from datetime import datetime
from dateutil.relativedelta import relativedelta

# ── CONSTANTS ──
PROJECT_MONTHS = 60
MONTH_COL_START = 3  # Column C
MONTH_COL_END = MONTH_COL_START + PROJECT_MONTHS - 1
VAL_COL_SUM = MONTH_COL_END + 1
VAL_COL_BUDGET = MONTH_COL_END + 2
VAL_COL_DIFF = MONTH_COL_END + 3
PROJECT_START = datetime(2026, 7, 1)
DEFAULT_RATE = 7500

EXAMPLE_WPS = [
    {"id": "WP1", "name": "Project Management", "fte": 4.0, "start": 1, "end": 60},
    {"id": "WP2", "name": "Requirements & Design", "fte": 8.0, "start": 1, "end": 18},
    {"id": "WP3", "name": "Core Development", "fte": 24.0, "start": 7, "end": 42},
    {"id": "WP4", "name": "Integration & Testing", "fte": 15.0, "start": 19, "end": 48},
    {"id": "WP5", "name": "Pilot Deployment", "fte": 10.0, "start": 37, "end": 54},
    {"id": "WP6", "name": "Dissemination", "fte": 6.0, "start": 13, "end": 60},
    {"id": "WP7", "name": "Ethics & Compliance", "fte": 4.0, "start": 1, "end": 60},
    {"id": "WP8", "name": "Exploitation", "fte": 5.0, "start": 25, "end": 60},
]
NUM_WPS = len(EXAMPLE_WPS)

# ── STYLES ──
BLUE_INPUT = Font(name="Arial", color="0000FF", size=10)
BLACK_FORMULA = Font(name="Arial", color="000000", size=10)
GREEN_XREF = Font(name="Arial", color="008000", size=10)
HEADER_FONT = Font(name="Arial", bold=True, size=10)
TITLE_FONT = Font(name="Arial", bold=True, size=14)
SECTION_FONT = Font(name="Arial", bold=True, size=12, color="FFFFFF")
SECTION_FILL = PatternFill("solid", fgColor="2F5496")
GREY_FILL = PatternFill("solid", fgColor="D9D9D9")
LIGHT_BLUE_FILL = PatternFill("solid", fgColor="D6E4F0")
LIGHT_GREEN_FILL = PatternFill("solid", fgColor="E2EFDA")
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)
CENTER = Alignment(horizontal="center", vertical="center")
WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)

def style_cell(ws, row, col, value=None, font=None, fill=None, alignment=None, fmt=None, border=True):
    cell = ws.cell(row=row, column=col, value=value)
    if font: cell.font = font
    if fill: cell.fill = fill
    if alignment: cell.alignment = alignment
    if fmt: cell.number_format = fmt
    if border: cell.border = THIN_BORDER
    return cell

def col_letter(c):
    return get_column_letter(c)

# ── WORKBOOK ──
wb = Workbook()

# ═══════════════════════════════════════════
# SHEET 1: _META (AI Agent Documentation)
# ═══════════════════════════════════════════
meta = wb.active
meta.title = "_META"
meta.sheet_properties.tabColor = "FF0000"

meta_text = """WORKBOOK: Programme Project Budget Planner
VERSION: 1.0
CREATED: {date}
PURPOSE: Track planned and actual FTE allocation per work package per month for an EU-funded (programme) project. Derive euro amounts from FTE using a configurable rate.

═══ SHEET INDEX ═══
1. _META        — This sheet. AI-agent-readable documentation of workbook structure, conventions, and modification instructions.
2. Plan         — Main operational sheet. Three blocks: CONFIG (WP definitions), PLANNED FTE (monthly allocation), ACTUAL FTE (logged time).
3. Summary      — Annual aggregation. Planned vs Actual FTE and EUR per WP per year. Budget health indicators.

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
    Cols {first_mcol}-{last_mcol} = FTE value for that month (NUMBER, input, blue text)
      - Pre-filled with even distribution: Total_FTE / Active_Months
      - Cells in inactive months pre-filled with 0
      - USER EDITS THESE DIRECTLY — they are values, not formulas
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
  NUMBER FORMATS:
    FTE values: 0.00
    EUR values: €#,##0
    Percentages: 0.0%

═══ MODIFICATION GUIDE (for AI agents) ═══

TO ADD A WORK PACKAGE:
  1. Insert a row in the config table (rows 6-{wp_last_config}), update totals row
  2. Insert a corresponding row in PLANNED block (rows {plan_first_wp}-{plan_last_wp})
  3. Insert a corresponding row in ACTUAL block (rows {act_first_wp}-{act_last_wp})
  4. Insert a row in Summary sheet (rows 5-{sum_last_wp})
  5. Update all SUM ranges in totals rows and Summary formulas
  6. Pre-fill new planned row with even distribution or 0

TO CHANGE PROJECT DURATION:
  1. Add/remove month columns in both PLANNED and ACTUAL blocks (cols {first_mcol}-{last_mcol})
  2. Update year header merges
  3. Update validation SUM ranges
  4. Update Summary sheet year ranges
  5. Update month number headers

TO CHANGE FTE RATE:
  Edit cell E2 on Plan sheet. All EUR calculations update automatically.

TO REPLACE EXAMPLE DATA:
  Edit config rows 6-{wp_last_config}: change WP IDs, names, FTE budgets, start/end months.
  Then either manually update planned FTE cells or run a script to re-distribute evenly.

NAMED RANGES:
  project_start = Plan!$B$2
  fte_rate = Plan!$E$2

FORMULAS USE ABSOLUTE REFERENCES for config lookups and mixed references for month iterations.
Row/column indices are documented in comments within the spreadsheet.
""".format(
    date=datetime.now().strftime("%Y-%m-%d"),
    config_end=5 + NUM_WPS + 1,
    wp_last_config=5 + NUM_WPS,
    config_total=5 + NUM_WPS + 1,
    pm=PROJECT_MONTHS,
    plan_header=5 + NUM_WPS + 3,
    plan_year_row=5 + NUM_WPS + 4,
    plan_month_row=5 + NUM_WPS + 5,
    plan_first_wp=5 + NUM_WPS + 6,
    plan_last_wp=5 + NUM_WPS + 5 + NUM_WPS,
    plan_total=5 + NUM_WPS + 6 + NUM_WPS,
    act_header=5 + NUM_WPS + 6 + NUM_WPS + 2,
    act_first_wp=5 + NUM_WPS + 6 + NUM_WPS + 5,
    act_last_wp=5 + NUM_WPS + 6 + NUM_WPS + 4 + NUM_WPS,
    act_total=5 + NUM_WPS + 6 + NUM_WPS + 5 + NUM_WPS,
    first_mcol=col_letter(MONTH_COL_START),
    last_mcol=col_letter(MONTH_COL_END),
    val_sum_col=col_letter(VAL_COL_SUM),
    val_budget_col=col_letter(VAL_COL_BUDGET),
    val_diff_col=col_letter(VAL_COL_DIFF),
    sum_last_wp=4 + NUM_WPS,
    sum_total=5 + NUM_WPS,
)

meta["A1"] = meta_text
meta["A1"].font = Font(name="Consolas", size=10)
meta["A1"].alignment = WRAP
meta.column_dimensions["A"].width = 120

# ═══════════════════════════════════════════
# SHEET 2: Plan
# ═══════════════════════════════════════════
plan = wb.create_sheet("Plan")
plan.sheet_properties.tabColor = "2F5496"

# ── Row tracking ──
r = 1

# ── GLOBAL PARAMETERS ──
style_cell(plan, r, 1, "Project Title:", font=HEADER_FONT, border=False)
style_cell(plan, r, 2, "Programme Project Budget Plan", font=Font(name="Arial", bold=True, size=14, color="0000FF"), border=False)
r += 1
style_cell(plan, r, 1, "Project Start Date:", font=HEADER_FONT, border=False)
c = style_cell(plan, r, 2, PROJECT_START, font=BLUE_INPUT, border=False, fmt="MMM YYYY")
style_cell(plan, r, 4, "Monthly FTE Rate (€):", font=HEADER_FONT, border=False)
style_cell(plan, r, 5, DEFAULT_RATE, font=BLUE_INPUT, border=False, fmt="€#,##0")
r += 1

# Named ranges
from openpyxl.workbook.defined_name import DefinedName
wb.defined_names.add(DefinedName("project_start", attr_text="Plan!$B$2"))
wb.defined_names.add(DefinedName("fte_rate", attr_text="Plan!$E$2"))

# ── CONFIG TABLE ──
r += 1
CONFIG_HEADER_ROW = r
plan.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
style_cell(plan, r, 1, "CONFIGURATION", font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER)
for c2 in range(2, 8):
    style_cell(plan, r, c2, fill=SECTION_FILL)
r += 1

CONFIG_COL_HEADERS_ROW = r
config_headers = ["WP ID", "WP Name", "Total FTE (PM)", "Start Month", "End Month", "Active Months", "Budget (€)"]
for ci, h in enumerate(config_headers, 1):
    style_cell(plan, r, ci, h, font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
r += 1

CONFIG_FIRST_WP_ROW = r
for i, wp in enumerate(EXAMPLE_WPS):
    row = r + i
    style_cell(plan, row, 1, wp["id"], font=BLUE_INPUT)
    style_cell(plan, row, 2, wp["name"], font=BLUE_INPUT)
    style_cell(plan, row, 3, wp["fte"], font=BLUE_INPUT, fmt="0.00")
    style_cell(plan, row, 4, wp["start"], font=BLUE_INPUT, fmt="0")
    style_cell(plan, row, 5, wp["end"], font=BLUE_INPUT, fmt="0")
    style_cell(plan, row, 6, font=BLACK_FORMULA, fmt="0").value = f"=E{row}-D{row}+1"
    style_cell(plan, row, 7, font=BLACK_FORMULA, fmt="€#,##0").value = f"=C{row}*$E$2"
CONFIG_LAST_WP_ROW = CONFIG_FIRST_WP_ROW + NUM_WPS - 1
r = CONFIG_LAST_WP_ROW + 1

CONFIG_TOTAL_ROW = r
style_cell(plan, r, 1, "", fill=LIGHT_BLUE_FILL)
style_cell(plan, r, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_BLUE_FILL)
style_cell(plan, r, 3, font=BLACK_FORMULA, fill=LIGHT_BLUE_FILL, fmt="0.00").value = f"=SUM(C{CONFIG_FIRST_WP_ROW}:C{CONFIG_LAST_WP_ROW})"
style_cell(plan, r, 4, "", fill=LIGHT_BLUE_FILL)
style_cell(plan, r, 5, "", fill=LIGHT_BLUE_FILL)
style_cell(plan, r, 6, "", fill=LIGHT_BLUE_FILL)
style_cell(plan, r, 7, font=BLACK_FORMULA, fill=LIGHT_BLUE_FILL, fmt="€#,##0").value = f"=SUM(G{CONFIG_FIRST_WP_ROW}:G{CONFIG_LAST_WP_ROW})"
r += 2

# ── HELPER: Build a month block (planned or actuals) ──
def build_month_block(ws, start_row, block_title, is_actuals=False):
    r = start_row

    # Section header
    merge_end = VAL_COL_DIFF
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=merge_end)
    style_cell(ws, r, 1, block_title, font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER)
    for c2 in range(2, merge_end + 1):
        style_cell(ws, r, c2, fill=SECTION_FILL)
    r += 1
    YEAR_ROW = r

    # Year header row
    style_cell(ws, r, 1, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, 2, "", fill=LIGHT_BLUE_FILL)
    m = 0
    while m < PROJECT_MONTHS:
        year_start_col = MONTH_COL_START + m
        months_left = PROJECT_MONTHS - m
        span = min(12, months_left)
        year_end_col = year_start_col + span - 1
        cal_date = PROJECT_START + relativedelta(months=m)
        year_label = f"Year {m // 12 + 1} ({cal_date.year})"
        if span > 1:
            ws.merge_cells(start_row=r, start_column=year_start_col, end_row=r, end_column=year_end_col)
        style_cell(ws, r, year_start_col, year_label, font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
        for c2 in range(year_start_col + 1, year_end_col + 1):
            style_cell(ws, r, c2, fill=LIGHT_BLUE_FILL)
        m += span
    # Validation headers in year row
    style_cell(ws, r, VAL_COL_SUM, "Sum", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_BUDGET, "Budget FTE", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    label = "Remaining" if is_actuals else "Variance"
    style_cell(ws, r, VAL_COL_DIFF, label, font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    r += 1
    MONTH_ROW = r

    # Month header row
    style_cell(ws, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    for m in range(PROJECT_MONTHS):
        c = MONTH_COL_START + m
        cal_date = PROJECT_START + relativedelta(months=m)
        label = f"M{m+1}\n{cal_date.strftime('%b %y')}"
        style_cell(ws, r, c, label, font=Font(name="Arial", size=8), fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, VAL_COL_SUM, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, VAL_COL_BUDGET, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, r, VAL_COL_DIFF, "", fill=LIGHT_BLUE_FILL)
    r += 1
    FIRST_WP_ROW = r

    # WP rows
    for i, wp in enumerate(EXAMPLE_WPS):
        row = r + i
        config_row = CONFIG_FIRST_WP_ROW + i
        style_cell(ws, row, 1, font=GREEN_XREF).value = f"=A{config_row}"
        style_cell(ws, row, 2, font=GREEN_XREF).value = f"=B{config_row}"

        active_months = wp["end"] - wp["start"] + 1
        even_fte = round(wp["fte"] / active_months, 4) if active_months > 0 else 0

        for m in range(PROJECT_MONTHS):
            c = MONTH_COL_START + m
            month_num = m + 1
            if wp["start"] <= month_num <= wp["end"]:
                val = even_fte if not is_actuals else 0
                style_cell(ws, row, c, val, font=BLUE_INPUT, fmt="0.00")
            else:
                style_cell(ws, row, c, "", font=BLUE_INPUT, fill=GREY_FILL, fmt="0.00")

        # Validation columns
        sum_range = f"{col_letter(MONTH_COL_START)}{row}:{col_letter(MONTH_COL_END)}{row}"
        style_cell(ws, row, VAL_COL_SUM, font=BLACK_FORMULA, fmt="0.00").value = f"=SUM({sum_range})"
        style_cell(ws, row, VAL_COL_BUDGET, font=GREEN_XREF, fmt="0.00").value = f"=C{config_row}"
        if is_actuals:
            style_cell(ws, row, VAL_COL_DIFF, font=BLACK_FORMULA, fmt="0.00").value = f"={col_letter(VAL_COL_BUDGET)}{row}-{col_letter(VAL_COL_SUM)}{row}"
        else:
            style_cell(ws, row, VAL_COL_DIFF, font=BLACK_FORMULA, fmt="0.00").value = f"={col_letter(VAL_COL_BUDGET)}{row}-{col_letter(VAL_COL_SUM)}{row}"

    LAST_WP_ROW = FIRST_WP_ROW + NUM_WPS - 1
    r = LAST_WP_ROW + 1

    # Totals row
    TOTAL_ROW = r
    style_cell(ws, r, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, r, 2, "MONTHLY TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
    for m in range(PROJECT_MONTHS):
        c = MONTH_COL_START + m
        col_range = f"{col_letter(c)}{FIRST_WP_ROW}:{col_letter(c)}{LAST_WP_ROW}"
        style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt="0.00").value = f"=SUM({col_range})"
    sum_range = f"{col_letter(VAL_COL_SUM)}{FIRST_WP_ROW}:{col_letter(VAL_COL_SUM)}{LAST_WP_ROW}"
    style_cell(ws, r, VAL_COL_SUM, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt="0.00").value = f"=SUM({sum_range})"
    bud_range = f"{col_letter(VAL_COL_BUDGET)}{FIRST_WP_ROW}:{col_letter(VAL_COL_BUDGET)}{LAST_WP_ROW}"
    style_cell(ws, r, VAL_COL_BUDGET, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt="0.00").value = f"=SUM({bud_range})"
    diff_range = f"{col_letter(VAL_COL_DIFF)}{FIRST_WP_ROW}:{col_letter(VAL_COL_DIFF)}{LAST_WP_ROW}"
    style_cell(ws, r, VAL_COL_DIFF, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt="0.00").value = f"=SUM({diff_range})"

    # Conditional formatting: red fill on variance/remaining if negative (for actuals) or nonzero (for planned)
    diff_col_l = col_letter(VAL_COL_DIFF)
    cell_range = f"{diff_col_l}{FIRST_WP_ROW}:{diff_col_l}{LAST_WP_ROW}"
    if is_actuals:
        ws.conditional_formatting.add(cell_range, FormulaRule(
            formula=[f"{diff_col_l}{FIRST_WP_ROW}<0"], fill=RED_FILL
        ))
    else:
        ws.conditional_formatting.add(cell_range, FormulaRule(
            formula=[f"ABS({diff_col_l}{FIRST_WP_ROW})>0.01"], fill=YELLOW_FILL
        ))

    return {
        "header_row": start_row,
        "year_row": YEAR_ROW,
        "month_row": MONTH_ROW,
        "first_wp_row": FIRST_WP_ROW,
        "last_wp_row": LAST_WP_ROW,
        "total_row": TOTAL_ROW,
        "next_row": r + 2
    }

# Build planned block
planned = build_month_block(plan, r, "PLANNED FTE (PERSON-MONTHS)", is_actuals=False)

# Build actuals block
actuals = build_month_block(plan, planned["next_row"], "ACTUAL FTE (PERSON-MONTHS)", is_actuals=True)

# Column widths
plan.column_dimensions["A"].width = 10
plan.column_dimensions["B"].width = 28
for m in range(PROJECT_MONTHS):
    plan.column_dimensions[col_letter(MONTH_COL_START + m)].width = 8
plan.column_dimensions[col_letter(VAL_COL_SUM)].width = 12
plan.column_dimensions[col_letter(VAL_COL_BUDGET)].width = 12
plan.column_dimensions[col_letter(VAL_COL_DIFF)].width = 12

# Freeze panes: freeze WP ID/Name columns and header rows for planned block
plan.freeze_panes = f"C{planned['first_wp_row']}"

# ═══════════════════════════════════════════
# SHEET 3: Summary
# ═══════════════════════════════════════════
summary = wb.create_sheet("Summary")
summary.sheet_properties.tabColor = "008000"

r = 1
style_cell(summary, r, 1, "ANNUAL BUDGET SUMMARY", font=TITLE_FONT, border=False)
r += 2

# Calculate years
num_years = (PROJECT_MONTHS + 11) // 12

# Year group headers
SUMMARY_WP_COL = 1
SUMMARY_NAME_COL = 2
YEAR_GROUP_START = 3
COLS_PER_YEAR = 5  # Planned FTE, Actual FTE, Variance FTE, Planned EUR, Actual EUR

# Project total group
TOTAL_GROUP_START = YEAR_GROUP_START + num_years * COLS_PER_YEAR
TOTAL_COLS = 7  # Total Planned FTE, Total Actual FTE, Total Variance, Total Planned EUR, Total Actual EUR, Remaining EUR, % Consumed

# Year headers row
YEAR_HEADER_ROW = r
style_cell(summary, r, 1, "", fill=LIGHT_BLUE_FILL)
style_cell(summary, r, 2, "", fill=LIGHT_BLUE_FILL)
for y in range(num_years):
    start_col = YEAR_GROUP_START + y * COLS_PER_YEAR
    end_col = start_col + COLS_PER_YEAR - 1
    cal_year = PROJECT_START.year + y
    summary.merge_cells(start_row=r, start_column=start_col, end_row=r, end_column=end_col)
    style_cell(summary, r, start_col, f"Year {y+1} ({cal_year})", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    for c2 in range(start_col + 1, end_col + 1):
        style_cell(summary, r, c2, fill=LIGHT_BLUE_FILL)

summary.merge_cells(start_row=r, start_column=TOTAL_GROUP_START, end_row=r, end_column=TOTAL_GROUP_START + TOTAL_COLS - 1)
style_cell(summary, r, TOTAL_GROUP_START, "PROJECT TOTAL", font=HEADER_FONT, fill=SECTION_FILL, alignment=CENTER)
style_cell(summary, r, TOTAL_GROUP_START, font=Font(name="Arial", bold=True, size=10, color="FFFFFF"), fill=SECTION_FILL, alignment=CENTER).value = "PROJECT TOTAL"
for c2 in range(TOTAL_GROUP_START + 1, TOTAL_GROUP_START + TOTAL_COLS):
    style_cell(summary, r, c2, fill=SECTION_FILL)
r += 1

# Sub-headers row
SUB_HEADER_ROW = r
style_cell(summary, r, 1, "WP ID", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
style_cell(summary, r, 2, "WP Name", font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
year_sub = ["Plan FTE", "Act FTE", "Var FTE", "Plan €", "Act €"]
for y in range(num_years):
    for si, sh in enumerate(year_sub):
        c = YEAR_GROUP_START + y * COLS_PER_YEAR + si
        style_cell(summary, r, c, sh, font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)

total_sub = ["Plan FTE", "Act FTE", "Var FTE", "Plan €", "Act €", "Rem €", "% Used"]
for si, sh in enumerate(total_sub):
    c = TOTAL_GROUP_START + si
    style_cell(summary, r, c, sh, font=HEADER_FONT, fill=LIGHT_GREEN_FILL, alignment=CENTER)
r += 1

SUMMARY_FIRST_WP_ROW = r

# WP data rows
for i, wp in enumerate(EXAMPLE_WPS):
    row = r + i
    config_row = CONFIG_FIRST_WP_ROW + i
    style_cell(summary, row, 1, font=GREEN_XREF).value = f"=Plan!A{config_row}"
    style_cell(summary, row, 2, font=GREEN_XREF).value = f"=Plan!B{config_row}"

    plan_wp_row = planned["first_wp_row"] + i
    act_wp_row = actuals["first_wp_row"] + i

    for y in range(num_years):
        m_start = y * 12
        m_end = min((y + 1) * 12, PROJECT_MONTHS) - 1
        first_col_l = col_letter(MONTH_COL_START + m_start)
        last_col_l = col_letter(MONTH_COL_START + m_end)

        sc = YEAR_GROUP_START + y * COLS_PER_YEAR
        # Planned FTE for year
        style_cell(summary, row, sc, font=GREEN_XREF, fmt="0.00").value = f"=SUM(Plan!{first_col_l}{plan_wp_row}:{last_col_l}{plan_wp_row})"
        # Actual FTE for year
        style_cell(summary, row, sc + 1, font=GREEN_XREF, fmt="0.00").value = f"=SUM(Plan!{first_col_l}{act_wp_row}:{last_col_l}{act_wp_row})"
        # Variance
        style_cell(summary, row, sc + 2, font=BLACK_FORMULA, fmt="0.00").value = f"={col_letter(sc)}{row}-{col_letter(sc+1)}{row}"
        # Planned EUR
        style_cell(summary, row, sc + 3, font=BLACK_FORMULA, fmt="€#,##0").value = f"={col_letter(sc)}{row}*Plan!$E$2"
        # Actual EUR
        style_cell(summary, row, sc + 4, font=BLACK_FORMULA, fmt="€#,##0").value = f"={col_letter(sc+1)}{row}*Plan!$E$2"

    # Project totals
    tc = TOTAL_GROUP_START
    # Sum planned FTE across years
    plan_year_cols = [col_letter(YEAR_GROUP_START + y * COLS_PER_YEAR) for y in range(num_years)]
    plan_sum = "+".join([f"{cl}{row}" for cl in plan_year_cols])
    style_cell(summary, row, tc, font=BLACK_FORMULA, fmt="0.00").value = f"={plan_sum}"
    # Sum actual FTE
    act_year_cols = [col_letter(YEAR_GROUP_START + y * COLS_PER_YEAR + 1) for y in range(num_years)]
    act_sum = "+".join([f"{cl}{row}" for cl in act_year_cols])
    style_cell(summary, row, tc + 1, font=BLACK_FORMULA, fmt="0.00").value = f"={act_sum}"
    # Total variance
    style_cell(summary, row, tc + 2, font=BLACK_FORMULA, fmt="0.00").value = f"={col_letter(tc)}{row}-{col_letter(tc+1)}{row}"
    # Total planned EUR
    style_cell(summary, row, tc + 3, font=BLACK_FORMULA, fmt="€#,##0").value = f"={col_letter(tc)}{row}*Plan!$E$2"
    # Total actual EUR
    style_cell(summary, row, tc + 4, font=BLACK_FORMULA, fmt="€#,##0").value = f"={col_letter(tc+1)}{row}*Plan!$E$2"
    # Remaining EUR
    style_cell(summary, row, tc + 5, font=BLACK_FORMULA, fmt="€#,##0").value = f"=Plan!G{config_row}-{col_letter(tc+4)}{row}"
    # % consumed
    style_cell(summary, row, tc + 6, font=BLACK_FORMULA, fmt="0.0%").value = f"=IF(Plan!C{config_row}=0,0,{col_letter(tc+1)}{row}/Plan!C{config_row})"

SUMMARY_LAST_WP_ROW = SUMMARY_FIRST_WP_ROW + NUM_WPS - 1
r = SUMMARY_LAST_WP_ROW + 1

# Totals row
SUMMARY_TOTAL_ROW = r
style_cell(summary, r, 1, "", fill=LIGHT_GREEN_FILL)
style_cell(summary, r, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_GREEN_FILL)
for y in range(num_years):
    for si in range(COLS_PER_YEAR):
        c = YEAR_GROUP_START + y * COLS_PER_YEAR + si
        fmt = "€#,##0" if si >= 3 else "0.00"
        col_range = f"{col_letter(c)}{SUMMARY_FIRST_WP_ROW}:{col_letter(c)}{SUMMARY_LAST_WP_ROW}"
        style_cell(summary, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=fmt).value = f"=SUM({col_range})"

for si in range(TOTAL_COLS):
    c = TOTAL_GROUP_START + si
    fmt = "0.0%" if si == 6 else ("€#,##0" if si >= 3 else "0.00")
    col_range = f"{col_letter(c)}{SUMMARY_FIRST_WP_ROW}:{col_letter(c)}{SUMMARY_LAST_WP_ROW}"
    if si == 6:
        style_cell(summary, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=fmt).value = \
            f"=IF({col_letter(TOTAL_GROUP_START)}{r}=0,0,{col_letter(TOTAL_GROUP_START+1)}{r}/{col_letter(TOTAL_GROUP_START)}{r})"
    else:
        style_cell(summary, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL, fmt=fmt).value = f"=SUM({col_range})"

# Conditional formatting: red on negative remaining EUR
rem_col = col_letter(TOTAL_GROUP_START + 5)
rem_range = f"{rem_col}{SUMMARY_FIRST_WP_ROW}:{rem_col}{SUMMARY_LAST_WP_ROW}"
summary.conditional_formatting.add(rem_range, FormulaRule(
    formula=[f"{rem_col}{SUMMARY_FIRST_WP_ROW}<0"], fill=RED_FILL
))

# Column widths
summary.column_dimensions["A"].width = 10
summary.column_dimensions["B"].width = 28
for c in range(YEAR_GROUP_START, TOTAL_GROUP_START + TOTAL_COLS + 1):
    summary.column_dimensions[col_letter(c)].width = 12

summary.freeze_panes = f"C{SUMMARY_FIRST_WP_ROW}"

# ── SAVE ──
OUTPUT = "/home/claude/budget_plan.xlsx"
wb.save(OUTPUT)
print(f"Saved to {OUTPUT}")

# Print key row references for verification
print(f"\n── KEY ROWS (Plan sheet) ──")
print(f"Config: rows {CONFIG_FIRST_WP_ROW}-{CONFIG_LAST_WP_ROW}")
print(f"Planned: rows {planned['first_wp_row']}-{planned['last_wp_row']}")
print(f"Actuals: rows {actuals['first_wp_row']}-{actuals['last_wp_row']}")
print(f"Month cols: {col_letter(MONTH_COL_START)}-{col_letter(MONTH_COL_END)}")
print(f"Validation cols: {col_letter(VAL_COL_SUM)}, {col_letter(VAL_COL_BUDGET)}, {col_letter(VAL_COL_DIFF)}")
