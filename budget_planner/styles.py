"""Spreadsheet styling constants and helpers."""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter

# Font color convention: Blue = input, Black = formula, Green = cross-sheet ref
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
AMBER_FILL = PatternFill("solid", fgColor="FFD966")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")

# Indirect-view header fills. External-facing EUR columns (+i = with 25% indirect)
# fill green; internal reconciliation EUR columns (-i = without indirect) fill
# orange. Non-EUR headers keep LIGHT_BLUE_FILL.
WITH_INDIRECT_HEADER_FILL = PatternFill("solid", fgColor="E2EFDA")
WITHOUT_INDIRECT_HEADER_FILL = PatternFill("solid", fgColor="FBE5D6")

# CF (conditional formatting) fills — DXF context requires bgColor, not fgColor
# Use 8-char AARRGGBB hex; 6-char gets 00 alpha (transparent)
CF_YELLOW_FILL = PatternFill(bgColor="FFFFFF00")
CF_AMBER_FILL = PatternFill(bgColor="FFFFD966")
CF_RED_FILL = PatternFill(bgColor="FFFFC7CE")
CF_GREEN_FILL = PatternFill(bgColor="FFE2EFDA")
CF_GREY_FILL = PatternFill(bgColor="FFD9D9D9")

# Timelog Plan: darker grey for cells outside any person-WP window
# (out-of-project-window). GREY_FILL is the in-window-but-empty marker.
DARK_GREY_FILL = PatternFill("solid", fgColor="A6A6A6")

# Timelog Plan: pale yellow for vacation-month grid cells. Distinct tier —
# direct fill (no CF), locked cells, always empty. See
# `docs/spec.md` § Vacation cells are always locked and empty.
VACATION_FILL = PatternFill("solid", fgColor="FFF2CC")

# Timelog Plan: pastel palette cycled across WPs for the canvas background.
# Cell-formatting variant (fgColor; used if pre-fill is preferred).
PASTEL_FILLS = [
    PatternFill("solid", fgColor="DCE6F1"),  # pastel blue
    PatternFill("solid", fgColor="E2EFDA"),  # pastel green
    PatternFill("solid", fgColor="FFF2CC"),  # pastel yellow
    PatternFill("solid", fgColor="FCE4D6"),  # pastel orange
    PatternFill("solid", fgColor="EDEDED"),  # pastel grey
    PatternFill("solid", fgColor="E4DFEC"),  # pastel purple
    PatternFill("solid", fgColor="DAEEF3"),  # pastel cyan
    PatternFill("solid", fgColor="FCD5B4"),  # pastel salmon
    PatternFill("solid", fgColor="C6EFCE"),  # pastel mint
]
# CF variant (bgColor with explicit FF alpha, per CF_*_FILL convention).
CF_PASTEL_FILLS = [
    PatternFill(bgColor="FFDCE6F1"),
    PatternFill(bgColor="FFE2EFDA"),
    PatternFill(bgColor="FFFFF2CC"),
    PatternFill(bgColor="FFFCE4D6"),
    PatternFill(bgColor="FFEDEDED"),
    PatternFill(bgColor="FFE4DFEC"),
    PatternFill(bgColor="FFDAEEF3"),
    PatternFill(bgColor="FFFCD5B4"),
    PatternFill(bgColor="FFC6EFCE"),
]

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin")
)
CENTER = Alignment(horizontal="center", vertical="center")
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")
WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)


def style_cell(ws, row, col, value=None, font=None, fill=None, alignment=None, fmt=None, border=True, unlocked=False):
    """Apply styling to a worksheet cell."""
    cell = ws.cell(row=row, column=col, value=value)
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if fmt:
        cell.number_format = fmt
    if border:
        cell.border = THIN_BORDER
    # Auto-unlock blue input cells for sheet protection
    if unlocked or _is_blue_font(font):
        cell.protection = Protection(locked=False)
    return cell


def _is_blue_font(font):
    """Check if font color is blue (0000FF), indicating a user input cell."""
    if font is None:
        return False
    color = getattr(font, "color", None)
    if color is None:
        return False
    rgb = getattr(color, "rgb", None)
    if rgb is None:
        return False
    return str(rgb).upper().endswith("0000FF")


def col_letter(c):
    """Wrapper for get_column_letter."""
    return get_column_letter(c)


# Comment box sizing — openpyxl defaults (144×79 pt) truncate any comment
# longer than ~2 short lines. These helpers compute a box big enough for the
# actual text. Heuristic: 7pt per char wide, 16pt per line tall, with a
# min/max clamp so single-word comments don't get a postage-stamp box and
# paragraph-length ones don't cover the whole screen.
_COMMENT_CHAR_WIDTH_PT = 7
_COMMENT_LINE_HEIGHT_PT = 16
_COMMENT_MIN_WIDTH_PT = 220
_COMMENT_MAX_WIDTH_PT = 520
_COMMENT_MIN_HEIGHT_PT = 60
_COMMENT_MAX_HEIGHT_PT = 400


def sized_comment(text: str, author: str = "Programme Budget Tool"):
    """Return an openpyxl Comment whose box is sized to fit ``text``.

    openpyxl's default box (144×79 pt) silently clips most multi-sentence
    comments. This helper estimates a box large enough for the full text,
    clamped to a readable range. Width is chosen from the longest wrapped
    line; height from the resulting wrapped line count.
    """
    from openpyxl.comments import Comment as _Comment

    lines = text.splitlines() or [""]
    longest = max((len(line) for line in lines), default=0)
    width = min(_COMMENT_MAX_WIDTH_PT, max(_COMMENT_MIN_WIDTH_PT, longest * _COMMENT_CHAR_WIDTH_PT))
    # Chars that fit on one visual line at the chosen width.
    chars_per_line = max(1, width // _COMMENT_CHAR_WIDTH_PT)
    wrapped_lines = 0
    for line in lines:
        if not line:
            wrapped_lines += 1
        else:
            wrapped_lines += -(-len(line) // chars_per_line)  # ceil div
    height = min(_COMMENT_MAX_HEIGHT_PT, max(_COMMENT_MIN_HEIGHT_PT, wrapped_lines * _COMMENT_LINE_HEIGHT_PT))
    c = _Comment(text, author)
    c.width = width
    c.height = height
    return c


def write_section_header(ws, row, title, merge_end):
    """Write a merged section header row with SECTION_FONT/SECTION_FILL."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_end)
    style_cell(ws, row, 1, title, font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER)
    for c in range(2, merge_end + 1):
        style_cell(ws, row, c, fill=SECTION_FILL)


def apply_year_groups(ws, month_col_start, project_months, current_month=None, *, project_start=None):
    """Apply column grouping by calendar year (outline level 1).

    Groups month columns by calendar-year boundaries. If current_month is
    provided, expand the group containing that month and collapse others.
    If current_month is None, expand the first group and collapse rest.

    Args:
        ws: Worksheet to apply grouping to.
        month_col_start: First month column number (typically 3 = column C).
        project_months: Total number of project months.
        current_month: 1-based month number to keep visible (optional).
        project_start: Project start date (datetime). Required for
            calendar-year grouping. Falls back to fixed 12-month spans
            if not provided (backward compat).
    """
    if project_start is not None:
        from budget_planner.layout import _calendar_year_groups

        groups = list(_calendar_year_groups(project_start, project_months))
    else:
        # Backward compat: fixed 12-month spans
        groups = []
        m = 0
        while m < project_months:
            span = min(12, project_months - m)
            groups.append((m, span))
            m += span

    # Determine calendar year for each group
    for group_start, span in groups:
        year_start_col = month_col_start + group_start
        year_end_col = year_start_col + span - 1

        # Leave the first month of each year ungrouped so that
        # adjacent year-groups don't merge into one big group in Excel.
        group_col_start = year_start_col + 1

        # Determine calendar year for this group
        if project_start is not None:
            from dateutil.relativedelta import relativedelta

            cal_year = (project_start + relativedelta(months=group_start)).year
        else:
            cal_year = None

        # Expand 2025-2026, collapse all other years
        if cal_year is not None:
            expand = cal_year in (2025, 2026)
        elif current_month is not None:
            expand = group_start < current_month <= group_start + span
        else:
            expand = group_start == 0

        if group_col_start <= year_end_col:
            ws.column_dimensions.group(
                col_letter(group_col_start),
                col_letter(year_end_col),
                outline_level=1,
                hidden=not expand,
            )


def auto_fit_columns(ws, *, min_width=8, max_width=40):
    """Set column widths based on non-formula cell content."""
    for col_cells in ws.columns:
        max_len = 0
        for cell in col_cells:
            val = cell.value
            if val is None or (isinstance(val, str) and val.startswith("=")):
                continue
            max_len = max(max_len, len(str(val)))
        if max_len > 0:
            col_letter_str = col_cells[0].column_letter
            ws.column_dimensions[col_letter_str].width = min(max(max_len + 2, min_width), max_width)
