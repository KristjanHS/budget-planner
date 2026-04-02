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

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)
CENTER = Alignment(horizontal="center", vertical="center")
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
    color = getattr(font, 'color', None)
    if color is None:
        return False
    rgb = getattr(color, 'rgb', None)
    if rgb is None:
        return False
    return str(rgb).upper().endswith('0000FF')


def col_letter(c):
    """Wrapper for get_column_letter."""
    return get_column_letter(c)


def apply_year_groups(ws, month_col_start, project_months, current_month=None):
    """Apply column grouping by year (outline level 1).

    Groups 12 month columns per year. If current_month is provided,
    expand the group containing that month and collapse others.
    If current_month is None, expand year 1 and collapse rest.

    Args:
        ws: Worksheet to apply grouping to.
        month_col_start: First month column number (typically 3 = column C).
        project_months: Total number of project months.
        current_month: 1-based month number to keep visible (optional).
    """
    m = 0
    while m < project_months:
        year_start_col = month_col_start + m
        span = min(12, project_months - m)
        year_end_col = year_start_col + span - 1
        year_num = m // 12 + 1

        # Determine if this year's group should be collapsed
        if current_month is not None:
            # Expand the year containing current_month
            expand = (m < current_month <= m + span)
        else:
            # Default: expand year 1
            expand = (year_num == 1)

        ws.column_dimensions.group(
            col_letter(year_start_col),
            col_letter(year_end_col),
            outline_level=1,
            hidden=not expand,
        )

        m += span
