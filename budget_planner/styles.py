"""Spreadsheet styling constants and helpers."""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
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


def style_cell(ws, row, col, value=None, font=None, fill=None, alignment=None, fmt=None, border=True):
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
    return cell


def col_letter(c):
    """Wrapper for get_column_letter."""
    return get_column_letter(c)
