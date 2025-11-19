from openpyxl.styles import PatternFill, Border, Font, Side, Alignment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter


# ========================================================
# BASIC COLORS AND FONTS
# ========================================================
RED = "FF0000"
WHITE = "FFFFFF"
DARK_BLUE = "4F81BD"
BLACK = "000000"
GREY = "D9D9D9"
SLATE_GREY = "768692"

BOLD_FONT = Font(bold=True)

# Borders
INNER_CELL_BORDER = Border(
    left=Side(style="dotted"),
    right=Side(style="dotted"),
    top=Side(style="dotted"),
    bottom=Side(style="dotted"),
)
LEFT_EDGE_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="dotted"),
    top=Side(style="dotted"),
    bottom=Side(style="dotted"),
)
TOP_EDGE_BORDER = Border(
    left=Side(style="dotted"),
    right=Side(style="dotted"),
    top=Side(style="thin"),
    bottom=Side(style="dotted"),
)
RIGHT_EDGE_BORDER = Border(
    left=Side(style="dotted"),
    right=Side(style="thin"),
    top=Side(style="dotted"),
    bottom=Side(style="dotted"),
)
BOTTOM_EDGE_BORDER = Border(
    left=Side(style="dotted"),
    right=Side(style="dotted"),
    top=Side(style="dotted"),
    bottom=Side(style="double"),  # part of footer
)
# Corners
TOP_LEFT_CORNER_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="dotted"),
    top=Side(style="thin"),
    bottom=Side(style="dotted"),
)
TOP_RIGHT_CORNER_BORDER = Border(
    left=Side(style="dotted"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="dotted"),
)
BOTTOM_LEFT_CORNER_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="dotted"),
    top=Side(style="dotted"),
    bottom=Side(style="double"),  # part of footer
)
BOTTOM_RIGHT_CORNER_BORDER = Border(
    left=Side(style="dotted"),
    right=Side(style="thin"),
    top=Side(style="dotted"),
    bottom=Side(style="double"),  # part of footer
)

# Conditional formatting colors and fills
LIGHT_RED = "ff6c6c"
ORANGE = "ff9966"
DARK_YELLOW = "febf00"
LIGHT_YELLOW = "fdf099"
LIGHT_GREEN = "91d04f"
DARK_GREEN = "05b050"

PINK = "ffc7cd"
PALE_GREEN = "c6efcd"

# [ ] TODO: For conditional formatting, Patternfill doesn't seem to work without end_color. Check why
LIGHT_RED_FILL = PatternFill(
    fill_type="solid", start_color=LIGHT_RED, end_color=LIGHT_RED
)
ORANGE_FILL = PatternFill(fill_type="solid", start_color=ORANGE, end_color=ORANGE)
DARK_YELLOW_FILL = PatternFill(
    fill_type="solid", start_color=DARK_YELLOW, end_color=DARK_YELLOW
)
LIGHT_YELLOW_FILL = PatternFill(
    fill_type="solid", start_color=LIGHT_YELLOW, end_color=LIGHT_YELLOW
)
LIGHT_GREEN_FILL = PatternFill(
    fill_type="solid", start_color=LIGHT_GREEN, end_color=LIGHT_GREEN
)
DARK_GREEN_FILL = PatternFill(
    fill_type="solid", start_color=DARK_GREEN, end_color=DARK_GREEN
)
PINK_FILL = PatternFill(fill_type="solid", start_color=PINK, end_color=PINK)
PALE_GREEN_FILL = PatternFill(
    fill_type="solid", start_color=PALE_GREEN, end_color=PALE_GREEN
)

# Autofit Column dimensions
MAX_COL_WIDTH = 30
COLUMN_PADDING = 2

# Title, Header, and Footer styles
TITLE_FONT = Font(bold=True, color=RED)

HEADER_FONT = Font(bold=True, color=WHITE)
HEADER_FILL = PatternFill(fill_type="solid", start_color=DARK_BLUE)
HEADER_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

FOOTER_FONT = Font(bold=True, color=WHITE)
FOOTER_FILL = PatternFill(fill_type="solid", start_color=SLATE_GREY)
FOOTER_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="double"),
)

GROUP_ROW_FILL = PatternFill(fill_type="solid", start_color=GREY)


# ========================================================
# BASIC FORMATTING
# ========================================================

# [ ] TODO: Function - Indent the subgroups below parent groups


def autofit_colums(ws, start_col, end_col, padding=COLUMN_PADDING, limit_width=True):
    """Set column width based on maximum length of content inside column"""
    for col in range(start_col, end_col + 1):
        col_letter = get_column_letter(col)
        max_len = 0
        for cell in ws[col_letter]:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
                if limit_width:
                    # max width should be <= MAX_COL_WIDTH
                    max_len = min(MAX_COL_WIDTH, max_len)
        ws.column_dimensions[col_letter].width = max_len + padding


def format_title(ws, row, col):
    """Style a header cell with bold font and title color"""
    cell = ws.cell(row=row, column=col)
    cell.font = TITLE_FONT


def format_header(ws, start_row, start_col, end_col):
    """Style a header with bold font, fill color"""
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=start_row, column=start_col + col - 1)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = HEADER_BORDER


def format_footer(ws, end_row, start_col, end_col):
    """Style the last row with bold font and fill color"""
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=end_row, column=start_col + col - 1)
        cell.font = FOOTER_FONT
        cell.fill = FOOTER_FILL
        cell.border = FOOTER_BORDER
        print(f"FOOTER: {end_row}, {start_col + col - 1} - {cell.value}")


def format_group_rows(
    ws, start_row, start_col, end_row, end_col, group_col, group_list
):
    """Fill the entire Group row with a grey colour.
    - group_col is the column number where the group names appear
    - group_list is a list of group names to check
    """
    for row in range(start_row, end_row + 1):
        cell = ws.cell(row=row, column=group_col)
        if cell.value in group_list:
            # Colour the entire row
            for col in range(start_col, end_col + 1):
                cell = ws.cell(row=row, column=col)
                cell.fill = GROUP_ROW_FILL
                cell.font = BOLD_FONT


def indent_child_rows(ws, start_row, end_row, group_col, group_list):
    """Indent all items appearing below the groups (from group_list)
    - group_col is the column with the groups and children
    """
    for row in range(start_row, end_row + 1):
        cell = ws.cell(row=row, column=group_col)
        if cell.value not in group_list and cell.value not in (None, ""):
            cell.alignment = Alignment(indent=2)


def format_table_data_cells(ws, start_row, start_col, end_row, end_col):
    """Style the cells inside the table
    Values passed exclude title, header, footer"""

    # --------------------------------------------
    # 1. Set all the corners
    cell = ws.cell(row=start_row, column=start_col)
    cell.border = TOP_LEFT_CORNER_BORDER
    print(f"TOP_LEFT_CORNER_BORDER: ({start_row}, {start_col}), {cell.value}")

    cell = ws.cell(row=start_row, column=end_col)
    cell.border = TOP_RIGHT_CORNER_BORDER
    print(f"TOP_RIGHT_CORNER_BORDER: ({start_row}, {end_col})), {cell.value}")

    cell = ws.cell(row=end_row, column=start_col)
    cell.border = BOTTOM_LEFT_CORNER_BORDER
    print(f"BOTTOM_LEFT_CORNER_BORDER: ({end_row}, {start_col}), {cell.value}")

    cell = ws.cell(row=end_row, column=end_col)
    cell.border = BOTTOM_RIGHT_CORNER_BORDER
    print(f"BOTTOM_RIGHT_CORNER_BORDER: ({end_row}, {end_col}), {cell.value}")

    print("=" * 60)
    # --------------------------------------------
    # 2. Set all edges (excluding corners)
    # Top edge
    for col in range(start_col + 1, end_col):
        cell = ws.cell(start_row, col)
        cell.border = TOP_EDGE_BORDER
        print(f"TOP_EDGE_BORDER: ({start_row}, {col}), {cell.value}")

    # Bottom edge
    for col in range(start_col + 1, end_col):
        cell = ws.cell(end_row, col)
        cell.border = BOTTOM_EDGE_BORDER
        print(f"BOTTOM_EDGE_BORDER: ({end_row}, {col}), {cell.value}")

    # Left edge
    for row in range(start_row + 1, end_row):
        cell = ws.cell(row, start_col)
        cell.border = LEFT_EDGE_BORDER
        print(f"LEFT_EDGE_BORDER: ({row}, {start_col}), {cell.value}")

    # Right edge
    for row in range(start_row + 1, end_row):
        cell = ws.cell(row, end_col)
        cell.border = RIGHT_EDGE_BORDER
        print(f"RIGHT_EDGE_BORDER: ({row}, {end_col}), {cell.value}")

    print("=" * 60)
    # --------------------------------------------
    # 3. Inner cells
    for row in range(start_row + 1, end_row):
        for col in range(start_col + 1, end_col):
            cell = ws.cell(row, col)
            cell.border = INNER_CELL_BORDER
            print(f"INNER_CELL_BORDER: ({row}, {col}), {cell.value}")
    print("=" * 60)


def conditional_format_number_5color_scale(ws, cell_range, rule_list=None):
    """Conditional format a number range"""
    # rules = [CellIsRule(operator="rule[0]['operator']")]
    rules = [
        CellIsRule(operator="equal", formula=['""'], stopIfTrue=True),
        CellIsRule(operator="greaterThanOrEqual", formula=["30"], fill=LIGHT_RED_FILL),
        CellIsRule(operator="between", formula=["20", "29"], fill=ORANGE_FILL),
        CellIsRule(operator="between", formula=["10", "19"], fill=DARK_YELLOW_FILL),
        CellIsRule(operator="between", formula=["5", "9"], fill=LIGHT_YELLOW_FILL),
        CellIsRule(operator="between", formula=["2", "4"], fill=LIGHT_GREEN_FILL),
        CellIsRule(operator="between", formula=["0", "1"], fill=DARK_GREEN_FILL),
    ]

    for rule in rules:
        ws.conditional_formatting.add(cell_range, rule)


def conditional_format_number_positive_negative(ws, cell_range):
    """Conditional format for >0 and <0"""
    rules = [
        CellIsRule(operator="equal", formula=['""'], stopIfTrue=True),
        CellIsRule(operator="greaterThan", formula=["0"], fill=PINK_FILL),
        CellIsRule(operator="lessThan", formula=["0"], fill=PALE_GREEN_FILL),
    ]

    for rule in rules:
        ws.conditional_formatting.add(cell_range, rule)
