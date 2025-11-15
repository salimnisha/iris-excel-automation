# ======================================
# IMPORTS
# ======================================
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Styles
HEADER_FILL = PatternFill(fill_type="solid", start_color="B8CCE4")
SECTION_FILL = PatternFill(fill_type="solid", start_color="DEE6F0")
THIN_BOTTOM = Border(bottom=Side(style="thin"))
THIN_TOP = Border(top=Side(style="thin"))
THIN_ALL_SIDES = Border(
    top=Side(style="thin"),
    bottom=Side(style="thin"),
    right=Side(style="thin"),
    left=Side(style="thin"),
)
RED = "FF0000"


def autofit_colums(ws, start_col, end_col, padding=3):
    """Set column width based on maximum length of content inside column"""
    for col in range(start_col, end_col + 1):
        col_letter = get_column_letter(col)
        max_len = 0
        for cell in ws[col_letter]:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_len + padding


def write_multi_index_pivot(
    ws, pivot_df, start_row, start_col, title=None, debug=False
):
    """Writes a multi-index pivot to the worksheet
    - First index (main group) is written bold
    - Second index (items under main group) is indented under corresponding main group
    - Only one Values column

    Returns a dict of start and end row and column values
    {"start_row": 3, "end_row": 10, "start_col": 1, "end_col": 2}
    """

    # Check if dataframe is actually multi-index
    if not isinstance(pivot_df.index, pd.MultiIndex):
        raise ValueError("⚠️ Error: Pivot dataframe must have 2-level MultiIndex")

    level0_name, level1_name = pivot_df.index.names
    value_col_name = pivot_df.columns.to_list()[0]

    # Initialize address to return
    pivot_address = {}

    # ----------------------------------------------------------------
    # 1. Write the header row(s)
    # ----------------------------------------------------------------
    # Title header
    if title:
        header_cell = ws.cell(row=start_row, column=start_col, value=title)
        header_cell.font = Font(bold=True)
        header_cell.fill = HEADER_FILL
        header_cell = ws.cell(row=start_row, column=start_col + 1, value="")
        header_cell.fill = HEADER_FILL

        start_row = start_row + 2  # Leave one row below before pivot data header

    # Pivot data header
    header_cell = ws.cell(row=start_row, column=start_col, value="Row Labels")
    header_cell.font = Font(bold=True)
    header_cell.fill = SECTION_FILL
    header_cell = ws.cell(row=start_row, column=start_col + 1, value=value_col_name)
    header_cell.font = Font(bold=True)
    header_cell.fill = SECTION_FILL

    pivot_address["start_row"] = start_row
    pivot_address["start_col"] = start_col
    pivot_address["end_col"] = start_col + 1  # this is our second header_cell

    # ----------------------------------------------------------------
    # 2. Walk through the MultiIndex and write the rows
    # ----------------------------------------------------------------
    current_row = start_row + 1  # We start writing the row data from here
    current_group = None  # Track the main group, or level0 index ('Audit' or 'TA')

    # Calculate the totals for the level0 groups ('Audit' and 'TA')
    totals = pivot_df.groupby(level=0)[value_col_name].sum()
    grand_total = 0

    for (assigned_group, allocated_to), row_value in pivot_df.iterrows():
        # 🤓 df.iterrows() lets you iterate over DataFrame rows as (index, Series) pairs.
        # index: label or tuple of labels; Series: data of the row as a series. Example below
        #   Index: ('TA', 'Anika Parkar')
        #   Row: Overdue    3

        # If assigned group changes, write bold group header
        if assigned_group != current_group:
            group_cell = ws.cell(
                row=current_row, column=start_col, value=assigned_group
            )
            group_cell.font = Font(bold=True)
            group_cell.border = THIN_BOTTOM

            # Write total for the group
            total = totals[assigned_group]
            total_cell = ws.cell(row=current_row, column=start_col + 1, value=total)
            total_cell.font = Font(bold=True)
            total_cell.border = THIN_BOTTOM

            grand_total += total

            current_group = assigned_group
            current_row += 1

        # Write second-level rows of allocated_to, indented below assigned_group
        allocated_to_cell = ws.cell(
            row=current_row, column=start_col, value=allocated_to
        )
        allocated_to_cell.alignment = Alignment(indent=1)
        # Write value for the allocated_to person, retrieved using name of the value column as key
        # e.g. row_value["Overdue"] = 3  -- row_value is a variable from the for loop
        ws.cell(row=current_row, column=start_col + 1, value=row_value[value_col_name])

        current_row += 1

    # Write the grand total at the very end
    grand_cell = ws.cell(row=current_row, column=start_col, value="Grand Total")
    grand_cell.font = Font(bold=True)
    grand_cell.fill = SECTION_FILL
    grand_cell.border = THIN_TOP

    grand_cell = ws.cell(row=current_row, column=start_col + 1, value=grand_total)
    grand_cell.font = Font(bold=True)
    grand_cell.fill = SECTION_FILL
    grand_cell.border = THIN_TOP

    # Set column width
    autofit_colums(ws, start_col=start_col, end_col=start_col + 1)

    if debug:
        print(
            "\n🐞 ====== DEBUG BLOCK START: write_multi_index_pivot (report_builder.py) ======"
        )
        print("[DEBUG] Pivot title:", title)
        print("[DEBUG] Groups written:", pivot_df.index.levels[0].to_list())
        print("[DEBUG] Total rows:", current_row - start_row)
        print(
            "🐞 ====== DEBUG BLOCK END: write_multi_index_pivot (report_builder.py) ======\n"
        )

    pivot_address["end_row"] = current_row

    # return current_row
    return pivot_address


def write_pivot_tables_to_sheet(pivots, ws, debug=False):
    # -----------------------------------------------------
    # Write the pivot tables to Calculations tab
    # -----------------------------------------------------
    ws_calc = ws
    overdue_pivot = pivots["overdue"]
    due_date_pivot = pivots["due_date"]
    count_of_content_pivot = pivots["count_of_content"]
    addressed_status_pivot = pivots["addressed_status"]

    pivots_ranges = {
        "overdue": {},
        "due_date": {},
        "count_of_content": {},
        "addressed_status": {},
    }

    # Pivot1: Write Overdue pivot
    title = "Filter: 'Aged' > 0"
    address = write_multi_index_pivot(
        ws=ws_calc,
        pivot_df=overdue_pivot,
        start_row=1,
        start_col=1,
        title=title,
        debug=debug,
    )
    pivots_ranges["overdue"] = address
    print(
        "\n✅ 1. Overdue pivot written to Calculations tab up to row",
        address["end_row"],
    )

    # Pivot2: Write Due within 1-14 days pivot
    title = "Filter: 'Due Date' 1-14 days"
    address = write_multi_index_pivot(
        ws=ws_calc,
        pivot_df=due_date_pivot,
        start_row=1,
        start_col=4,
        title=title,
        debug=debug,
    )
    pivots_ranges["due_date"] = address
    print(
        "\n✅ 2. Due Date pivot written to Calculations tab up to row",
        address["end_row"],
    )

    # Pivot3: Write Count of Content
    title = "Filter: None Applied"
    address = write_multi_index_pivot(
        ws=ws_calc,
        pivot_df=count_of_content_pivot,
        start_row=1,
        start_col=8,
        title=title,
        debug=debug,
    )
    pivots_ranges["count_of_content"] = address
    print(
        "\n✅ 3. Count of Content pivot written to Calculations tab up to row",
        address["end_row"],
    )

    # Pivot4: Write Status is Addressed
    title = "Filter: 'Status' == 'Addressed'"
    address = write_multi_index_pivot(
        ws=ws_calc,
        pivot_df=addressed_status_pivot,
        start_row=1,
        start_col=11,
        title=title,
        debug=debug,
    )
    pivots_ranges["addressed_status"] = address
    print(
        "\n✅ 4. Addressed Status pivot written to Calculations tab up to row",
        address["end_row"],
    )

    return pivots_ranges


def write_table(ws, start_row, start_col, title, header, rows, row_border=False):
    """Write a simple table into an Openpyxl worksheet

    Args:
        ws (Worksheet): Worksheet object to write
        start_row (int): Top left row to place title
        start_col (int): Top left column to place title
        title (str): Title above the header
        header (list[str]): list of column names
        rows (list[list)]: Table data as list of rows
    """

    # Table address to return
    table_address = {}

    # Write the title
    cell = ws.cell(row=start_row, column=start_col, value=title)
    cell.font = Font(bold=True, color=RED)

    table_address["start_row"] = start_row
    table_address["start_col"] = start_col

    # Write the column header one row below
    header_row = start_row + 1
    for j, header in enumerate(header):
        cell = ws.cell(row=header_row, column=start_col + j, value=header)
        cell.font = Font(bold=True)
        cell.border = THIN_ALL_SIDES

    last_row = 0
    last_col = 0
    # Write table contents
    data_start_row = start_row + 2
    for i_row, row in enumerate(rows):
        for j_col, value in enumerate(row):
            cell = ws.cell(
                row=data_start_row + i_row, column=start_col + j_col, value=value
            )
            if row_border:
                cell.border = THIN_ALL_SIDES
            last_row = data_start_row + i_row
            last_col = start_col + j_col

    table_address["end_row"] = last_row
    table_address["end_col"] = last_col

    return table_address


def write_summary_tables_to_sheet(tables, ws, start_row, debug=False):
    # -----------------------------------------------------
    # Write the summary tables to Calculations tab
    # -----------------------------------------------------
    ws_calc = ws
    open_notes_table = tables["open_notes"]
    addressed_notes_table = tables["addressed_notes"]

    table_ranges = {
        "open_notes": {},
        "addressed_notes": {},
    }

    # Table1: Write Open Notes Summary table
    title = open_notes_table["title"]
    header = open_notes_table["header"]
    rows = open_notes_table["rows"]
    address = write_table(
        ws=ws_calc,
        start_row=start_row,
        start_col=1,  # Starts at first column
        title=title,
        header=header,
        rows=rows,
    )
    table_ranges["open_notes"] = address

    # Table2: Write Addressed Notes Summary table
    title = addressed_notes_table["title"]
    header = addressed_notes_table["header"]
    rows = addressed_notes_table["rows"]
    address = write_table(
        ws=ws_calc,
        start_row=start_row,
        start_col=9,  # starts at 9th column
        title=title,
        header=header,
        rows=rows,
    )
    table_ranges["addressed_notes"] = address

    return table_ranges
