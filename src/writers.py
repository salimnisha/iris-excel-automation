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
    """

    # Check if dataframe is actually multi-index
    if not isinstance(pivot_df.index, pd.MultiIndex):
        raise ValueError("⚠️ Error: Pivot dataframe must have 2-level MultiIndex")

    level0_name, level1_name = pivot_df.index.names
    value_col_name = pivot_df.columns.to_list()[0]

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

    return current_row


def write_pivot_tables_to_sheet(pivots, ws, debug=False):
    # -----------------------------------------------------
    # Write the pivot tables to Calculations tab
    # -----------------------------------------------------
    ws_calc = ws
    overdue_pivot = pivots["overdue"]
    due_date_pivot = pivots["due_date"]
    count_of_content_pivot = pivots["count_of_content"]
    addressed_status_pivot = pivots["addressed_status"]

    final_written_row = 0

    # Pivot1: Write Overdue pivot
    title = "Filter: 'Aged' > 0"
    last_row = write_multi_index_pivot(
        ws=ws_calc,
        pivot_df=overdue_pivot,
        start_row=1,
        start_col=1,
        title=title,
        debug=debug,
    )
    final_written_row = max(final_written_row, last_row)
    print("\n✅ 1. Overdue pivot written to Calculations tab up to row", last_row)

    # Pivot2: Write Due within 1-14 days pivot
    title = "Filter: 'Due Date' between 0-14 days (includes starting date)"
    last_row = write_multi_index_pivot(
        ws=ws_calc,
        pivot_df=due_date_pivot,
        start_row=1,
        start_col=4,
        title=title,
        debug=debug,
    )
    final_written_row = max(final_written_row, last_row)
    print("\n✅ 2. Due Date pivot written to Calculations tab up to row", last_row)

    # Pivot3: Write Count of Content
    title = "Filter: None Applied"
    last_row = write_multi_index_pivot(
        ws=ws_calc,
        pivot_df=count_of_content_pivot,
        start_row=1,
        start_col=8,
        title=title,
        debug=debug,
    )
    final_written_row = max(final_written_row, last_row)
    print(
        "\n✅ 3. Count of Content pivot written to Calculations tab up to row", last_row
    )

    # Pivot4: Write Status is Addressed
    title = "Filter: 'Status' == 'Addressed'"
    last_row = write_multi_index_pivot(
        ws=ws_calc,
        pivot_df=addressed_status_pivot,
        start_row=1,
        start_col=11,
        title=title,
        debug=debug,
    )
    final_written_row = max(final_written_row, last_row)
    print(
        "\n✅ 4. Addressed Status pivot written to Calculations tab up to row", last_row
    )

    return final_written_row
