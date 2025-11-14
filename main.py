# ======================================
# IMPORTS
# ======================================
from openpyxl import load_workbook

# Imports from internal project modules
from src.excel_io import (
    force_excel_recalc,
    make_copy,
    extract_base_date,
    read_excel_dataframe,
)
from src.pivots import get_all_pivot_tables
from src.writers import write_pivot_tables_to_sheet

# ======================================
# CONSTANTS & CONFIG
# ======================================
SOURCE_FILE = r"Ongoing Deliverable_US-1-US AU-1896858.1_Synopsys Inc._GDC EMSS PM Support_10.29.2025.xlsx"
REVIEW_NOTE_AGING_HEADER = 6  # Data starts in excel row 7 => header 7

# ======================================
# MAIN WORKFLOW
# ======================================
DEBUG = False


def main():
    # === Make a copy of the source file to do all further processing ===
    working_copy_file = make_copy(SOURCE_FILE, debug=DEBUG)
    force_excel_recalc(working_copy_file)  # recalculate all formulas

    # Open the workbook. Close it after writing all pivot tables
    wb = load_workbook(working_copy_file)

    # Extract base date for reports and for filtering due date pivot
    base_date = extract_base_date(ws=wb["ReviewNoteAging"], cell="B4")

    # === Read ReviewNoteAging tab and load dataframe ===
    df = read_excel_dataframe(
        file_name=working_copy_file,
        sheet_name="ReviewNoteAging",
        header_start=REVIEW_NOTE_AGING_HEADER,
        debug=DEBUG,
    )
    df.columns = df.columns.str.strip()

    pivots = get_all_pivot_tables(df, base_date, debug=DEBUG)
    last_row_written = write_pivot_tables_to_sheet(
        pivots, wb["Calculations"], debug=DEBUG
    )

    print("----- Very last row of pivot tables:", last_row_written)

    wb.save(working_copy_file)
    wb.close()


# ======================================
# SCRIPT ENTRY POINT
# ======================================
if __name__ == "__main__":
    main()
