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
from src.writers import write_pivot_tables_to_sheet, write_summary_tables_to_sheet
from src.tables import get_all_tables
import src.constants as C

# ======================================
# MAIN WORKFLOW
# ======================================
DEBUG = False


def main():
    # ===================================================================
    # PROCESS EXCEL
    #   Make a copy of the source file to do all further processing
    working_copy_file = make_copy(C.SOURCE_FILE, debug=DEBUG)
    force_excel_recalc(working_copy_file)  # recalculate all formulas

    #   Open the workbook. To be closed after writing all pivots, tables, and reports
    wb = load_workbook(working_copy_file)

    #   Extract base date for reports and for filtering due date pivot
    base_date = extract_base_date(ws=wb[C.BASE_DATE_SHEET], cell=C.BASE_DATE_CELL)

    # ===================================================================
    # PIVOT TABLES
    #   Read ReviewNoteAging tab and load dataframe to build pivots
    df = read_excel_dataframe(
        file_name=working_copy_file,
        sheet_name=C.DF1_SHEET,
        header_start=C.DF1_SHEET_HEADER,
        debug=DEBUG,
    )
    df.columns = df.columns.str.strip()

    #   Build and write pivots to sheet
    pivots = get_all_pivot_tables(df, base_date, debug=DEBUG)
    pivot_ranges = write_pivot_tables_to_sheet(pivots, wb[C.CALC_SHEET], debug=DEBUG)

    # ===================================================================
    # SUMMARY TABLES
    #   Prepare arguments to build and write summary tables under pivots
    max_val = max(pivot["end_row"] for pivot in pivot_ranges.values())
    table_start_row = max_val + C.BUFFER_LINES  # buffer rows after pivots
    base_date_str = base_date.strftime("%m/%d/%Y")

    #   Build and write summary tables to sheet
    tables = get_all_tables(base_date_str, pivot_ranges, table_start_row, debug=DEBUG)
    table_ranges = write_summary_tables_to_sheet(
        tables, wb[C.CALC_SHEET], table_start_row, debug=DEBUG
    )

    wb.save(working_copy_file)
    wb.close()


# ======================================
# SCRIPT ENTRY POINT
# ======================================
if __name__ == "__main__":
    main()
