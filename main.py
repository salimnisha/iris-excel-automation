# ======================================
# IMPORTS
# ======================================
from openpyxl import load_workbook

# Imports from internal project modules
from src.excel_io import (
    load_formula_workbook,
    force_excel_recalc,
    make_copy,
    extract_base_date,
    read_excel_dataframe,
)
from src.pivots import get_all_pivot_tables
from src.writers import (
    write_pivot_tables_to_sheet,
    write_summary_tables_to_sheet,
    copy_table_to_report,
)
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

    #   Open the workbook to for calculations and writing pivots. To be closed after writing all pivots, tables, and reports
    wb_main = load_formula_workbook(working_copy_file)

    #   Extract base date for reports and for filtering due date pivot
    base_date = extract_base_date(ws=wb_main[C.BASE_DATE_SHEET], cell=C.BASE_DATE_CELL)

    # ===================================================================
    # PIVOT TABLES
    #   Read ReviewNoteAging and Signoff Aging tabs, and load dataframe to build pivots
    df_reviewnote_aging = read_excel_dataframe(
        file_name=working_copy_file,
        sheet_name=C.DF1_SHEET,
        header_start=C.DF1_SHEET_HEADER,
        debug=DEBUG,
    )
    df_reviewnote_aging.columns = df_reviewnote_aging.columns.str.strip()

    df_signoff_aging = read_excel_dataframe(
        file_name=working_copy_file,
        sheet_name=C.DF2_SHEET,
        header_start=C.DF2_SHEET_HEADER,
        debug=DEBUG,
    )
    df_signoff_aging.columns = df_signoff_aging.columns.str.strip()

    dfs = {"reviewnote_aging": df_reviewnote_aging, "signoff_aging": df_signoff_aging}

    #   Build and write pivots to sheet
    pivots = get_all_pivot_tables(dfs, base_date, debug=DEBUG)
    pivot_ranges = write_pivot_tables_to_sheet(
        pivots, wb_main[C.CALC_SHEET], debug=DEBUG
    )

    # ===================================================================
    # SUMMARY TABLES
    #   Prepare arguments to build and write summary tables under pivots
    max_val = max(pivot["end_row"] for pivot in pivot_ranges.values())
    table_start_row = max_val + C.BUFFER_LINES  # buffer rows after pivots
    base_date_str = base_date.strftime("%m/%d/%Y")

    #   Build and write summary tables to sheet
    tables = get_all_tables(base_date_str, pivot_ranges, table_start_row, debug=DEBUG)
    table_ranges = write_summary_tables_to_sheet(
        tables, wb_main[C.CALC_SHEET], table_start_row, debug=DEBUG
    )

    # Save the file before generating the reports
    wb_main.save(working_copy_file)

    # ===================================================================
    # GENERATE FORMATTED REPORTS
    #   Prepare reports in 'Report' sheet and format them

    r1_range = copy_table_to_report(
        src_path=working_copy_file,
        wb_src=wb_main,
        table_range=table_ranges["open_notes"],
        report_start_row=1,
        report_start_col=1,
    )

    wb_main.save(working_copy_file)
    wb_main.close()

    print("Range:", r1_range)
    print("Test value", wb_main["Report"]["A4"].value)


# ======================================
# SCRIPT ENTRY POINT
# ======================================
if __name__ == "__main__":
    main()
