import json

# Prepare tables to be written


def build_open_review_notes_table(base_date_str, pivot_ranges, start_row, debug=False):
    """Prepare the first summary table to be written under the pivot tables"""

    # Initialize table data to return
    table = {}

    title = f"All Open/Reopen Audit review notes to be addressed as of {base_date_str}"

    # Header row values
    header = [
        "Assigned To",
        "Overdue",
        "Due Soon",
        "Pending",
        "Grand Total",
        "As of [PREV DATE]",
        "Difference",
    ]

    # ==================================================================
    # Prepare the row values
    #   The table uses VLOOKUP formula to look up pivot tables above
    # ==================================================================

    # Pivot ranges used for VLOOKUP in the table
    p1 = pivot_ranges["overdue"]
    p2 = pivot_ranges["due_date"]
    p3 = pivot_ranges["count_of_content"]

    p1_start_row = p1["start_row"] + 1  # Add 1 because start_row is "Row Labels"
    p1_end_row = p1["end_row"]

    p2_start_row = p2["start_row"] + 1  # Add 1 because start_row is "Row Labels"
    p2_end_row = p2["end_row"]

    p3_start_row = p3["start_row"] + 1  # Add 1 because start_row is "Row Labels"
    p3_end_row = p3["end_row"]
    p3_num_rows = p3_end_row - p3_start_row + 1  # Number of rows of the table

    header_row = start_row + 1
    rows = []
    for row_val in range(p3_num_rows):
        current_row = header_row + 1 + row_val  # row_val starts at 0
        p3_row = p3_start_row + row_val  # row in pivot3

        row_content = []

        # 1. 'Assigned To' column (column A)
        # 'Assigned To' column takes values from pivot3 (count_of_content)
        # Formula '=H2'
        formula_str = f"=H{p3_row}"
        row_content.append(formula_str)

        # 2. 'Overdue' column (column B)
        # Takes VLOOKUP values from pivot1
        # =IF(OR(A50="Audit",A50="TA"), "",VLOOKUP(A50,$A$4:$B$24,2,FALSE))
        # Add IFERROR to VLOOKUP to replace #N/A with 0
        formula_str = f'=IF(OR(A{current_row}="Audit",A{current_row}="TA"), "", IFERROR(VLOOKUP(A{current_row},$A${p1_start_row}:$B${p1_end_row},2,FALSE), 0))'
        row_content.append(formula_str)

        # 3. Due Soon, pivot2
        # =IF(OR(A50="Audit",A50="TA"),"",VLOOKUP(A50,$D$4:$E$24,2,FALSE))
        formula_str = f'=IF(OR(A{current_row}="Audit",A{current_row}="TA"),"",IFERROR(VLOOKUP(A{current_row},$D${p2_start_row}:$E${p2_end_row},2,FALSE), 0))'
        row_content.append(formula_str)

        # 4. Pending, diff
        # =IF(OR(A50="Audit", A50="TA"), "",E50-C50-B50)
        formula_str = f'=IF(OR(A{current_row}="Audit", A{current_row}="TA"), "",E{current_row}-C{current_row}-B{current_row})'
        row_content.append(formula_str)

        # 5. Grand Total, pivot3
        # =IF(OR(A50="Audit",A50="TA"),"",VLOOKUP(A50,$H$2:$I$28,2,FALSE))
        formula_str = f'=IF(OR(A{current_row}="Audit",A{current_row}="TA"),"",IFERROR(VLOOKUP(A{current_row},$H${p3_start_row}:$I${p3_end_row},2,FALSE), 0))'
        row_content.append(formula_str)

        # 6. As of [Prev Date], blank
        # Leave this column blank to fill in manually
        # row_content.append("") # add a blank value

        # << ------ Temporary hardcoding ----
        # Temporarily hard-coding 'as of previous date' values, so we can generate the reports properly
        # [ ] TODO: Delete this block after deciding how to get prev date values. For now, getting the values from a temp tab called 'PrevDate' with the values
        # =VLOOKUP(A51,PrevDate!$A$1:$B$35,2,FALSE)
        formula_str = f'=IF(OR(A{current_row}="Audit",A{current_row}="TA"),"",IFERROR(VLOOKUP(A{current_row},PrevDate!$A$2:$B$36,2,FALSE), 0))'
        row_content.append(formula_str)

        # 7. Difference, diff
        # =IF(OR(A50="Audit",A50="TA"),"",E50-F50)
        formula_str = f'=IF(OR(A{current_row}="Audit",A{current_row}="TA"),"",E{current_row}-F{current_row})'
        row_content.append(formula_str)

        rows.append(row_content)

    table["title"] = title
    table["header"] = header
    table["rows"] = rows

    if debug:
        file_path = "debug/open_review_note_table.json"
        with open(file_path, "w") as f:
            json.dump(table, f, indent=2)

    return table


def build_addressed_review_notes_table(
    base_date_str, pivot_ranges, start_row, debug=False
):
    """Prepare the second summary table under the pivot tables"""

    # Initialize table data to return
    table = {}

    title = f"All Addressed review notes to be cleared as of {base_date_str}"

    # Header row values
    header = [
        "Created By",
        "Addressed",
        "As of [PREV DATE]",
        "Difference",
    ]

    # ==================================================================
    # Prepare the row values
    #   The table uses VLOOKUP formula to look up the pivot table above
    #   p4 - addressed_status pivot
    # ==================================================================

    # Pivot ranges used for VLOOKUP in the table
    p4 = pivot_ranges["addressed_status"]

    p4_start_row = p4["start_row"] + 1  # Add 1 because start_row is "Row Labels"
    p4_end_row = p4["end_row"]
    p4_num_rows = p4_end_row - p4_start_row + 1  # Number of rows of the table

    header_row = start_row + 1  # start_row is the row where table starts (with title)
    first_data_row = header_row + 1
    rows = []
    for row_val in range(p4_num_rows):
        current_row = header_row + 1 + row_val  # row_val starts at 0
        p4_row = p4_start_row + row_val  # row in pivot4

        row_content = []

        # 1. 'Created By' - column I
        # Formula '=K4'
        formula_str = f"=K{p4_row}"
        row_content.append(formula_str)

        # 2. 'Addressed' column (column J)
        # Takes VLOOKUP values from pivot4
        # =IF(OR(I50="Audit",I50="TA"), "",VLOOKUP(I50,K$4:L$24,2,FALSE)
        # Add IFERROR to VLOOKUP to replace #N/A with 0
        formula_str = f'=IF(OR(I{current_row}="Audit",I{current_row}="TA"), "", IFERROR(VLOOKUP(I{current_row},$K${p4_start_row}:$L${p4_end_row},2,FALSE), 0))'
        row_content.append(formula_str)

        # 3. As of [Prev Date], blank
        # Leave this column blank to fill in manually
        # row_content.append("")

        # << ------ Temporary hardcoding ----
        # Temporarily hard-coding 'as of previous date' values, so we can generate the reports properly
        # [ ] TODO: Delete this block after deciding how to get prev date values. For now, getting the values from a temp tab called 'PrevDate' with the values
        # =VLOOKUP(A51,PrevDate!$D:$E,2,FALSE)
        # The very last row is Total, so check for last row and amend formula
        if current_row == (first_data_row + p4_num_rows - 1):
            formula_str = f"=SUM(K{first_data_row}:K{current_row - 1})"
        else:
            formula_str = f'=IF(OR(I{current_row}="Audit",I{current_row}="TA"),"",IFERROR(VLOOKUP(I{current_row},PrevDate!$D:$E,2,FALSE), 0))'
        row_content.append(formula_str)

        # 4. Difference, diff
        # =IF(OR(I50="Audit",I50="TA"),"",J50-K50)
        formula_str = f'=IF(OR(I{current_row}="Audit",I{current_row}="TA"),"",J{current_row}-K{current_row})'
        row_content.append(formula_str)

        rows.append(row_content)

    table["title"] = title
    table["header"] = header
    table["rows"] = rows

    if debug:
        file_path = "debug/addressed_review_note_table.json"
        with open(file_path, "w") as f:
            json.dump(table, f, indent=2)

    return table


def build_signoff_aging_table(sync_time_str, pivot_ranges, start_row, debug=False):
    """Prepare the 3rd summary table under pivots"""

    # Initialize table data to return (keys: title, header, rows)
    table = {}

    title = sync_time_str

    # Header row values
    header = ["Row Labels", "Count of Workflow", "Previous Count", "Differences"]

    # ==================================================================
    # Prepare the row values
    #   The table uses formulas to look up the pivot table above.
    #   p5 - signoff_aging pivot
    #   This table is essentially the same as the pivot table, plus 2 columns
    #   to calculate the difference from prev date's values
    # ==================================================================

    # Pivot range
    p5 = pivot_ranges["signoff_aging"]

    p5_start_row = p5["start_row"] + 1  # Add 1 because start_row is "Row Labels"
    p5_end_row = p5["end_row"]
    p5_num_rows = p5_end_row - p5_start_row + 1  # number of rows of the table

    header_row = start_row + 1  # start_row is the row where table starts (with title)
    first_data_row = header_row + 1
    rows = []
    for row_val in range(p5_num_rows):
        current_row = first_data_row + row_val  # row_val starts at 0
        p5_row = p5_start_row + row_val  # row in pivot5

        row_content = []

        # 1. 'Row Labels' - Column N
        # Formula '=N4'
        formula_str = f"=N{p5_row}"
        row_content.append(formula_str)

        # 2. 'Count of Workflow' - Column O
        # Formula '=O4'
        formula_str = f"=O{p5_row}"
        row_content.append(formula_str)

        # 3. 'Previous Count', blank
        # Leave this column blank to fill in manually
        # row_content.append("")

        # << ------ Temporary hardcoding ----
        # Temporarily hard-coding 'as of previous date' values, so we can generate the reports properly
        # [ ] TODO: Delete this block after deciding how to get prev date values. For now, getting the values from a temp tab called 'PrevDate' with the values
        # =VLOOKUP(N51,PrevDate!$G:$H,2,FALSE) - reference whole column
        # The very last row is Total, so check for last row and amend formula
        if current_row == (first_data_row + p5_num_rows - 1):
            formula_str = f"=SUM(P{first_data_row}:P{current_row - 1})"
        else:
            formula_str = f"=IFERROR(VLOOKUP(N{current_row},PrevDate!$G:$H,2,FALSE), 0)"
        row_content.append(formula_str)

        # 4. 'Differences'
        # =(O51 - P51)
        formula_str = f"=O{current_row}-P{current_row}"
        row_content.append(formula_str)

        rows.append(row_content)

    table["title"] = title
    table["header"] = header
    table["rows"] = rows

    if debug:
        file_path = "debug/signoff_aging.json"
        with open(file_path, "w") as f:
            json.dump(table, f, indent=2)

    return table


def get_all_tables(base_date_str, last_sync_str, pivot_ranges, start_row, debug=False):
    # -----------------------------------------------------
    # Prepare the tables to write
    # -----------------------------------------------------
    # Table 1: Open Review Notes Table
    open_notes_table = build_open_review_notes_table(
        base_date_str, pivot_ranges, start_row, debug=debug
    )

    # Table 2: Addressed Review Notes Table
    addressed_notes_table = build_addressed_review_notes_table(
        base_date_str, pivot_ranges, start_row, debug=debug
    )

    # Table 3: Signoff Aging table
    signoff_aging_table = build_signoff_aging_table(
        last_sync_str, pivot_ranges, start_row, debug=debug
    )

    tables = {
        "open_notes": open_notes_table,
        "addressed_notes": addressed_notes_table,
        "signoff_aging": signoff_aging_table,
    }

    return tables
