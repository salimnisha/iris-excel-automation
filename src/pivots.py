# ======================================
# IMPORTS
# ======================================
import pandas as pd
from datetime import datetime


def build_overdue_pivot(df: pd.DataFrame, debug: bool = False):
    """Build a pivot for Overdue counts by 'Assigned group' and 'Allocated To', filtered to rows where Aged > 0"""

    rows = ["Assigned group", "Allocated To"]
    values = "Content"  # To be renamed as Overdue

    required = ["Assigned group", "Allocated To", "Content"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Missing columns for overdue_pivot: {missing}")

    # Filter for Aged
    filtered_df = df[df["Aged"] > 0].copy()

    # Build pivot (keep the #N/A values because we want to display them as well)
    pivot = filtered_df.pivot_table(values=values, index=rows, aggfunc="count").rename(
        columns={"Content": "Overdue"}
    )

    if debug:
        print("\n🐞 ====== DEBUG BLOCK START: build_overdue_pivot (pivots.py) ======")
        # Check if the dataframe structure is correct. MultiIndex expected
        print("[DEBUG] Pivot Index Type:", type(pivot.index))  # Expected: MultiIndex
        print(
            "[DEBUG] Pivot Index Names:", pivot.index.names
        )  # Expected: ['Assigned group', 'Allocated To']
        print(
            "[DEBUG] Pivot Columns:", pivot.columns.to_list()
        )  # Expected: ['Overdue']
        print(
            "[DEBUG] Pivot Preview:", pivot.head(5)
        )  # Expected: top of multiindex pivot data

        # Write to debug files
        pivot.to_csv("debug/debug_pivot1.csv", index=True)
        pivot.to_pickle("debug/debug_pivot1.pkl")
        print("🐞 Saved intermediate pivot1 to debug_pivot1.csv and debug_pivot1.pkl")
        print("🐞 ====== DEBUG BLOCK END: build_overdue_pivot (pivots.py) ====== \n")

    return pivot


def build_due_date_pivot(df: pd.DataFrame, base_date: datetime, debug: bool = False):
    """Build a pivot for Due Date count by 'Assigned group' and 'Allocated To', filtered for 'Due within 1-14 Days' start day inclusive"""

    rows = ["Assigned group", "Allocated To"]
    values = "Content"  # To be renamed 'Due within 1-14 Days'

    required = ["Assigned group", "Allocated To", "Content"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Missing columns for due_date_pivot: {missing}")

    # Apply filter for due dates between 1-14 days (including today)
    # TODO: Verify with Iris - might need to change to 1-13 days in filter
    filter_due_in_14_days = (df["Due Date"] - base_date).dt.days.between(
        0, 14, inclusive="both"
    )
    df1 = df[filter_due_in_14_days]

    # Build pivot
    pivot1 = df1.pivot_table(values=values, index=rows, aggfunc="count").rename(
        columns={"Content": "Due within 1-14 Days"}
    )

    if debug:
        print("\n🐞 ====== DEBUG BLOCK START: build_due_date_pivot (pivots.py) ======")
        # Check dataframe structure. MultiIndex expected
        print("[DEBUG] Pivot Index type:", type(pivot1.index))
        print(
            "[DEBUG] Pivot index names:", pivot1.index.names
        )  # Expected: ['Assigned group', 'Allocated To']
        print(
            "[DEBUG] Pivot columns:", pivot1.columns.to_list()
        )  # Expected: ['Due within 1-14 Days']
        print("[DEBUG] Pivot preview:", pivot1.head(5))

        # Write to debug files
        pivot1.to_csv("debug/debug_pivot2.csv", index=True)
        pivot1.to_pickle("debug/debug_pivot2.pkl")
        print("🐞 Saved pivot2 dump into debug_pivot2.csv and debug_pivot2.pkl")
        print("🐞 ====== DEBUG BLOCK END: build_due_date_pivot (pivots.py) ====== \n")

    return pivot1


def build_count_of_content_pivot(df: pd.DataFrame, debug: bool = False):
    """Build a pivot with no filters, indexed to 'Assigned group' and 'Allocated To', count of Content value"""

    rows = ["Assigned group", "Allocated To"]
    values = "Content"

    required = ["Assigned group", "Allocated To", "Content"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Missing columns for count_of_content_pivot: {missing}")

    # Build pivot
    pivot = df.pivot_table(values=values, index=rows, aggfunc="count")

    if debug:
        print(
            "\n🐞 ====== DEBUG BLOCK START: build_count_of_content_pivot (pivots.py) ======"
        )
        # Check if the dataframe structure is correct. MultiIndex expected
        print("[DEBUG] Pivot Index Type:", type(pivot.index))  # Expected: MultiIndex
        print(
            "[DEBUG] Pivot Index Names:", pivot.index.names
        )  # Expected: ['Assigned group', 'Allocated To']
        print(
            "[DEBUG] Pivot Columns:", pivot.columns.to_list()
        )  # Expected: ['Count of Content']
        print(
            "[DEBUG] Pivot Preview:", pivot.head(5)
        )  # Expected: top of multiindex pivot data

        # Write to debug files
        pivot.to_csv("debug/debug_pivot3.csv", index=True)
        pivot.to_pickle("debug/debug_pivot3.pkl")
        print("🐞 Saved intermediate pivot1 to debug_pivot3.csv and debug_pivot3.pkl")
        print(
            "🐞 ====== DEBUG BLOCK END: build_count_of_content_pivot (pivots.py) ====== \n"
        )

    return pivot


def build_addressed_status_pivot(df: pd.DataFrame, debug: bool = False):
    rows = ["Created by group", "Created By"]
    values = "Content"

    required = ["Created by group", "Created By", "Content", "Status"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Missing columns for addressed_status_pivot: {missing}")

    # Filter for 'Status' = 'Addressed'
    filtered_df = df[df["Status"] == "Addressed"]

    # Build pivot table
    pivot = filtered_df.pivot_table(values=values, index=rows, aggfunc="count").rename(
        columns={"Content": "Addressed"}
    )

    if debug:
        print(
            "\n🐞 ====== DEBUG BLOCK START: build_addressed_status_pivot (pivots.py) ======"
        )
        # Check if the dataframe structure is correct. MultiIndex expected
        print("[DEBUG] Pivot Index Type:", type(pivot.index))  # Expected: MultiIndex
        print(
            "[DEBUG] Pivot Index Names:", pivot.index.names
        )  # Expected: ['Created by group', 'Created By']
        print(
            "[DEBUG] Pivot Columns:", pivot.columns.to_list()
        )  # Expected: ['Addressed']
        print(
            "[DEBUG] Pivot Preview:", pivot.head(5)
        )  # Expected: top of multiindex pivot data

        # Write to debug files
        pivot.to_csv("debug/debug_pivot4.csv", index=True)
        pivot.to_pickle("debug/debug_pivot4.pkl")
        print("🐞 Saved intermediate pivot to debug_pivot4.csv and debug_pivot4.pkl")
        print(
            "🐞 ====== DEBUG BLOCK END: build_addressed_status_pivot (pivots.py) ====== \n"
        )

    return pivot


def build_signoff_aging_pivot(df: pd.DataFrame, debug: bool = False):
    df_row = "Assignee"
    df_value = "Workflow"
    df_filter_col = "Signoff Role"

    required = [df_row, df_value, df_filter_col]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Missing columns for signoff_aging pivot: {missing}")

    # Filter for Sign-off Role
    filtered_df = df[
        (df["Signoff Role"] != "In-Charge") & (df["Signoff Role"] != "Senior")
    ]

    # Build pivot table
    pivot = filtered_df.pivot_table(values=df_value, index=df_row, aggfunc="count")

    if debug:
        print(
            "🐞 ====== DEBUG BLOCK START: build_signoff_aging_pivot (pivots.py) ====== \n"
        )
        print(f"[DEBUG] Pivot index names: {pivot.index.names}")
        print(f"[DEBUG] Pivot columns: {pivot.columns}")
        print(f"[DEBUG] Pivot previes: {pivot.head(5)}")

        # Write to debug file
        pivot.to_csv("debug/debug_pivot5.csv", index=True)
        pivot.to_pickle("debug/debug_pivot5.pkl")
        print("🐞 Saved intermediate pivot to debug_pivot5.csv and debug_pivot5.pkl")
        print(
            "🐞 ====== DEBUG BLOCK END: build_signoff_aging_pivot (pivots.py) ====== \n"
        )

    return pivot


def get_all_pivot_tables(dfs, base_date, debug=False):
    """Prepare the required pivot tables and return as a list"""

    # Unpack the dfs dict that contain two dataframes from 'ReviewNoteAging' and 'SignoffAging'
    df_reviewnote = dfs["reviewnote_aging"]
    df_signoff = dfs["signoff_aging"]

    # -----------------------------------------------------
    # Prepare the pivot tables
    # -----------------------------------------------------
    # Pivot1: Overdue pivot
    overdue_pivot = build_overdue_pivot(df=df_reviewnote, debug=debug)

    # Pivot2: Overdue pivot
    if base_date:
        due_date_pivot = build_due_date_pivot(
            df=df_reviewnote, base_date=base_date, debug=debug
        )
    else:
        # Empty pivot
        due_date_pivot = df.iloc[0:0].pivot_table(
            values="Content", index=["Assigned group", "Allocated To"], aggfunc="count"
        )

    # Pivot3: Count of content pivot
    count_of_content_pivot = build_count_of_content_pivot(df=df_reviewnote, debug=debug)

    # Pivot4: Addressed Status pivot
    addressed_status_pivot = build_addressed_status_pivot(df=df_reviewnote, debug=debug)

    # Pivot5: Signoff Aging pivot
    signoff_aging_pivot = build_signoff_aging_pivot(df=df_signoff, debug=debug)

    pivots = {
        "overdue": overdue_pivot,
        "due_date": due_date_pivot,
        "count_of_content": count_of_content_pivot,
        "addressed_status": addressed_status_pivot,
        "signoff_aging": signoff_aging_pivot,
    }

    return pivots
