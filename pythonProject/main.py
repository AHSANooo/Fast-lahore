#!/usr/bin/env python3
import os
import sys
import pandas as pd
from dateutil import parser

def load_schedule(path: str) -> pd.DataFrame:
    """
    Load your Spring 2025 date sheet (row 3 zero-based index 2 header),
    then melt morning/afternoon into one table with parsed dates.
    """
    raw = pd.read_excel(path, header=None, engine="openpyxl")

    HEADER = 2  # zero-based index for the row containing “Day | Date | ...”
    # Read the two time-slot labels from that header row
    morning_time   = str(raw.iat[HEADER, 3]).strip()
    afternoon_time = str(raw.iat[HEADER, 5]).strip()

    # Extract only the columns we need under that header
    block = raw.iloc[HEADER+1:, [1, 2, 3, 4, 5]].copy()
    block.columns = [
        "Date",
        "Morning Code", "Morning Course",
        "Afternoon Code", "Afternoon Course"
    ]
    block = block.dropna(subset=["Date"])

    # Use dateutil to parse any month format robustly
    def parse_date(x):
        # Convert Excel serials or strings to Python datetime
        if isinstance(x, (int, float)):
            # Excel serial date
            return pd.to_datetime(x, unit='d', origin='1899-12-30')
        else:
            return parser.parse(str(x), dayfirst=True)
    block["Date"] = block["Date"].apply(parse_date)

    # Build “long” table of exams
    morning = block[["Date", "Morning Code", "Morning Course"]].dropna(subset=["Morning Code"])
    morning = morning.rename(columns={
        "Morning Code":   "Course_Code",
        "Morning Course": "Course_Name"
    })
    morning["Time"] = morning_time

    afternoon = block[["Date", "Afternoon Code", "Afternoon Course"]].dropna(subset=["Afternoon Code"])
    afternoon = afternoon.rename(columns={
        "Afternoon Code":   "Course_Code",
        "Afternoon Course": "Course_Name"
    })
    afternoon["Time"] = afternoon_time

    df = pd.concat([morning, afternoon], ignore_index=True)
    df["Course_Code"] = df["Course_Code"].astype(str).str.strip().str.lower()

    return df[["Date", "Time", "Course_Code", "Course_Name"]].sort_values(["Date", "Time"])


def ask_int(prompt: str, min_val: int = 1) -> int:
    """Prompt until the user enters an integer ≥ min_val."""
    while True:
        try:
            n = int(input(prompt).strip())
            if n < min_val:
                raise ValueError()
            return n
        except ValueError:
            print(f" ➤ Please enter an integer ≥ {min_val}.")


def ask_codes(n: int, valid: set) -> list:
    """Prompt n times for course codes (case-insensitive)."""
    chosen = []
    for i in range(1, n+1):
        while True:
            code = input(f"Enter course code #{i}: ").strip().lower()
            if code in valid:
                chosen.append(code)
                break
            else:
                print(f" ➤ “{code}” not found. Try again.")
    return chosen


def main():
    print("\n🏫  Fast Lahore Final‐Exam Lookup  –  Spring 2025\n")

    fname = "final date sheet-Spring2025-STU-v1.3.xlsx"
    if not os.path.exists(fname):
        print(f"❗ Cannot find '{fname}' in this folder.")
        sys.exit(1)

    try:
        schedule = load_schedule(fname)
    except Exception as e:
        print("❗ Failed to load schedule:", e)
        sys.exit(1)

    total = ask_int("How many courses would you like to look up? ")
    codes = ask_codes(total, set(schedule["Course_Code"]))

    # Filter & display
    result = schedule[schedule["Course_Code"].isin(codes)].copy()
    if result.empty:
        print("\nNo matching exams found.")
    else:
        result["Date"] = result["Date"].dt.strftime("%d-%b-%Y")
        print("\nHere’s your personalized exam schedule:\n")
        print(result[["Date", "Time", "Course_Name"]].to_string(index=False))


if __name__ == "__main__":
    main()
