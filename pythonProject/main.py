#!/usr/bin/env python3
import sys
import pandas as pd

def load_schedule(path: str) -> pd.DataFrame:
    """
    Reads the Excel or CSV and returns a DataFrame with columns:
      ['Date', 'Time', 'Course Code', 'Course Name']
    """
    # 1) Read file
    if path.lower().endswith(('.xls', '.xlsx')):
        raw = pd.read_excel(path, header=None)
    elif path.lower().endswith('.csv'):
        raw = pd.read_csv(path, header=None)
    else:
        raise ValueError("Unsupported file type. Please provide .xlsx or .csv")

    # 2) Find header row (where “Day” appears)
    header_row = raw[raw.apply(lambda row: row.astype(str).str.contains('Day', case=False).any(), axis=1)].index[0]
    data = raw.iloc[header_row+1:].reset_index(drop=True)

    # 3) Assign column names based on what we know:
    data = data.rename(columns={
        0: 'Day',
        1: 'Date',
        2: 'Morning Code',
        3: 'Morning Course',
        4: 'Afternoon Code',
        5: 'Afternoon Course'
    })

    # 4) Melt into long form
    morning = data[['Date', 'Morning Code', 'Morning Course']].dropna(subset=['Morning Code'])
    morning = morning.rename(columns={
        'Morning Code': 'Course Code',
        'Morning Course': 'Course Name'
    })
    morning['Time'] = '09:00 – 12:00'

    afternoon = data[['Date', 'Afternoon Code', 'Afternoon Course']].dropna(subset=['Afternoon Code'])
    afternoon = afternoon.rename(columns={
        'Afternoon Code': 'Course Code',
        'Afternoon Course': 'Course Name'
    })
    afternoon['Time'] = '13:00 – 16:00'

    long = pd.concat([morning, afternoon], ignore_index=True)

    # 5) Normalize course codes (upper-case, trimmed)
    long['Course Code'] = long['Course Code'].astype(str).str.strip().str.upper()

    # 6) Parse dates for sorting
    long['Date'] = pd.to_datetime(long['Date'], dayfirst=True)

    return long[['Date', 'Time', 'Course Code', 'Course Name']].sort_values(['Date', 'Time'])


def ask_int(prompt: str, min_val: int = 1) -> int:
    """Prompt until the user enters a valid integer ≥ min_val."""
    while True:
        try:
            n = int(input(prompt).strip())
            if n < min_val:
                raise ValueError()
            return n
        except ValueError:
            print(f" ➤ Please enter an integer ≥ {min_val}.")


def ask_codes(n: int, valid_codes: set) -> list:
    """Prompt the user n times for a course code, validating against valid_codes."""
    chosen = []
    for i in range(1, n+1):
        while True:
            code = input(f"Enter course code #{i}: ").strip().upper()
            if code in valid_codes:
                chosen.append(code)
                break
            else:
                print(f" ➤ “{code}” not found. Try again.")
    return chosen


def main():
    if len(sys.argv) != 2:
        print("Usage: python exam_lookup.py <path_to_schedule.xlsx|.csv>")
        sys.exit(1)

    path = sys.argv[1]
    try:
        schedule = load_schedule(path)
    except Exception as e:
        print("Error loading schedule:", e)
        sys.exit(1)

    print("\n🏫  Fast Lahore Final‐Exam Lookup  –  Spring 2025\n")
    total = ask_int("How many courses would you like to look up? ")
    codes = ask_codes(total, set(schedule['Course Code']))

    # Filter and display
    result = schedule[schedule['Course Code'].isin(codes)]
    if result.empty:
        print("\nNo matching exams found.")
    else:
        # Format date back to d-M-YYYY
        result['Date'] = result['Date'].dt.strftime('%d-%b-%Y')
        print("\nHere’s your personalized exam schedule:\n")
        print(result[['Date', 'Time', 'Course Name']].to_string(index=False))

if __name__ == '__main__':
    main()
