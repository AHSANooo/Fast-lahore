#!/usr/bin/env python3
import os
import streamlit as st
import pandas as pd
from dateutil import parser

@st.cache_data(show_spinner=False)
def load_schedule():
    fname = "final date sheet-Spring2025-STU-v1.3.xlsx"
    if not os.path.exists(fname):
        st.error(f"❗ Cannot find '{fname}' in this folder.")
        st.stop()

    raw = pd.read_excel(fname, header=None, engine="openpyxl")
    HEADER = 2
    morning_time   = str(raw.iat[HEADER, 3]).strip()
    afternoon_time = str(raw.iat[HEADER, 5]).strip()

    block = raw.iloc[HEADER+1:, [1,2,3,4,5]].copy()
    block.columns = [
        "Date",
        "Morning Code", "Morning Course",
        "Afternoon Code", "Afternoon Course"
    ]
    block = block.dropna(subset=["Date"])

    def parse_date(x):
        if isinstance(x, (int, float)):
            return pd.to_datetime(x, unit='d', origin='1899-12-30')
        return parser.parse(str(x), dayfirst=True)

    block["Date"] = block["Date"].apply(parse_date)

    morning = (
        block[["Date","Morning Code","Morning Course"]]
        .dropna(subset=["Morning Code"])
        .rename(columns={"Morning Code":"Course_Code","Morning Course":"Course_Name"})
    )
    morning["Time"] = morning_time

    afternoon = (
        block[["Date","Afternoon Code","Afternoon Course"]]
        .dropna(subset=["Afternoon Code"])
        .rename(columns={"Afternoon Code":"Course_Code","Afternoon Course":"Course_Name"})
    )
    afternoon["Time"] = afternoon_time

    df = pd.concat([morning, afternoon], ignore_index=True)
    df["Course_Code"] = df["Course_Code"].astype(str).str.strip().str.lower()
    return df[["Date","Time","Course_Code","Course_Name"]].sort_values(["Date","Time"])


st.set_page_config(page_title="FAST Lahore Exam Lookup", layout="centered")
st.title("📅 FAST Lahore Final-Exam Lookup")
st.markdown("**Spring 2025** – Best of luck of the Finals!")

schedule = load_schedule()

n = st.number_input("How many courses to look up?", min_value=1, step=1, value=1)

codes = []
cols = st.columns(2)
for i in range(n):
    with cols[i % 2]:
        codes.append(st.text_input(f"Course code #{i+1}", "").strip().lower())

if st.button("Show My Schedule"):
    valid = set(schedule["Course_Code"])
    bad = [c for c in codes if c and c not in valid]
    if bad:
        st.warning(f"❗ These codes were not found: {', '.join(bad)}")
    else:
        df = schedule[schedule["Course_Code"].isin(codes)].copy()
        if df.empty:
            st.info("No exams found for the entered codes.")
        else:
            # format the date
            df["Date"] = df["Date"].dt.strftime("%d-%b-%Y")
            # reset index to start from 1, and name the index column "No."
            df_display = df[["Date","Time","Course_Name"]].reset_index(drop=True)
            df_display.index = df_display.index + 1
            df_display.index.name = "No."
            st.markdown("### 🎓 Your Exam Schedule")
            st.table(df_display)

# optional footer
st.markdown("---")
st.markdown("📧 **For any issues or support, please contact:** [i230553@isb.nu.edu.pk](mailto:i230553@isb.nu.edu.pk)")
st.markdown("🔗 **Connect with me on LinkedIn:** [Muhammad Ahsan](https://www.linkedin.com/in/muhammad-ahsan-7612701a7)")
