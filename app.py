import streamlit as st
import pandas as pd
import numpy as np
from rapidfuzz import fuzz, process

st.set_page_config(page_title="WIR & ITP Tracker", layout="wide")

st.title("📊 WIR & ITP Tracker")

# File uploader
st.sidebar.header("Upload your files (CSV/Excel)")
itp_activities_file = st.sidebar.file_uploader("Upload ITP Activities", type=["csv", "xlsx"])
itp_log_file = st.sidebar.file_uploader("Upload ITP Log", type=["csv", "xlsx"])
wir_log_file = st.sidebar.file_uploader("Upload WIR Log", type=["csv", "xlsx"])

progress = st.empty()

if st.sidebar.button("🚀 Start Processing"):
    if not (itp_activities_file and itp_log_file and wir_log_file):
        st.error("Please upload all 3 files!")
    else:
        try:
            # Read files
            def read_file(file):
                if file.name.endswith(".csv"):
                    return pd.read_csv(file)
                return pd.read_excel(file)

            st.info("Reading Excel/CSV files...")
            itp_activities = read_file(itp_activities_file)
            itp_log = read_file(itp_log_file)
            wir_log = read_file(wir_log_file)

            st.success("Files loaded successfully ✅")
            st.write("ITP Activities:", itp_activities.head())
            st.write("ITP Log:", itp_log.head())
            st.write("WIR Log:", wir_log.head())

            # Example fuzzy matching (demo)
            st.info("Matching WIRs to ITP activities...")

            matches = []
            total = len(itp_activities)
            for i, row in itp_activities.iterrows():
                act = str(row.get("Activiy Description", ""))
                best = process.extractOne(act, wir_log["Title / Description"].astype(str), scorer=fuzz.partial_ratio)
                matches.append((act, best))
                if i % 50 == 0:
                    progress.progress(int(i / total * 100))

            st.success("✅ Matching completed!")
            results = pd.DataFrame(matches, columns=["Activity", "Best Match"])
            st.dataframe(results.head(20))

        except Exception as e:
            st.error(f"❌ Error: {e}")
