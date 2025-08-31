import streamlit as st
import pandas as pd
from rapidfuzz import fuzz, process

st.set_page_config(page_title="WIR-ITP Tracker", layout="wide")

st.title("📊 WIR ↔ ITP Tracker")

# ---------------------------
# File Upload
# ---------------------------
st.sidebar.header("📂 Upload Excel Files")
itp_activities_file = st.sidebar.file_uploader("Upload ITP Activities Log", type=["xlsx"])
itp_log_file = st.sidebar.file_uploader("Upload ITP Log", type=["xlsx"])
wir_file = st.sidebar.file_uploader("Upload WIR Log", type=["xlsx"])

if itp_activities_file and itp_log_file and wir_file:
    st.info("📖 Reading Excel files...")

    itp_df = pd.read_excel(itp_activities_file)
    itp_log_df = pd.read_excel(itp_log_file)
    wir_df = pd.read_excel(wir_file)

    # ---------------------------
    # Expand WIR activities
    # ---------------------------
    st.write("🔄 Expanding multi-activity WIR rows...")
    wir_rows = []
    total_rows = len(wir_df)
    progress = st.progress(0, text="Starting WIR expansion...")

    for idx, (_, row) in enumerate(wir_df.iterrows(), 1):
        activities = str(row.get("Title / Description", "")).split(",")
        for act in activities:
            new_row = row.copy()
            new_row["Activity"] = act.strip()
            wir_rows.append(new_row)

        if idx % 100 == 0 or idx == total_rows:
            progress.progress(idx / total_rows, text=f"Processed {idx}/{total_rows} WIR rows")

    wir_expanded = pd.DataFrame(wir_rows)
    st.success(f"✅ Expanded WIR activities: {len(wir_expanded)} rows total")

    # ---------------------------
    # Matching WIRs to ITP
    # ---------------------------
    st.write("🔍 Matching WIR activities to ITP activities...")
    results = []
    total_activities = len(itp_df)
    match_progress = st.progress(0, text="Starting matching...")

    itp_activities = itp_df["Activiy Description"].astype(str).tolist()

    for idx, (_, wir_row) in enumerate(wir_expanded.iterrows(), 1):
        wir_activity = str(wir_row.get("Activity", ""))
        best_matches = process.extract(
            wir_activity,
            itp_activities,
            scorer=fuzz.token_sort_ratio,
            limit=3  # get top 3 matches
        )

        for match_text, score, match_idx in best_matches:
            result_row = wir_row.copy()
            result_row["Matched ITP Activity"] = match_text
            result_row["Score"] = score
            results.append(result_row)

        if idx % 500 == 0 or idx == len(wir_expanded):
            match_progress.progress(idx / len(wir_expanded), text=f"Processed {idx}/{len(wir_expanded)} WIR activities")

    final_df = pd.DataFrame(results)

    st.success("✅ Matching completed!")

    # ---------------------------
    # Download
    # ---------------------------
    st.write("📥 Download Results")
    @st.cache_data
    def convert_df(df):
        return df.to_excel(index=False, engine="openpyxl")

    st.download_button(
        label="💾 Download Excel File",
        data=convert_df(final_df),
        file_name="wir_itp_matches.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("⬅️ Please upload all three Excel files to start.")
