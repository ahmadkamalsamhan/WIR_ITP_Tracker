# ------------------------------
# Ultra-High Performance WIR → ITP Tracker
# ------------------------------

import streamlit as st
import pandas as pd
from rapidfuzz import fuzz, cdist
import re
import io
import numpy as np

st.set_page_config(page_title="Ultra High-Performance WIR → ITP Tracker", layout="wide")
st.title("WIR → ITP Activity Tracking Tool (Ultra High Performance)")

# ------------------------------
# Upload Excel Files
# ------------------------------
st.header("Upload Excel Files")
wir_file = st.file_uploader("Upload WIR Log", type=["xlsx"])
itp_file = st.file_uploader("Upload ITP Log", type=["xlsx"])
activity_file = st.file_uploader("Upload ITP Activities Log", type=["xlsx"])

threshold = st.slider("Fuzzy Match Threshold (%)", 70, 100, 90)

# ------------------------------
# Start Processing Button
# ------------------------------
if wir_file and itp_file and activity_file:
    if st.button("Start Processing"):

        st.info("Reading Excel files...")
        wir_df = pd.read_excel(wir_file)
        itp_df = pd.read_excel(itp_file)
        act_df = pd.read_excel(activity_file)

        # ------------------------------
        # Clean Columns
        # ------------------------------
        def clean_columns(df):
            df.columns = df.columns.str.strip()
            df.columns = df.columns.str.replace('\n',' ').str.replace('\r',' ')
            return df

        wir_df = clean_columns(wir_df)
        itp_df = clean_columns(itp_df)
        act_df = clean_columns(act_df)

        # ------------------------------
        # Detect Columns
        # ------------------------------
        wir_pm_col = [c for c in wir_df.columns if 'PM Web Code' in c][0]
        wir_title_col = [c for c in wir_df.columns if 'Title / Description2' in c][0]

        act_itp_col = [c for c in act_df.columns if 'ITP Reference' in c][0]
        act_desc_col = [c for c in act_df.columns if 'Activiy Description' in c][0]

        # ------------------------------
        # Normalize WIR
        # ------------------------------
        wir_df['PM Web Code'] = wir_df[wir_pm_col]
        wir_df['PM_Code_Num'] = wir_df['PM Web Code'].map(lambda x: 1 if str(x).upper() in ['A','B'] else 2 if str(x).upper() in ['C','D'] else 0)

        st.info("Expanding multi-activity WIR titles (vectorized)...")
        wir_df['ActivitiesList'] = wir_df[wir_title_col].astype(str).str.split(r',|\+|/| and |&')
        wir_exp_df = wir_df.explode('ActivitiesList').reset_index(drop=True)
        wir_exp_df['ActivityNorm'] = wir_exp_df['ActivitiesList'].astype(str).str.upper().str.strip().str.replace("-", "").str.replace(" ", "")
        st.success(f"Expanded WIR activities: {len(wir_exp_df)} rows total")

        # ------------------------------
        # Normalize ITP Activities
        # ------------------------------
        act_df['ITP_Ref_Norm'] = act_df[act_itp_col].astype(str).str.upper().str.strip().str.replace("-", "").str.replace(" ", "")
        act_df['ActivityDescNorm'] = act_df[act_desc_col].astype(str).str.upper().str.strip().str.replace("-", "").str.replace(" ", "")

        # ------------------------------
        # Matching using RapidFuzz cdist (vectorized)
        # ------------------------------
        st.info("Matching WIRs to ITP activities (vectorized)...")
        progress_bar = st.progress(0)
        status_text = st.empty()

        wir_list = wir_exp_df['ActivityNorm'].tolist()
        wir_pm_list = wir_exp_df['PM_Code_Num'].tolist()
        itp_activities = act_df['ActivityDescNorm'].tolist()

        # Compute similarity matrix
        sim_matrix = cdist(itp_activities, wir_list, scorer=fuzz.ratio)
        best_idx = np.argmax(sim_matrix, axis=1)
        best_score = np.max(sim_matrix, axis=1)
        pm_codes = [wir_pm_list[i] if s >= threshold else 0 for i,s in zip(best_idx, best_score)]

        # Build matches
        match_df = pd.DataFrame({
            'ITP Reference': act_df[act_itp_col],
            'Activity Description': act_df[act_desc_col],
            'Status': pm_codes
        })

        # Build audit for low-confidence matches
        audit_df = pd.DataFrame({
            'ITP Reference': act_df[act_itp_col],
            'Activity Description': act_df[act_desc_col],
            'Best Match': [wir_exp_df['ActivityNorm'][i] for i in best_idx],
            'Score': best_score
        })
        audit_df = audit_df[audit_df['Score'] < threshold]

        progress_bar.progress(1.0)
        status_text.text("Matching completed!")

        # ------------------------------
        # Pivot Table
        # ------------------------------
        pivot_df = match_df.pivot_table(index='ITP Reference',
                                        columns='Activity Description',
                                        values='Status',
                                        fill_value=0).reset_index()

        st.subheader("Activity Completion Status Pivot Table")
        st.dataframe(pivot_df)

        # ------------------------------
        # Excel Output with Audit Sheet
        # ------------------------------
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pivot_df.to_excel(writer, index=False, sheet_name='PivotedStatus')
            if not audit_df.empty:
                audit_df.to_excel(writer, index=False, sheet_name='Audit_LowConfidence')
        output.seek(0)

        st.download_button(
            label="Download Excel",
            data=output.getvalue(),
            file_name="ITP_WIR_Activity_Status.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("Processing completed successfully!")
