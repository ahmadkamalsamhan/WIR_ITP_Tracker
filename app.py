# ------------------------------
# Ultra High-Performance WIR → ITP Tracker (CSV input)
# ------------------------------

import streamlit as st
import pandas as pd
import numpy as np
from rapidfuzz import fuzz
import io

st.set_page_config(page_title="Ultra High-Performance WIR → ITP Tracker", layout="wide")
st.title("WIR → ITP Activity Tracking Tool (Ultra High Performance)")

# ------------------------------
# Upload CSV Files
# ------------------------------
st.header("Upload CSV Files")
wir_file = st.file_uploader("Upload WIR Log (CSV)", type=["csv"])
itp_file = st.file_uploader("Upload ITP Log (CSV)", type=["csv"])
activity_file = st.file_uploader("Upload ITP Activities Log (CSV)", type=["csv"])

threshold = st.slider("Fuzzy Match Threshold (%)", 70, 100, 90)

# ------------------------------
# Start Processing Button
# ------------------------------
if wir_file and itp_file and activity_file:
    if st.button("Start Processing"):

        st.info("Reading CSV files...")
        wir_df = pd.read_csv(wir_file)
        itp_df = pd.read_csv(itp_file)
        act_df = pd.read_csv(activity_file)

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
        # Detect Columns (adjust if needed)
        # ------------------------------
        wir_pm_col = [c for c in wir_df.columns if 'PM Web Code' in c][0]
        wir_title_col = [c for c in wir_df.columns if 'Title' in c or 'Description' in c][0]

        act_itp_col = [c for c in act_df.columns if 'ITP Reference' in c][0]
        act_desc_col = [c for c in act_df.columns if 'Activiy Description' in c][0]

        # ------------------------------
        # Normalize WIR
        # ------------------------------
        wir_df['PM_Code_Num'] = wir_df[wir_pm_col].map(lambda x: 1 if str(x).upper() in ['A','B'] else 2 if str(x).upper() in ['C','D'] else 0)

        st.info("Expanding multi-activity WIR titles...")
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
        # Matching with progress bar
        # ------------------------------
        st.info("Matching WIRs to ITP activities...")
        progress_bar = st.progress(0)
        status_text = st.empty()

        wir_list = wir_exp_df['ActivityNorm'].tolist()
        wir_pm_list = wir_exp_df['PM_Code_Num'].tolist()

        match_results = []
        audit_results = []

        total = len(act_df)
        for idx, row in act_df.iterrows():
            act_name = row['ActivityDescNorm']
            best_score = 0
            best_idx = -1
            for i, wir_name in enumerate(wir_list):
                score = fuzz.ratio(act_name, wir_name)
                if score > best_score:
                    best_score = score
                    best_idx = i
            pm_code = wir_pm_list[best_idx] if best_score >= threshold else 0
            match_results.append({
                'ITP Reference': row[act_itp_col],
                'Activity Description': row[act_desc_col],
                'Status': pm_code
            })
            if best_score < threshold:
                audit_results.append({
                    'ITP Reference': row[act_itp_col],
                    'Activity Description': row[act_desc_col],
                    'Best Match': wir_list[best_idx],
                    'Score': best_score
                })
            if idx % 50 == 0:
                progress_bar.progress(idx/total)
                status_text.text(f"Processing {idx}/{total} ITP activities...")

        progress_bar.progress(1.0)
        status_text.text("Matching completed!")

        match_df = pd.DataFrame(match_results)
        audit_df = pd.DataFrame(audit_results)

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
