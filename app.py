# ------------------------------
# WIR → ITP Activity Tracking App
# ------------------------------

import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
import re
import io

st.set_page_config(page_title="WIR → ITP Tracker", layout="wide")
st.title("WIR → ITP Activity Tracking Tool (Expert Level)")

# ------------------------------
# Upload Files in Main Page
# ------------------------------
st.header("Upload Excel Files")
wir_file = st.file_uploader("Upload WIR Log", type=["xlsx"])
itp_file = st.file_uploader("Upload ITP Log", type=["xlsx"])
activity_file = st.file_uploader("Upload ITP Activities Log", type=["xlsx"])

threshold = st.slider("Fuzzy Match Threshold (%)", 70, 100, 90)

# ------------------------------
# Process Button
# ------------------------------
if wir_file and itp_file and activity_file:
    if st.button("Start Processing"):

        st.info("Reading Excel files...")
        # ------------------------------
        # Read Excel Files
        # ------------------------------
        wir_df = pd.read_excel(wir_file)
        itp_df = pd.read_excel(itp_file)
        act_df = pd.read_excel(activity_file)

        # ------------------------------
        # Normalize Columns (strip spaces/newlines)
        # ------------------------------
        def clean_columns(df):
            df.columns = df.columns.str.strip()
            df.columns = df.columns.str.replace('\n',' ').str.replace('\r',' ')
            return df

        wir_df = clean_columns(wir_df)
        itp_df = clean_columns(itp_df)
        act_df = clean_columns(act_df)

        # ------------------------------
        # Detect Required Columns
        # ------------------------------

        # WIR Log
        wir_pm_col = [c for c in wir_df.columns if 'PM Web Code' in c][0]
        wir_title_col = [c for c in wir_df.columns if 'Title / Description2' in c][0]

        wir_df['PM Web Code'] = wir_df[wir_pm_col]
        wir_df['TitleNorm'] = wir_df[wir_title_col].astype(str).str.upper().str.strip().str.replace("-", "").str.replace(" ", "")
        wir_df['PM_Code_Num'] = wir_df['PM Web Code'].map(lambda x: 1 if str(x).upper() in ['A','B'] else 2 if str(x).upper() in ['C','D'] else 0)

        # ITP Activities Log
        act_itp_col = [c for c in act_df.columns if 'ITP Reference' in c][0]
        act_desc_col = [c for c in act_df.columns if 'Activiy Description' in c][0]

        act_df['ITP_Ref_Norm'] = act_df[act_itp_col].astype(str).str.upper().str.strip().str.replace("-", "").str.replace(" ", "")
        act_df['ActivityDescNorm'] = act_df[act_desc_col].astype(str).str.upper().str.strip().str.replace("-", "").str.replace(" ", "")

        # ------------------------------
        # Split Multi-Activity WIR Titles
        # ------------------------------
        def split_activities(title):
            parts = re.split(r',|\+|/| and |&', title)
            return [p.strip() for p in parts if p.strip() != '']

        st.info("Expanding multi-activity WIR titles...")
        wir_expanded = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, row in wir_df.iterrows():
            activities = split_activities(row[wir_title_col])
            for act in activities:
                new_row = row.copy()
                new_row['SingleActivity'] = act
                new_row['ActivityNorm'] = str(act).upper().strip().replace("-", "").replace(" ", "")
                wir_expanded.append(new_row)
            progress_bar.progress((idx+1)/len(wir_df))
            status_text.text(f"Splitting WIR activities: {idx+1}/{len(wir_df)} rows processed")

        wir_exp_df = pd.DataFrame(wir_expanded)

        # ------------------------------
        # Match WIR → ITP Activities
        # ------------------------------
        st.info("Matching WIRs to ITP activities...")
        matches = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, act_row in act_df.iterrows():
            activity_desc = act_row['ActivityDescNorm']
            candidate_wirs = wir_exp_df
            matched_pm_code = 0

            for _, wir_row in candidate_wirs.iterrows():
                score = fuzz.ratio(activity_desc, wir_row['ActivityNorm'])
                if score >= threshold:
                    if wir_row['PM_Code_Num'] > matched_pm_code:
                        matched_pm_code = wir_row['PM_Code_Num']

            matches.append({
                'ITP Reference': act_row[act_itp_col],
                'Activity Description': act_row[act_desc_col],
                'Status': matched_pm_code
            })

            progress_bar.progress((idx+1)/len(act_df))
            status_text.text(f"Matching WIRs to activities: {idx+1}/{len(act_df)} activities processed")

        match_df = pd.DataFrame(matches)

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
        # Excel Output
        # ------------------------------
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pivot_df.to_excel(writer, index=False, sheet_name='PivotedStatus')
        output.seek(0)

        st.download_button(
            label="Download Excel",
            data=output.getvalue(),
            file_name="ITP_WIR_Activity_Status.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("Processing completed!")
