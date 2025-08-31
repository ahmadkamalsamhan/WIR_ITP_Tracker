# ------------------------------
# WIR → ITP Activity Tracking App
# ------------------------------

import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
import re
import io
from tqdm import tqdm

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

        # ------------------------------
        # Read Excel Files
        # ------------------------------
        wir_df = pd.read_excel(wir_file)
        itp_df = pd.read_excel(itp_file)
        act_df = pd.read_excel(activity_file)

        # ------------------------------
        # Normalize Column Names
        # ------------------------------
        def clean_columns(df):
            df.columns = df.columns.str.strip()
            df.columns = df.columns.str.replace('\n',' ').str.replace('\r',' ')
            return df

        wir_df = clean_columns(wir_df)
        itp_df = clean_columns(itp_df)
        act_df = clean_columns(act_df)

        # ------------------------------
        # Detect PM Web Code Column Robustly
        # ------------------------------
        pm_col_candidates = [col for col in wir_df.columns if 'PM' in col.upper() and 'CODE' in col.upper()]
        if len(pm_col_candidates) == 0:
            st.error("Cannot find PM Web Code column in WIR log!")
            st.stop()
        wir_df['PM Web Code'] = wir_df[pm_col_candidates[0]]

        # ------------------------------
        # Normalize Text Columns
        # ------------------------------
        def normalize_text(x):
            return str(x).upper().strip().replace("-", "").replace(" ", "")

        wir_df['TitleNorm'] = wir_df['Title / Description2'].astype(str).apply(normalize_text)
        wir_df['ProjectNorm'] = wir_df['Project Name'].astype(str).apply(normalize_text)
        wir_df['DisciplineNorm'] = wir_df['Discipline'].astype(str).apply(normalize_text)
        wir_df['PhaseNorm'] = wir_df['Phase'].astype(str).apply(normalize_text)
        wir_df['PM_Code_Num'] = wir_df['PM Web Code'].map(lambda x: 1 if str(x).upper() in ['A','B'] else 2 if str(x).upper() in ['C','D'] else 0)

        itp_df['ITP_Ref_Norm'] = itp_df['ITP Reference'].astype(str).apply(normalize_text)
        act_df['ITP_Ref_Norm'] = act_df['ITP Reference'].astype(str).apply(normalize_text)
        act_df['ActivityDescNorm'] = act_df['Activity Description'].astype(str).apply(normalize_text)

        # ------------------------------
        # Split Multi-Activity WIR Titles
        # ------------------------------
        def split_activities(title):
            parts = re.split(r',|\+|/| and |&', title.upper())
            return [p.strip() for p in parts if p.strip() != '']

        wir_expanded = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, row in wir_df.iterrows():
            activities = split_activities(row['Title / Description2'])
            for act in activities:
                new_row = row.copy()
                new_row['SingleActivity'] = act
                wir_expanded.append(new_row)
            progress_bar.progress((idx+1)/len(wir_df))
            status_text.text(f"Splitting WIR activities: {idx+1}/{len(wir_df)} rows processed")

        wir_exp_df = pd.DataFrame(wir_expanded)
        wir_exp_df['ActivityNorm'] = wir_exp_df['SingleActivity'].astype(str).apply(normalize_text)

        # ------------------------------
        # Match WIR → ITP Activities
        # ------------------------------
        matches = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, act_row in act_df.iterrows():
            itp_ref = act_row['ITP_Ref_Norm']
            activity_desc = act_row['ActivityDescNorm']

            candidate_wirs = wir_exp_df

            matched_pm_code = 0

            for widx, wir_row in candidate_wirs.iterrows():
                title_score = fuzz.ratio(activity_desc, wir_row['ActivityNorm'])
                if title_score >= threshold:
                    if wir_row['PM_Code_Num'] > matched_pm_code:
                        matched_pm_code = wir_row['PM_Code_Num']

            matches.append({
                'ITP Reference': act_row['ITP Reference'],
                'Activity Description': act_row['Activity Description'],
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
