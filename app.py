# ------------------------------
# 1. Install required packages
# ------------------------------
# pip install streamlit pandas openpyxl rapidfuzz tqdm

import streamlit as st
import pandas as pd
from rapidfuzz import fuzz
import re
from tqdm import tqdm
import io

# ------------------------------
# 2. Streamlit App Layout
# ------------------------------
st.set_page_config(page_title="WIR → ITP Activity Tracker", layout="wide")
st.title("WIR → ITP Activity Tracking Tool (Expert Level)")

# ------------------------------
# 3. File Uploads
# ------------------------------
st.sidebar.header("Upload Excel Files")
wir_file = st.sidebar.file_uploader("Upload WIR Log", type=["xlsx"])
itp_file = st.sidebar.file_uploader("Upload ITP Log", type=["xlsx"])
activity_file = st.sidebar.file_uploader("Upload ITP Activities Log", type=["xlsx"])

threshold = st.sidebar.slider("Fuzzy Match Threshold (%)", 70, 100, 90)

if wir_file and itp_file and activity_file:

    # ------------------------------
    # 4. Read Excel Files
    # ------------------------------
    wir_df = pd.read_excel(wir_file)
    itp_df = pd.read_excel(itp_file)
    act_df = pd.read_excel(activity_file)

    # ------------------------------
    # 5. Data Normalization
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
    # 6. Split Multi-Activity WIR Titles
    # ------------------------------
    def split_activities(title):
        parts = re.split(r',|\+|/| and |&', title.upper())
        return [p.strip() for p in parts if p.strip() != '']

    wir_expanded = []
    for idx, row in tqdm(wir_df.iterrows(), total=len(wir_df), desc="Splitting WIR activities"):
        activities = split_activities(row['Title / Description2'])
        for act in activities:
            new_row = row.copy()
            new_row['SingleActivity'] = act
            wir_expanded.append(new_row)

    wir_exp_df = pd.DataFrame(wir_expanded)
    wir_exp_df['ActivityNorm'] = wir_exp_df['SingleActivity'].astype(str).apply(normalize_text)

    # ------------------------------
    # 7. Matching WIR → ITP Activities
    # ------------------------------
    matches = []
    for idx, act_row in tqdm(act_df.iterrows(), total=len(act_df), desc="Matching WIRs to Activities"):
        itp_ref = act_row['ITP_Ref_Norm']
        activity_desc = act_row['ActivityDescNorm']

        # Filter WIRs for the same ITP (Transmittal/Document matching can also be added)
        candidate_wirs = wir_exp_df

        best_match_score = 0
        matched_pm_code = 0

        for widx, wir_row in candidate_wirs.iterrows():
            title_score = fuzz.ratio(activity_desc, wir_row['ActivityNorm'])
            # Optional: add weights for project, discipline, phase
            project_match = 1 if wir_row['ProjectNorm'] == normalize_text(itp_df.loc[itp_df['ITP_Ref_Norm']==itp_ref,'Title / Description'].values[0]) else 0
            # Weighted score (simplified here)
            score = title_score
            if score >= threshold:
                if wir_row['PM_Code_Num'] > matched_pm_code:
                    matched_pm_code = wir_row['PM_Code_Num']

        matches.append({
            'ITP Reference': act_row['ITP Reference'],
            'Activity Description': act_row['Activity Description'],
            'Status': matched_pm_code
        })

    match_df = pd.DataFrame(matches)

    # ------------------------------
    # 8. Pivot Table
    # ------------------------------
    pivot_df = match_df.pivot_table(index='ITP Reference', columns='Activity Description', values='Status', fill_value=0).reset_index()

    st.subheader("Activity Completion Status Pivot Table")
    st.dataframe(pivot_df)

    # ------------------------------
    # 9. Download Result
    # ------------------------------
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pivot_df.to_excel(writer, index=False, sheet_name='PivotedStatus')
        writer.save()
        processed_data = output.getvalue()

    st.download_button(label="Download Excel", data=processed_data, file_name="ITP_WIR_Activity_Status.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

