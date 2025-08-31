# ===============================
# ITP-WIR Matching Tool (Interactive Columns)
# ===============================

import streamlit as st
import pandas as pd
import re
from sentence_transformers import SentenceTransformer, util

st.title("📊 ITP-WIR Matching Tool (Online)")

# -------------------------------
# Load Sentence Transformer
# -------------------------------
@st.cache_resource
def load_model():
    return SentenceTransformer('all-MiniLM-L6-v2')

model = load_model()

# -------------------------------
# Preprocessing
# -------------------------------
def preprocess_text(text):
    if pd.isna(text):
        return []
    text = str(text).lower()
    text = re.sub(r'[^a-z0-9\s]', '', text)
    return text.split()

# -------------------------------
# Match activity to WIR
# -------------------------------
def match_activity_to_wir(activity, candidate_wirs, wir_title_col):
    best_match = None
    max_score = -1

    for idx, wir in candidate_wirs.iterrows():
        activity_tokens = set(preprocess_text(activity))
        wir_tokens = set(preprocess_text(wir[wir_title_col]))
        token_score = len(activity_tokens & wir_tokens)

        activity_emb = model.encode(activity, convert_to_tensor=True)
        wir_emb = model.encode(str(wir[wir_title_col]), convert_to_tensor=True)
        semantic_score = util.cos_sim(activity_emb, wir_emb).item()

        total_score = token_score + semantic_score

        if total_score > max_score:
            max_score = total_score
            best_match = wir

    return best_match

# -------------------------------
# Assign Status Code
# -------------------------------
def assign_status(wir, wir_pm_col):
    if wir is None or pd.isna(wir.get(wir_pm_col)):
        return 0
    code = str(wir[wir_pm_col]).strip().upper()
    if code in ['A','B']:
        return 1
    elif code in ['C','D']:
        return 2
    else:
        return 0

# -------------------------------
# Upload Files
# -------------------------------
st.header("Step 1: Upload Excel Files")
itp_file = st.file_uploader("Upload ITP Log", type=["xlsx"])
activity_file = st.file_uploader("Upload ITP Activities Log", type=["xlsx"])
wir_file = st.file_uploader("Upload Document Control Log (WIR)", type=["xlsx"])

if itp_file and activity_file and wir_file:
    itp_log = pd.read_excel(itp_file)
    activity_log = pd.read_excel(activity_file)
    wir_log = pd.read_excel(wir_file)

    st.success("✅ Files uploaded successfully!")

    # -------------------------------
    # Column Selection for Each File
    # -------------------------------
    st.header("Step 2: Select Columns")

    # ITP Log
    st.subheader("ITP Log Columns")
    st.write(itp_log.columns.tolist())
    itp_no_col = st.selectbox("Select column for ITP No.", options=itp_log.columns.tolist())
    itp_title_col = st.selectbox("Select column for ITP Title", options=itp_log.columns.tolist())

    # Activity Log
    st.subheader("ITP Activity Log Columns")
    st.write(activity_log.columns.tolist())
    activity_desc_col = st.selectbox("Select column for Activity Description", options=activity_log.columns.tolist())
    itp_ref_col = st.selectbox("Select column for ITP Reference", options=activity_log.columns.tolist())
    activity_no_col = st.selectbox("Select column for Activity No.", options=activity_log.columns.tolist())

    # WIR Log
    st.subheader("WIR Log Columns")
    st.write(wir_log.columns.tolist())
    wir_title_col = st.selectbox("Select column for WIR Title", options=wir_log.columns.tolist())
    wir_func_col = st.selectbox("Select column for Function", options=wir_log.columns.tolist())
    wir_pm_col = st.selectbox("Select column for PM Web Code", options=wir_log.columns.tolist())

    st.header("Step 3: Generate Matrix")

    if st.button("Generate ITP-WIR Matrix"):
        st.info("Processing... This may take a few minutes depending on the file size.")

        # Extract unique activities
        unique_activities = activity_log[activity_desc_col].dropna().unique().tolist()
        itp_nos = itp_log[itp_no_col].unique()
        matrix = pd.DataFrame(0, index=itp_nos, columns=unique_activities)

        progress = st.progress(0)
        total = len(itp_nos)

        for i, itp_no in enumerate(itp_nos):
            itp_row = itp_log[itp_log[itp_no_col]==itp_no].iloc[0]
            itp_title = itp_row[itp_title_col]

            # Get activities for this ITP
            activities = activity_log[activity_log[itp_ref_col]==itp_no]

            # Candidate WIRs (optionally, can filter by Function later)
            candidate_wirs = wir_log

            for _, activity_row in activities.iterrows():
                activity_desc = activity_row[activity_desc_col]
                best_wir = match_activity_to_wir(activity_desc, candidate_wirs, wir_title_col)
                status_code = assign_status(best_wir, wir_pm_col)
                matrix.at[itp_no, activity_desc] = status_code

            progress.progress((i+1)/total)

        matrix.reset_index(inplace=True)
        matrix.rename(columns={'index':'ITP No.'}, inplace=True)

        st.success("✅ ITP-WIR Matrix Generated!")
        st.dataframe(matrix)

        st.download_button(
            label="📥 Download Matrix as Excel",
            data=matrix.to_excel(index=False, engine='openpyxl'),
            file_name="ITP_WIR_Matrix.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
