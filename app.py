import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("📊 ITP-WIR Matching Tool (Optimized & Fast)")

# -------------------------------
# Preprocessing
# -------------------------------
def preprocess_text(text):
    if pd.isna(text):
        return ""
    return str(text).lower().strip()

# -------------------------------
# Assign Status
# -------------------------------
def assign_status(wir_row, pm_col):
    if pm_col not in wir_row or pd.isna(wir_row[pm_col]):
        return 0
    code = str(wir_row[pm_col]).strip().upper()
    if code in ['A','B']:
        return 1
    elif code in ['C','D']:
        return 2
    return 0

# -------------------------------
# Upload Files
# -------------------------------
itp_file = st.file_uploader("Upload ITP Log", type=["xlsx"])
activity_file = st.file_uploader("Upload ITP Activities Log", type=["xlsx"])
wir_file = st.file_uploader("Upload Document Control Log (WIR)", type=["xlsx"])

if itp_file and activity_file and wir_file:
    itp_log = pd.read_excel(itp_file)
    activity_log = pd.read_excel(activity_file)
    wir_log = pd.read_excel(wir_file)

    # Strip spaces/newlines from columns to prevent KeyErrors
    itp_log.columns = itp_log.columns.str.strip().str.replace('\n','').str.replace('\r','')
    activity_log.columns = activity_log.columns.str.strip().str.replace('\n','').str.replace('\r','')
    wir_log.columns = wir_log.columns.str.strip().str.replace('\n','').str.replace('\r','')

    st.success("✅ Files uploaded successfully!")

    # -------------------------------
    # Column Selection
    # -------------------------------
    st.subheader("ITP Log Columns")
    itp_no_col = st.selectbox("Select ITP No. column", options=itp_log.columns.tolist())
    itp_title_col = st.selectbox("Select ITP Title column", options=itp_log.columns.tolist())

    st.subheader("Activity Log Columns")
    activity_desc_col = st.selectbox("Select Activity Description column", options=activity_log.columns.tolist())
    itp_ref_col = st.selectbox("Select ITP Reference column", options=activity_log.columns.tolist())

    st.subheader("WIR Log Columns")
    wir_title_col = st.selectbox("Select WIR Title column (Title / Description2)", options=wir_log.columns.tolist())
    wir_pm_col = st.selectbox("Select PM Web Code column", options=wir_log.columns.tolist())

    # -------------------------------
    # Generate Matrix
    # -------------------------------
    if st.button("Generate ITP-WIR Matrix"):
        st.info("Processing...")

        # Build WIR lookup dictionary (title -> PM Web Code)
        wir_lookup = {}
        for idx, row in wir_log.iterrows():
            title = preprocess_text(row[wir_title_col])
            wir_lookup[title] = row[wir_pm_col]

        # Prepare matrix
        unique_activities = activity_log[activity_desc_col].dropna().unique().tolist()
        itp_nos = itp_log[itp_no_col].unique()
        matrix = pd.DataFrame(0, index=itp_nos, columns=unique_activities)

        # Fill matrix
        for itp_no in itp_nos:
            activities = activity_log[activity_log[itp_ref_col]==itp_no]
            for _, activity_row in activities.iterrows():
                activity_desc = preprocess_text(activity_row[activity_desc_col])
                status_code = 0

                # Fast substring match in WIR lookup
                for wir_title, pm_code in wir_lookup.items():
                    if activity_desc in wir_title:
                        status_code = assign_status({wir_pm_col: pm_code}, wir_pm_col)
                        break

                matrix.at[itp_no, activity_row[activity_desc_col]] = status_code

        matrix.reset_index(inplace=True)
        matrix.rename(columns={'index':'ITP No.'}, inplace=True)

        st.success("✅ ITP-WIR Matrix Generated!")
        st.dataframe(matrix)

        # -------------------------------
        # Download Excel
        # -------------------------------
        output = BytesIO()
        matrix.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button(
            label="📥 Download Matrix as Excel",
            data=output,
            file_name="ITP_WIR_Matrix.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
