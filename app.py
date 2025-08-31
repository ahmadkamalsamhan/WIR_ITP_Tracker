import streamlit as st
import pandas as pd
import re

st.title("📊 ITP-WIR Matching Tool (Optimized)")

# -------------------------------
# Preprocessing
# -------------------------------
def preprocess_text(text):
    if pd.isna(text):
        return ""
    return str(text).lower().strip()

# -------------------------------
# Assign Status Code
# -------------------------------
def assign_status(wir_row):
    code = str(wir_row['PM Web Code']).strip().upper()
    if code in ['A','B']:
        return 1
    elif code in ['C','D']:
        return 2
    else:
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

    st.success("✅ Files uploaded successfully!")

    # -------------------------------
    # Column Selection
    # -------------------------------
    st.subheader("Select Columns in ITP Log")
    itp_no_col = st.selectbox("ITP No.", options=itp_log.columns.tolist())
    itp_title_col = st.selectbox("ITP Title", options=itp_log.columns.tolist())

    st.subheader("Select Columns in Activity Log")
    activity_desc_col = st.selectbox("Activity Description", options=activity_log.columns.tolist())
    itp_ref_col = st.selectbox("ITP Reference", options=activity_log.columns.tolist())

    st.subheader("WIR Log Columns")
    wir_title_col = st.selectbox("WIR Title (Title / Description2)", options=wir_log.columns.tolist())
    wir_pm_col = st.selectbox("PM Web Code", options=wir_log.columns.tolist())

    if st.button("Generate Matrix"):
        st.info("Processing...")

        unique_activities = activity_log[activity_desc_col].dropna().unique().tolist()
        itp_nos = itp_log[itp_no_col].unique()
        matrix = pd.DataFrame(0, index=itp_nos, columns=unique_activities)

        for itp_no in itp_nos:
            activities = activity_log[activity_log[itp_ref_col]==itp_no]
            for _, activity_row in activities.iterrows():
                activity_desc = preprocess_text(activity_row[activity_desc_col])

                # Find matching WIR row (simple substring match)
                matched_wir = wir_log[wir_log[wir_title_col].str.lower().str.contains(activity_desc, na=False)]
                if not matched_wir.empty:
                    status_code = assign_status(matched_wir.iloc[0])
                else:
                    status_code = 0
                matrix.at[itp_no, activity_row[activity_desc_col]] = status_code

        matrix.reset_index(inplace=True)
        matrix.rename(columns={'index':'ITP No.'}, inplace=True)

        st.success("✅ Matrix Generated!")
        st.dataframe(matrix)
        st.download_button(
            label="📥 Download Matrix",
            data=matrix.to_excel(index=False, engine='openpyxl'),
            file_name="ITP_WIR_Matrix.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
