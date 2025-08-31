import streamlit as st
import pandas as pd
import numpy as np
from rapidfuzz import fuzz
import io

st.set_page_config(page_title="WIR → ITP Tracker", layout="wide")
st.title("WIR → ITP Activity Tracker")

# -------------------------
# Upload Files
# -------------------------
wir_file = st.file_uploader("Upload WIR Log (.xlsx)", type=["xlsx"])
itp_file = st.file_uploader("Upload ITP Log (.xlsx)", type=["xlsx"])
activity_file = st.file_uploader("Upload ITP Activities Log (.xlsx)", type=["xlsx"])

threshold = st.slider("Fuzzy Match Threshold (%)", 50, 100, 80)

if wir_file and itp_file and activity_file:
    if st.button("Start Processing"):
        st.info("Reading Excel files...")
        wir_df = pd.read_excel(wir_file)
        itp_df = pd.read_excel(itp_file)
        act_df = pd.read_excel(activity_file)

        # -------------------------
        # Clean column names
        # -------------------------
        def clean_columns(df):
            df.columns = df.columns.str.strip().str.replace('\n',' ').str.replace('\r',' ')
            return df

        wir_df = clean_columns(wir_df)
        itp_df = clean_columns(itp_df)
        act_df = clean_columns(act_df)

        wir_col = [c for c in wir_df.columns if 'Title / Description2' in c][0]
        itp_col = [c for c in itp_df.columns if 'Title / Description' in c][0]
        act_itp_col = [c for c in act_df.columns if 'ITP Reference' in c][0]
        act_desc_col = [c for c in act_df.columns if 'Activiy Description' in c][0]

        # -------------------------
        # Normalize
        # -------------------------
        wir_df['WIR_Norm'] = wir_df[wir_col].astype(str).str.upper().str.replace("-", "").str.replace(" ", "")
        itp_df['ITP_Norm'] = itp_df[itp_col].astype(str).str.upper().str.replace("-", "").str.replace(" ", "")
        act_df['ITP_Ref_Norm'] = act_df[act_itp_col].astype(str).str.upper().str.replace("-", "").str.replace(" ", "")
        act_df['ActivityDescNorm'] = act_df[act_desc_col].astype(str).str.upper().str.replace("-", "").str.replace(" ", "")

        # -------------------------
        # Matching WIR → ITP
        # -------------------------
        st.info("Matching WIRs to ITPs...")
        progress_bar = st.progress(0)
        status_text = st.empty()

        results = []

        total_wir = len(wir_df)
        for idx, wir_row in wir_df.iterrows():
            best_match_score = 0
            best_itp_idx = None
            for jdx, itp_row in itp_df.iterrows():
                score = fuzz.ratio(wir_row['WIR_Norm'], itp_row['ITP_Norm'])
                if score > best_match_score:
                    best_match_score = score
                    best_itp_idx = jdx
            if best_match_score >= threshold:
                matched_itp = itp_df.loc[best_itp_idx, itp_col]
                # check activities selected
                activities = act_df[act_df['ITP_Ref_Norm'] == str(itp_df.loc[best_itp_idx, itp_col]).upper().replace("-", "").replace(" ", "")]
                activities_list = activities[act_desc_col].tolist()
            else:
                matched_itp = None
                activities_list = []
            results.append({
                'WIR_Title': wir_row[wir_col],
                'Matched_ITP': matched_itp,
                'ITP_Activities': ", ".join(activities_list),
                'Score': best_match_score
            })
            progress_bar.progress((idx + 1) / total_wir)
            status_text.text(f"Processed {idx + 1}/{total_wir} WIR rows")

        result_df = pd.DataFrame(results)

        st.subheader("Matched Results")
        st.dataframe(result_df)

        # -------------------------
        # Save to Excel
        # -------------------------
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='MatchedResults')
        output.seek(0)

        st.download_button(
            label="Download Excel",
            data=output.getvalue(),
            file_name="WIR_ITP_Matched.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("Processing completed!")
