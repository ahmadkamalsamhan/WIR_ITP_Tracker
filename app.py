import streamlit as st
import pandas as pd
from io import BytesIO

# ----------------------
# Streamlit Page Config
# ----------------------
st.set_page_config(page_title="ITP Tracker", layout="wide")
st.title("ITP Tracker - Multi File Upload")

# ----------------------
# Upload Excel Files
# ----------------------
st.subheader("Upload your Excel files")

itp_file = st.file_uploader("Upload ITP Activities Log", type=["xlsx"])
submittal_file = st.file_uploader("Upload Submittal Reference file", type=["xlsx"])
checklist_file = st.file_uploader("Upload Checklist Reference file", type=["xlsx"])

if itp_file and submittal_file and checklist_file:
    try:
        # Read Excel files
        itp_df = pd.read_excel(itp_file)
        submittal_df = pd.read_excel(submittal_file)
        checklist_df = pd.read_excel(checklist_file)

        st.success("All files loaded successfully!")

        st.subheader("ITP Activities Log")
        st.dataframe(itp_df)

        st.subheader("Submittal Reference Table")
        st.dataframe(submittal_df)

        st.subheader("Checklist Reference Table")
        st.dataframe(checklist_df)

        # ----------------------
        # Filter by ITP Reference
        # ----------------------
        if 'ITP Reference' not in itp_df.columns:
            st.error("Column 'ITP Reference' not found in ITP Activities Log.")
        else:
            itp_options = itp_df['ITP Reference'].dropna().unique()
            selected_itps = st.multiselect("Select ITP Reference(s) to filter", itp_options)

            filtered_itp = itp_df.copy()
            if selected_itps:
                filtered_itp = filtered_itp[filtered_itp['ITP Reference'].isin(selected_itps)]

            # ----------------------
            # Filter by Clause Number (optional)
            # ----------------------
            if 'Clause Number' in filtered_itp.columns:
                clause_options = filtered_itp['Clause Number'].dropna().unique()
                selected_clauses = st.multiselect("Select Clause Number(s) to filter", clause_options)
                if selected_clauses:
                    filtered_itp = filtered_itp[filtered_itp['Clause Number'].isin(selected_clauses)]

            # ----------------------
            # Search in Activity Description
            # ----------------------
            if 'Activiy Description' in filtered_itp.columns:
                keywords = st.text_input("Search keywords in Activity Description (comma-separated)")
                if keywords:
                    keyword_list = [kw.strip().lower() for kw in keywords.split(",")]
                    filtered_itp = filtered_itp[
                        filtered_itp['Activiy Description'].str.lower().apply(
                            lambda x: any(kw in x for kw in keyword_list)
                        )
                    ]

            st.subheader("Filtered ITP Activities")
            st.dataframe(filtered_itp)

            # ----------------------
            # Merge with Submittal and Checklist
            # ----------------------
            merge_df = filtered_itp.merge(
                submittal_df, on='Submittal Reference', how='left'
            ).merge(
                checklist_df, on='Checklist Reference', how='left'
            )

            st.subheader("Merged Table (ITP + Submittal + Checklist)")
            st.dataframe(merge_df)

            # ----------------------
            # Download merged results
            # ----------------------
            def to_excel(df):
                output = BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                df.to_excel(writer, index=False, sheet_name='Merged')
                writer.save()
                processed_data = output.getvalue()
                return processed_data

            excel_data = to_excel(merge_df)
            st.download_button(
                label="Download Filtered & Merged Results as Excel",
                data=excel_data,
                file_name="merged_itp_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error loading Excel files: {e}")

else:
    st.info("Please upload all three Excel files to proceed.")
