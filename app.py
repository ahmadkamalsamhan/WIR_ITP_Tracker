import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ITP Activities Tracker", layout="wide")
st.title("ITP Activities Tracker")

# ----------------------
# Upload Excel File
# ----------------------
uploaded_file = st.file_uploader("Upload your Excel file (formatted as table)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read Excel file
        df = pd.read_excel(uploaded_file)
        st.success("Excel file loaded successfully!")

        # Show the full table
        st.subheader("Full Table")
        st.dataframe(df)

        # ----------------------
        # Filter by ITP Reference
        # ----------------------
        if 'ITP Reference' not in df.columns:
            st.error("Column 'ITP Reference' not found in the uploaded file.")
        else:
            itp_options = df['ITP Reference'].dropna().unique()
            selected_itps = st.multiselect("Select ITP Reference(s) to filter", itp_options)

            if selected_itps:
                filtered_df = df[df['ITP Reference'].isin(selected_itps)]
            else:
                filtered_df = df.copy()

            # ----------------------
            # Search in Activity Description
            # ----------------------
            if 'Activiy Description' in df.columns:
                keywords = st.text_input("Search keywords in Activity Description (comma-separated)")
                if keywords:
                    keyword_list = [kw.strip().lower() for kw in keywords.split(",")]
                    filtered_df = filtered_df[
                        filtered_df['Activiy Description'].str.lower().apply(
                            lambda x: any(kw in x for kw in keyword_list)
                        )
                    ]

            st.subheader("Filtered Results")
            st.dataframe(filtered_df)

            # ----------------------
            # Download filtered results
            # ----------------------
            def to_excel(df):
                output = BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                df.to_excel(writer, index=False, sheet_name='Filtered')
                writer.save()
                processed_data = output.getvalue()
                return processed_data

            excel_data = to_excel(filtered_df)
            st.download_button(
                label="Download Filtered Results as Excel",
                data=excel_data,
                file_name="filtered_itp_activities.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
else:
    st.info("Please upload an Excel file to start.")
