import streamlit as st
import pandas as pd
from time import sleep

st.set_page_config(page_title="ITP Activities Tracker", layout="wide")
st.title("ITP Activities Tracker")

# ------------------------------
# 1. File upload
# ------------------------------
uploaded_file = st.file_uploader("Upload your ITP CSV file", type=["csv"])
if uploaded_file is not None:
    st.info("Processing file, please wait...")

    # Read CSV
    df = pd.read_csv(uploaded_file)
    
    # Clean column names
    df.columns = df.columns.str.strip()
    
    # Check required columns
    required_cols = ['Submittal Reference', 'ITP Reference', 'Checklist Reference', 'Activity No.', 'Activiy Description']
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        st.error(f"Missing required columns: {missing_cols}")
    else:
        # ------------------------------
        # 2. Group by ITP Reference
        # ------------------------------
        grouped = []
        itp_refs = df['ITP Reference'].unique()
        progress_bar = st.progress(0)

        for i, itp in enumerate(itp_refs):
            temp = df[df['ITP Reference'] == itp]
            activities = [f"{row['Activity No.']} - {row['Activiy Description']}" for idx, row in temp.iterrows()]
            grouped.append({
                'ITP Reference': itp,
                'Submittal Reference': temp['Submittal Reference'].iloc[0],
                'Checklist Reference': temp['Checklist Reference'].iloc[0],
                'Activities': activities
            })
            progress_bar.progress((i + 1) / len(itp_refs))

        grouped_df = pd.DataFrame(grouped)

        # ------------------------------
        # 3. Display & download
        # ------------------------------
        st.subheader("Grouped ITP Activities")
        st.dataframe(grouped_df)

        # Download CSV
        csv = grouped_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="Download grouped CSV",
            data=csv,
            file_name='ITP_grouped.csv',
            mime='text/csv'
        )
else:
    st.info("Please upload a CSV file to process.")
