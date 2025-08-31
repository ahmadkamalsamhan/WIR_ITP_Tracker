import streamlit as st
import pandas as pd
from rapidfuzz import fuzz, process
import nltk
from nltk.corpus import stopwords

# Download NLTK stopwords (once)
nltk.download('stopwords')

# Streamlit page setup
st.set_page_config(page_title="ITP Tracker", layout="wide")
st.title("ITP Tracker - Fuzzy Matching App")

# File uploader
uploaded_file = st.file_uploader("Upload Excel or CSV file", type=['xlsx','csv'])

if uploaded_file:
    try:
        # Load file
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.success(f"File '{uploaded_file.name}' loaded successfully!")
        st.dataframe(df.head())

        # Select column for fuzzy search
        col_to_search = st.selectbox("Select column to search in", df.columns)

        # Input search terms (comma-separated)
        search_input = st.text_area("Enter search terms (comma-separated)")

        if search_input:
            search_terms = [term.strip() for term in search_input.split(",") if term.strip()]
            
            # Fuzzy match function
            def get_best_matches(value):
                matches = [(term, fuzz.partial_ratio(str(value), term)) for term in search_terms]
                matches.sort(key=lambda x: x[1], reverse=True)
                return matches[0]  # best match

            # Apply fuzzy match
            df['Best Match'] = df[col_to_search].apply(lambda x: get_best_matches(x)[0])
            df['Match Score'] = df[col_to_search].apply(lambda x: get_best_matches(x)[1])

            # Filter results by minimum score
            min_score = st.slider("Minimum match score", 0, 100, 60)
            filtered_df = df[df['Match Score'] >= min_score]

            st.subheader("Filtered Results")
            st.dataframe(filtered_df)

            # Download filtered results
            def convert_df(df):
                return df.to_csv(index=False).encode('utf-8')

            csv = convert_df(filtered_df)
            st.download_button(
                label="Download filtered results as CSV",
                data=csv,
                file_name='filtered_results.csv',
                mime='text/csv'
            )

    except Exception as e:
        st.error(f"Error loading file: {e}")

else:
    st.info("Please upload a CSV or Excel file to start.")
