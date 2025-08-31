import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io

# -----------------------------
# Normalization helper
# -----------------------------
def normalize_text(text):
    if pd.isna(text):
        return ""
    return str(text).strip().lower().replace("\n", " ")

# -----------------------------
# Main app
# -----------------------------
st.set_page_config(page_title="WIR ↔ ITP Activity Matcher", layout="wide")

st.title("🔎 WIR ↔ ITP Activity Smart Matcher (Top 3 Matches)")

st.write("Upload your CSV files for **ITP Activities Log**, **ITP Log**, and **WIR Log**.")

# File uploads
itp_act_file = st.file_uploader("Upload ITP Activities Log (CSV)", type=["csv"])
itp_log_file = st.file_uploader("Upload ITP Log (CSV)", type=["csv"])
wir_log_file = st.file_uploader("Upload WIR Log (CSV)", type=["csv"])

if itp_act_file and itp_log_file and wir_log_file:
    try:
        st.info("Reading CSV files...")
        itp_activities = pd.read_csv(itp_act_file)
        itp_log = pd.read_csv(itp_log_file)
        wir_log = pd.read_csv(wir_log_file)

        # Normalize activity descriptions
        itp_activities["ActivityNorm"] = itp_activities["Activiy Description"].apply(normalize_text)

        # Expand multi-activity WIR titles
        st.info("Expanding multi-activity WIR rows...")
        wir_expanded = []
        for _, row in wir_log.iterrows():
            activities = str(row["Title / Description"]).split(",")
            for act in activities:
                new_row = row.copy()
                new_row["WIR_Activity"] = normalize_text(act)
                wir_expanded.append(new_row)
        wir_df = pd.DataFrame(wir_expanded)
        st.success(f"Expanded WIR activities: {len(wir_df)} rows total")

        # Matching
        st.info("Matching WIR ↔ ITP activities...")
        matches = []
        progress = st.progress(0)

        wir_acts = wir_df["WIR_Activity"].tolist()
        itp_acts = itp_activities["ActivityNorm"].tolist()

        for i, wir_act in enumerate(wir_acts):
            top_matches = process.extract(
                wir_act,
                itp_acts,
                scorer=fuzz.token_sort_ratio,
                limit=3
            )
            for rank, (match, score, idx) in enumerate(top_matches, 1):
                matches.append({
                    "WIR_Activity": wir_act,
                    f"Match_{rank}": itp_activities.iloc[idx]["Activiy Description"],
                    f"Score_{rank}": score
                })

            if i % 100 == 0 or i == len(wir_acts)-1:
                progress.progress(int((i+1)/len(wir_acts)*100))

        st.success("✅ Matching completed!")

        result_df = pd.DataFrame(matches)

        st.write("### Sample Results (Top 3 matches per WIR activity):")
        st.dataframe(result_df.head(50), use_container_width=True)

        # Download button
        buffer = io.BytesIO()
        result_df.to_csv(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            label="📥 Download Full Results as CSV",
            data=buffer,
            file_name="wir_itp_matches.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"Error processing files: {e}")
