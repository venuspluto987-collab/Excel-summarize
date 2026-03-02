import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime

# Page config
st.set_page_config(page_title="Excel Name Summarizer", layout="wide")

# 🎨 Custom Pink-Red + Purple Theme
st.markdown("""
<style>

/* Main Background Gradient */
.stApp {
    background: linear-gradient(135deg, #ff4b6e, #8a2be2);
    color: white;
}

/* Make text white */
h1, h2, h3, h4, h5, h6, p, label {
    color: white !important;
}

/* File uploader browse button */
div[data-testid="stFileUploader"] button {
    background-color: #ff2e63 !important;
    color: white !important;
    border-radius: 10px !important;
    border: none !important;
    padding: 8px 18px !important;
    font-weight: 600 !important;
}

/* Hover effect */
div[data-testid="stFileUploader"] button:hover {
    background-color: #e6004c !important;
    color: white !important;
}

/* Upload box styling */
div[data-testid="stFileUploader"] {
    background-color: rgba(255, 255, 255, 0.1);
    padding: 20px;
    border-radius: 12px;
}

/* Download button styling */
.stDownloadButton>button {
    background-color: #9d4edd !important;
    color: white !important;
    border-radius: 10px !important;
    border: none !important;
    padding: 8px 18px !important;
}

.stDownloadButton>button:hover {
    background-color: #7b2cbf !important;
    color: white !important;
}

</style>
""", unsafe_allow_html=True)

st.title("📊 Excel Automation - Summarization by Name")

st.write("Upload an Excel file to automatically group and summarize data by Name.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    st.subheader("📄 Original Data")
    st.dataframe(df, use_container_width=True)

    if "Name" in df.columns:
        numeric_cols = df.select_dtypes(include="number").columns
        summary = df.groupby("Name")[numeric_cols].sum().reset_index()

        st.subheader("✅ Summarized Data")
        st.dataframe(summary, use_container_width=True)

        today = datetime.now().strftime("%Y-%m-%d")
        output_folder = Path(f"output_files/output_{today}")
        output_folder.mkdir(parents=True, exist_ok=True)

        output_path = output_folder / "summary.xlsx"
        summary.to_excel(output_path, index=False)

        st.success("Summary file generated successfully!")

        with open(output_path, "rb") as file:
            st.download_button(
                label="⬇ Download Summary Excel",
                data=file,
                file_name="summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("⚠ The Excel file must contain a column named 'Name'.")

else:
    st.info("Please upload an Excel file to begin.")