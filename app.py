import streamlit as st
import pandas as pd
from processor import generate_doc

st.set_page_config(page_title="Weather Bulletin Generator")

st.title("🌦️ Multi-Hazard Weather Bulletin Generator")

st.markdown("Upload IMD format Excel file and generate bulletin.")

# Upload file
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

if uploaded_file:
    st.success("✅ File uploaded successfully!")

    # Preview raw file
    df_preview = pd.read_excel(uploaded_file, header=None)
    st.subheader("Preview Data")
    st.dataframe(df_preview.head())

    if st.button("🚀 Generate Bulletin"):
        output_file = generate_doc(uploaded_file)

        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Download Bulletin",
                data=f,
                file_name="Weather_Bulletin.docx"
            )

st.markdown("---")
st.markdown("Developed for IMD Automation Project")