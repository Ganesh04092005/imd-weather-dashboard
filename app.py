# import streamlit as st
# import pandas as pd
# from processor import generate_doc

# st.set_page_config(page_title="Weather Bulletin Generator")

# st.title("🌦️ Multi-Hazard Weather Bulletin Generator")

# st.markdown("Upload IMD format Excel file and generate bulletin.")

# # Upload file
# uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

# if uploaded_file:
#     st.success("✅ File uploaded successfully!")

#     # Preview raw file
#     df_preview = pd.read_excel(uploaded_file, header=None)
#     st.subheader("Preview Data")
#     st.dataframe(df_preview.head())

#     if st.button("🚀 Generate Bulletin"):
#         output_file = generate_doc(uploaded_file)

#         with open(output_file, "rb") as f:
#             st.download_button(
#                 label="📥 Download Bulletin",
#                 data=f,
#                 file_name="Weather_Bulletin.docx"
#             )

# st.markdown("---")
# st.markdown("Developed for IMD Automation Project")
import streamlit as st
import pandas as pd
import time
from processor import generate_doc

# 🔷 Page Config
st.set_page_config(
    page_title="IMD Weather Dashboard",
    page_icon="🌦️",
    layout="wide"
)

# 🔷 HEADER (LOGO + TITLE)
col1, col2 = st.columns([1, 6])

with col1:
    st.image("logo.png", width=90)

with col2:
    st.markdown("""
    <h2 style="color:#4db8ff; margin-bottom:5px;">
        IMD Weather Forecast Dashboard
    </h2>
    <p style="margin:0; font-size:15px; color:#cccccc;">
        India Meteorological Department<br>
        Ministry of Earth Sciences<br>
        Government of India
    </p>
    """, unsafe_allow_html=True)

st.markdown("<hr style='border:1px solid #444;'>", unsafe_allow_html=True)

# 🔷 SUBTITLE (FIXED - USING COLUMNS, NOT HTML IMG)
col1, col2 = st.columns([1, 8])

with col1:
    st.image("logo1.png", width=70)

with col2:
    with st.container():
        st.markdown("""
        ### ⚡ Automated Multi-Hazard Bulletin Generation System
        """)
        st.caption("Generate accurate IMD weather bulletins instantly with structured data")

st.divider()

# 🔷 SIDEBAR
with st.sidebar:
    st.header("⚙️ Controls")

    st.info("""
✔ Upload IMD Excel file  
✔ Generate bulletin automatically  
✔ Download official report  
""")

# 🔷 MAIN LAYOUT
col1, col2 = st.columns([2, 1], gap="large")

# ================= LEFT =================
with col1:
    st.markdown("""
    <div style="background:#1f2a40; padding:20px; border-radius:12px; border-left:5px solid #1f77b4;">
        <h3>📤 Upload IMD Data</h3>
        <p style="color:gray;">Upload Multi-Hazard Excel file to generate bulletin</p>
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("", type=["xlsx"])

    if uploaded_file:
        st.success("✅ File uploaded successfully!")

        df_preview = pd.read_excel(uploaded_file, header=None)

        with st.expander("🔍 Preview Data"):
            st.dataframe(df_preview.head(10))

        st.markdown("### ⚡ Generate Bulletin")

        if st.button("🚀 Generate Bulletin"):

            with st.spinner("Generating Bulletin..."):

                progress = st.progress(0)

                for i in range(100):
                    time.sleep(0.01)
                    progress.progress(i + 1)

                output_file = generate_doc(uploaded_file)

            st.success("🎉 Bulletin Generated Successfully!")

            with open(output_file, "rb") as f:
                st.download_button(
                    label="📥 Download Bulletin",
                    data=f,
                    file_name="IMD_Weather_Bulletin.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

# ================= RIGHT =================
with col2:
    st.markdown("""
    <div style="background:#1f2a40; padding:20px; border-radius:12px;">
        <h3>📌 Instructions</h3>
        <ul>
            <li>Upload IMD Excel file</li>
            <li>Click Generate Bulletin</li>
            <li>Download output file</li>
        </ul>
        <p style="color:lightgreen;">✔ Fully automated</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("""
    <div style="background:#1f2a40; padding:20px; border-radius:12px;">
        <h3>📊 Features</h3>
        <ul>
            <li>Dynamic dates</li>
            <li>District classification</li>
            <li>Template-based output</li>
            <li>Professional formatting</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

st.divider()

# 🔷 FOOTER
st.markdown("""
<hr>
<p style='text-align:center; font-size:14px; color:gray;'>
Developed by Ganesh | IMD Weather Dashboard 🌦️
</p>
""", unsafe_allow_html=True)
