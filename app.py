import streamlit as st
import pandas as pd
import time
import os
import tempfile
from processor import generate_doc, get_district_preview

# ─── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="IMD Weather Dashboard",
    page_icon="🌦️",
    layout="wide"
)

# ─── Styling ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    [data-testid="stAppViewContainer"] { background-color: #0e1117; }
    [data-testid="stSidebar"] { background-color: #161b22; }
    body, .stMarkdown, p, li { color: #e0e0e0; }
</style>
""", unsafe_allow_html=True)

# ─── Header ─────────────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 6])
with col1:
    st.image("logo.png", width=90)
with col2:
    st.markdown("""
    <h2 style="color:#4db8ff; margin-bottom:5px;">
        IMD Weather Forecast Dashboard
    </h2>
    <p style="margin:0; font-size:15px; color:#cccccc;">
        India Meteorological Department &nbsp;|&nbsp;
        Ministry of Earth Sciences &nbsp;|&nbsp;
        Government of India
    </p>
    """, unsafe_allow_html=True)

st.markdown("<hr style='border:1px solid #444;'>", unsafe_allow_html=True)

col1, col2 = st.columns([1, 8])
with col1:
    st.image("logo1.png", width=70)
with col2:
    st.markdown("### ⚡ Automated Multi-Hazard Bulletin Generation System")
    st.caption("Generate accurate IMD weather bulletins instantly with structured data")

st.divider()

# ─── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Controls")
    st.info("✔ Upload IMD Excel file\n✔ Generate bulletin automatically\n✔ Download official report")

# ─── Main Layout ────────────────────────────────────────────────────────────────
col1, col2 = st.columns([2, 1], gap="large")

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "imd_template_fixed.docx")

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

        # ── Improved district preview ─────────────────────────────────────
        with st.expander("🔍 Preview Data"):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_prev:
                    tmp_prev.write(uploaded_file.getvalue())
                    tmp_prev_path = tmp_prev.name

                df_preview = get_district_preview(tmp_prev_path)
                os.remove(tmp_prev_path)

                st.markdown(f"**{len(df_preview)} districts loaded**")

                # Row-level color highlights based on D1 Warning severity
                def highlight_warning(row):
                    level = str(row.get("D1 Warning", "")).upper()
                    if "EXHVY" in level:
                        return ["background-color:#FF000044"] * len(row)
                    elif "VHVY" in level:
                        return ["background-color:#FF660044"] * len(row)
                    elif "IHVY" in level:
                        return ["background-color:#FFFF0033"] * len(row)
                    return [""] * len(row)

                st.dataframe(
                    df_preview.style.apply(highlight_warning, axis=1),
                    use_container_width=True,
                    height=min(400, 35 + len(df_preview) * 35)
                )

            except Exception as e:
                df_raw = pd.read_excel(uploaded_file, header=None)
                st.dataframe(df_raw.head(10))
                st.warning(f"District parse error: {e}")

        st.markdown("### ⚡ Generate Bulletin")

        if st.button("🚀 Generate Bulletin"):
            with st.spinner("Generating Bulletin..."):
                progress = st.progress(0)
                for i in range(80):
                    time.sleep(0.01)
                    progress.progress(i + 1)

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name

                output_file = generate_doc(tmp_path, template_path=TEMPLATE_PATH)

                for i in range(80, 100):
                    time.sleep(0.005)
                    progress.progress(i + 1)

            st.success("🎉 Bulletin Generated Successfully!")

            with open(output_file, "rb") as f:
                st.download_button(
                    label="📥 Download Bulletin",
                    data=f,
                    file_name="IMD_Weather_Bulletin.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

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
            <li>Dynamic dates &amp; time</li>
            <li>All-district preview</li>
            <li>Colour-coded warning cells</li>
            <li>District-level classification</li>
            <li>Severity grouping (EH / VH / H)</li>
            <li>Thunderstorm + Gusty Winds</li>
            <li>No warning fallback text</li>
            <li>Professional IMD formatting</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

st.divider()
