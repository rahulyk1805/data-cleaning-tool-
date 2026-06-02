# ============================================================
# app.py
# ============================================================

import streamlit as st
import tempfile
import os

from cleaner import (
    detect_file_type,
    clean_sales_file,
    clean_purchase_file,
    clean_stock_file,
    clean_jv_file,
    clean_sundry_creditors_file,
    save_cleaned_workbook,
    generate_summary_txt
)

# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Data Cleaning Tool",
    page_icon="📊",
    layout="wide"
)

# ============================================================
# BEAUTIFUL UI
# ============================================================

st.markdown("""
<style>

.main {
    background-color: #f4f7fc;
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}

.main-title {
    font-size: 58px;
    font-weight: 800;
    color: #1a1a2e;
    text-align: center;
    margin-bottom: 5px;
}

.sub-title {
    text-align: center;
    font-size: 24px;
    color: #4b5563;
    margin-bottom: 40px;
}

.feature-box {
    background: white;
    padding: 30px;
    border-radius: 18px;
    box-shadow: 0px 4px 18px rgba(0,0,0,0.08);
    margin-bottom: 25px;
}

.feature-title {
    font-size: 28px;
    font-weight: 700;
    color: #111827;
    margin-bottom: 15px;
}

.feature-text {
    font-size: 18px;
    color: #374151;
    line-height: 1.8;
}

.support-box {
    background: linear-gradient(135deg,#2563eb,#1d4ed8);
    padding: 30px;
    border-radius: 18px;
    color: white;
    box-shadow: 0px 4px 18px rgba(0,0,0,0.12);
}

.support-title {
    font-size: 30px;
    font-weight: bold;
    margin-bottom: 20px;
}

.support-item {
    font-size: 19px;
    margin-bottom: 10px;
}

.upload-box {
    background: white;
    padding: 25px;
    border-radius: 18px;
    box-shadow: 0px 4px 18px rgba(0,0,0,0.08);
    margin-top: 25px;
}

.file-card {
    background: white;
    padding: 25px;
    border-radius: 18px;
    box-shadow: 0px 4px 18px rgba(0,0,0,0.08);
    margin-top: 30px;
}

.footer {
    text-align: center;
    color: gray;
    margin-top: 40px;
    font-size: 16px;
}

</style>
""", unsafe_allow_html=True)

# ============================================================
# HEADER
# ============================================================

st.markdown("""
<div class="main-title">
    📊 DATA CLEANING TOOL
</div>

<div class="sub-title">
    Developed by <b>Jaishil Industries</b>
</div>
""", unsafe_allow_html=True)

# ============================================================
# INFO SECTION
# ============================================================

col1, col2 = st.columns([2,1])

with col1:

    st.markdown("""
    <div class="feature-box">

    <div class="feature-title">
    🚀 Smart Automated Excel Cleaning
    </div>

    <div class="feature-text">

    This intelligent system automatically:

    ✅ Detects file type from uploaded Excel files  
    ✅ Cleans raw accounting and inventory data  
    ✅ Removes unwanted entries  
    ✅ Corrects typo issues  
    ✅ Generates cleaned Excel output  
    ✅ Creates change logs automatically  
    ✅ Supports manual file type selection  
    ✅ Processes multiple files together  

    </div>

    </div>
    """, unsafe_allow_html=True)

with col2:

    st.markdown("""
    <div class="support-box">

    <div class="support-title">
    📁 Supported File Types
    </div>

    <div class="support-item">✔ Sales</div>
    <div class="support-item">✔ Purchase</div>
    <div class="support-item">✔ Stock</div>
    <div class="support-item">✔ JV</div>
    <div class="support-item">✔ HO EXP</div>
    <div class="support-item">✔ Sundry Creditors</div>
    <div class="support-item">✔ Unknown Files</div>

    </div>
    """, unsafe_allow_html=True)

# ============================================================
# UPLOAD
# ============================================================

st.markdown("<br>", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "📤 Upload Excel Files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# ============================================================
# PROCESS FILES
# ============================================================

if uploaded_files:

    for uploaded_file in uploaded_files:

        st.markdown("<hr>", unsafe_allow_html=True)

        st.markdown(f"""
        <div class="file-card">
        <h2>📄 {uploaded_file.name}</h2>
        </div>
        """, unsafe_allow_html=True)

        temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp_input.write(uploaded_file.read())
        temp_input.close()

        detected_type = detect_file_type(temp_input.name)

        st.success(f"Detected File Type: {detected_type}")

        # ====================================================
        # AUTO CLEANING
        # ====================================================

        if detected_type == "SALES":
            raw_df, cleaned_df, changes = clean_sales_file(temp_input.name)

        elif detected_type == "PURCHASE":
            raw_df, cleaned_df, changes = clean_purchase_file(temp_input.name)

        elif detected_type == "STOCK":
            raw_df, cleaned_df, changes = clean_stock_file(temp_input.name)

        elif detected_type == "JV":
            raw_df, cleaned_df, changes = clean_jv_file(temp_input.name)

        elif detected_type == "SUNDRY CREDITORS":
            raw_df, cleaned_df, changes = clean_sundry_creditors_file(temp_input.name)

        else:
            st.warning("Unable to detect file type. Generic cleaning applied.")
            raw_df, cleaned_df, changes = clean_sales_file(temp_input.name)

        output_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_excel.close()

        save_cleaned_workbook(
            raw_df,
            cleaned_df,
            changes,
            output_excel.name,
            "Cleaned_Data"
        )

        summary_file = tempfile.NamedTemporaryFile(delete=False, suffix=".txt")
        summary_file.close()

        generate_summary_txt(
            uploaded_file.name,
            detected_type,
            changes,
            summary_file.name
        )

        # ====================================================
        # DOWNLOADS
        # ====================================================

        col1, col2 = st.columns(2)

        with col1:

            with open(output_excel.name, "rb") as f:

                st.download_button(
                    "⬇ Download Cleaned Excel",
                    f,
                    file_name=f"cleaned_{uploaded_file.name}.xlsx"
                )

        with col2:

            with open(summary_file.name, "rb") as f:

                st.download_button(
                    "⬇ Download Changes Summary",
                    f,
                    file_name=f"summary_{uploaded_file.name}.txt"
                )

        # ====================================================
        # MANUAL FILE TYPE
        # ====================================================

        st.markdown("### 🔧 Manual File Type Selection")

        manual_type = st.selectbox(
            f"Select manual type for {uploaded_file.name}",
            [
                "SALES",
                "PURCHASE",
                "STOCK",
                "JV",
                "SUNDRY CREDITORS"
            ],
            key=uploaded_file.name
        )

        if st.button(f"Generate Manual Cleaned File - {uploaded_file.name}"):

            if manual_type == "SALES":
                raw_df, cleaned_df, changes = clean_sales_file(temp_input.name)

            elif manual_type == "PURCHASE":
                raw_df, cleaned_df, changes = clean_purchase_file(temp_input.name)

            elif manual_type == "STOCK":
                raw_df, cleaned_df, changes = clean_stock_file(temp_input.name)

            elif manual_type == "JV":
                raw_df, cleaned_df, changes = clean_jv_file(temp_input.name)

            elif manual_type == "SUNDRY CREDITORS":
                raw_df, cleaned_df, changes = clean_sundry_creditors_file(temp_input.name)

            manual_output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            manual_output.close()

            save_cleaned_workbook(
                raw_df,
                cleaned_df,
                changes,
                manual_output.name,
                "Manual_Cleaned"
            )

            with open(manual_output.name, "rb") as f:

                st.download_button(
                    "⬇ Download Manual Cleaned Excel",
                    f,
                    file_name=f"manual_cleaned_{uploaded_file.name}.xlsx"
                )

# ============================================================
# FOOTER
# ============================================================

st.markdown("""
<div class="footer">
© 2026 Jaishil Industries | Smart Data Cleaning Platform
</div>
""", unsafe_allow_html=True)
