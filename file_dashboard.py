import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="📁 File Dashboard", layout="centered")
st.title("📁 File Upload Dashboard")
st.markdown("Upload Excel or Word files, preview them, and download them.")

uploaded_file = st.file_uploader("📤 Upload Excel or Word File", type=["xlsx", "xls", "docx"])

if uploaded_file:
    file_name = uploaded_file.name
    file_bytes = uploaded_file.getvalue()

    st.success(f"✅ Uploaded: `{file_name}`")

    # Allow downloading the file
    st.download_button(
        label="📥 Download File",
        data=file_bytes,
        file_name=file_name
    )

    st.divider()

    # Preview Excel
    if file_name.endswith(('.xlsx', '.xls')):
        try:
            df = pd.read_excel(uploaded_file)
            st.subheader("📊 Excel Preview")
            st.dataframe(df.head(50))  # Limit to 50 rows
        except Exception as e:
            st.error(f"❌ Error reading Excel file: {e}")

    # Preview Word
    elif file_name.endswith('.docx'):
        try:
            doc = Document(io.BytesIO(file_bytes))
            text = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
            st.subheader("📝 Word Document Preview")
            st.text_area("Content", text, height=300)
        except Exception as e:
            st.error(f"❌ Error reading Word file: {e}")
