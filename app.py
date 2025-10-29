import streamlit as st
import pandas as pd
from docx import Document
import os, tempfile, zipfile, io

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(page_title="Dean's List Certificates Generator", page_icon="🏅", layout="centered")
st.title("🏅 Dean's List Certificates Generator")

st.write("""
Upload your **Excel file** (Student Data with complete details)
and the **Word template** (MS / MBA Template) to generate personalized
Dean’s List letters for all students.
""")

# --------------------------------------------------
# FILE UPLOADS
# --------------------------------------------------
uploaded_excel = st.file_uploader("📄 Upload Student Data Sheet", type=["xlsx"])
uploaded_template = st.file_uploader("📝 Upload Word Template", type=["docx"])

# --------------------------------------------------
# MAIN PROCESS
# --------------------------------------------------
if uploaded_excel and uploaded_template:
    if st.button("🚀 Generate Letters"):
        with st.spinner("Generating personalized DOCX letters... ⏳"):
            with tempfile.TemporaryDirectory() as td:
                # Save uploaded files
                excel_path = os.path.join(td, "students.xlsx")
                template_path = os.path.join(td, "template.docx")
                open(excel_path, "wb").write(uploaded_excel.read())
                open(template_path, "wb").write(uploaded_template.read())

                # Read Excel
                df = pd.read_excel(excel_path, engine="openpyxl")

                # Output directory
                output_docs = os.path.join(td, "docs")
                os.makedirs(output_docs, exist_ok=True)

                # Core function
                def generate_letter(student):
                    doc = Document(template_path)
                    for p in doc.paragraphs:
                        for key, value in student.items():
                            placeholder = f"[{key.upper()}]"
                            if placeholder in p.text:
                                for run in p.runs:
                                    if placeholder in run.text:
                                        run.text = run.text.replace(placeholder, str(value))

                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for key, value in student.items():
                                    placeholder = f"[{key.upper()}]"
                                    if placeholder in cell.text:
                                        cell.text = cell.text.replace(placeholder, str(value))

                    filename = f"{student['NAME']}_DeansList.docx"
                    filepath = os.path.join(output_docs, filename)
                    doc.save(filepath)

                # Generate all
                for _, row in df.iterrows():
                    generate_letter(row.to_dict())

                # Zip results
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for file in os.listdir(output_docs):
                        zf.write(os.path.join(output_docs, file), arcname=file)
                zip_buffer.seek(0)

                st.success("✅ All letters generated successfully!")
                st.download_button(
                    label="⬇️ Download All Letters (ZIP)",
                    data=zip_buffer,
                    file_name="DeansList_Letters.zip",
                    mime="application/zip"
                )
else:
    st.info("Please upload both the Excel file and Word template to proceed.")