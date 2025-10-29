import streamlit as st
import pandas as pd
from docx import Document
import pypandoc
import os, tempfile, zipfile, io

st.set_page_config(page_title="Dean's List Certificates Generator", page_icon="🏅", layout="centered")
st.title("🏅 Dean's List Certificates Generator")

st.write("""
Upload your **Excel file** (Student Data) and **Word template** to generate
personalized **PDF** certificates or letters for each student.
""")

uploaded_excel = st.file_uploader("📄 Upload Student Data Sheet", type=["xlsx"])
uploaded_template = st.file_uploader("📝 Upload Word Template", type=["docx"])

if uploaded_excel and uploaded_template:
    if st.button("🚀 Generate PDFs"):
        with st.spinner("Generating personalized PDFs... ⏳"):
            with tempfile.TemporaryDirectory() as td:
                excel_path = os.path.join(td, "students.xlsx")
                template_path = os.path.join(td, "template.docx")
                open(excel_path, "wb").write(uploaded_excel.read())
                open(template_path, "wb").write(uploaded_template.read())

                df = pd.read_excel(excel_path, engine="openpyxl")

                output_pdfs = os.path.join(td, "pdfs")
                os.makedirs(output_pdfs, exist_ok=True)

                def generate_pdf(student):
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

                    tmp_docx = os.path.join(td, f"{student['NAME']}.docx")
                    doc.save(tmp_docx)

                    pdf_path = os.path.join(output_pdfs, f"{student['NAME']}.pdf")
                    pypandoc.convert_text('', 'pdf', format='docx',
                                         outputfile=pdf_path,
                                         extra_args=['--pdf-engine=xelatex', tmp_docx])

                for _, row in df.iterrows():
                    generate_pdf(row.to_dict())

                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for file in os.listdir(output_pdfs):
                        zf.write(os.path.join(output_pdfs, file), arcname=file)
                zip_buffer.seek(0)

                st.success("✅ All PDFs generated successfully!")
                st.download_button(
                    label="⬇️ Download All PDFs (ZIP)",
                    data=zip_buffer,
                    file_name="DeansList_PDFs.zip",
                    mime="application/zip"
                )
else:
    st.info("Please upload both the Excel file and Word template to proceed.")