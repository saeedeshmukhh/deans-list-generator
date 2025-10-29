import streamlit as st
import pandas as pd
from docx import Document
from docx2pdf import convert
import os, tempfile, zipfile, io

st.set_page_config(page_title="Dean's List PDF Generator", page_icon="🏅", layout="centered")
st.title("🏅 Dean's List Certficates Generator")

st.write("Upload your Excel file (Student Data with complete details) and Word template ( MS / MBA Template) to generate personalized letters with Certificates for all students.")

uploaded_excel = st.file_uploader("📄 Upload Student Data sheet", type=["xlsx"])
uploaded_template = st.file_uploader("📝 Upload Word Template", type=["docx"])

if uploaded_excel and uploaded_template:
    if st.button("🚀 Generate PDFs"):
        with tempfile.TemporaryDirectory() as td:
            # Save uploads
            excel_path = os.path.join(td, "students.xlsx")
            template_path = os.path.join(td, "template.docx")
            open(excel_path, "wb").write(uploaded_excel.read())
            open(template_path, "wb").write(uploaded_template.read())

            # Read Excel
            df = pd.read_excel(excel_path, engine="openpyxl")

            output_docs = os.path.join(td, "docs")
            output_pdfs = os.path.join(td, "pdfs")
            os.makedirs(output_docs, exist_ok=True)
            os.makedirs(output_pdfs, exist_ok=True)

            # Core function (from your Program.py)
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

                import subprocess, shutil, tempfile

                def convert_to_pdf(filepath, output_pdf):
                    soffice = shutil.which("soffice")
                    if not soffice:
                        raise RuntimeError("LibreOffice not found — install with 'brew install --cask libreoffice'")
                    tmpdir = tempfile.mkdtemp()
                    subprocess.run(
                        [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, filepath],
                        check=True
                    )
                    pdf_file = os.path.join(tmpdir, os.path.splitext(os.path.basename(filepath))[0] + ".pdf")
                    shutil.move(pdf_file, output_pdf)
                    shutil.rmtree(tmpdir, ignore_errors=True)

                # ...
                convert_to_pdf(filepath, os.path.join(output_pdfs, f"{student['NAME']}.pdf"))

            # Iterate
            for _, row in df.iterrows():
                student_data = row.to_dict()
                generate_letter(student_data)

            # Zip results
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
    st.info("Please upload both the Excel file and the Word template to proceed.")