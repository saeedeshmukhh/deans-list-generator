import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
import os, tempfile, zipfile, io

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(page_title="Dean's List Certificates Generator", page_icon="🏅", layout="centered")
st.title("🏅 Dean's List Certificates Generator")

st.write("""
Upload your **Excel file** (student data) to generate personalized **Dean’s List PDFs**  
for each student, including name, program, term, and GPA.
""")

logo_path = "logo.jpg"  # put your SCU logo in same folder (rename if needed)

# --------------------------------------------------
# FILE UPLOAD
# --------------------------------------------------
uploaded_excel = st.file_uploader("📄 Upload Student Data Sheet", type=["xlsx"])

# --------------------------------------------------
# PDF CREATOR FUNCTION (ReportLab)
# --------------------------------------------------
def create_pdf(student, pdf_path):
    c = canvas.Canvas(pdf_path, pagesize=LETTER)
    width, height = LETTER

    # draw logo if available
    if os.path.exists(logo_path):
        c.drawImage(logo_path, width/2 - 0.7*inch, height - 1.5*inch, 1.4*inch, 1.4*inch, preserveAspectRatio=True, mask='auto')

    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(width/2, height - 2.2*inch, "Dean’s List Certificate")

    c.setFont("Helvetica", 14)
    c.drawCentredString(width/2, height - 3.0*inch, f"Presented to {student.get('NAME','')}")

    c.setFont("Helvetica-Oblique", 12)
    term = student.get("TERM","").capitalize()
    year = student.get("YEAR","")
    gpa = student.get("GPA","")
    c.drawCentredString(width/2, height - 3.6*inch,
        f"For outstanding academic achievement with a GPA of {gpa} in the {term} term of {year}.")

    c.setFont("Helvetica", 11)
    c.drawCentredString(width/2, height - 4.6*inch, student.get("PROGRAM",""))

    # signatures
    c.setFont("Helvetica", 10)
    c.drawString(1.0*inch, 1.4*inch, "Joshua Rosenthal, Ed.D. – Assistant Dean")
    c.drawString(1.0*inch, 1.1*inch, "Graduate Business Programs")
    c.drawRightString(width - 1.0*inch, 1.4*inch, "Michelle Kim – Senior Director")
    c.drawRightString(width - 1.0*inch, 1.1*inch, "Online & MS Programs")

    c.setFont("Helvetica-Oblique", 9)
    c.drawCentredString(width/2, 0.7*inch, "Santa Clara University – Leavey School of Business")

    c.showPage()
    c.save()

# --------------------------------------------------
# MAIN LOGIC
# --------------------------------------------------
if uploaded_excel:
    if st.button("🚀 Generate PDFs"):
        with st.spinner("Creating personalized certificates... ⏳"):
            with tempfile.TemporaryDirectory() as td:
                excel_path = os.path.join(td, "students.xlsx")
                open(excel_path, "wb").write(uploaded_excel.read())

                df = pd.read_excel(excel_path, engine="openpyxl")
                output_pdfs = os.path.join(td, "pdfs")
                os.makedirs(output_pdfs, exist_ok=True)

                for _, row in df.iterrows():
                    pdf_path = os.path.join(output_pdfs, f"{row['NAME']}.pdf")
                    create_pdf(row.to_dict(), pdf_path)

                # zip results
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for file in os.listdir(output_pdfs):
                        zf.write(os.path.join(output_pdfs, file), arcname=file)
                zip_buffer.seek(0)

                st.success("✅ All PDF certificates generated successfully!")
                st.download_button(
                    label="⬇️ Download All Certificates (ZIP)",
                    data=zip_buffer,
                    file_name="DeansList_Certificates.zip",
                    mime="application/zip"
                )
else:
    st.info("Please upload the Excel sheet to begin.")