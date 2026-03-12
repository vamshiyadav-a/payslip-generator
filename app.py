import streamlit as st
import pandas as pd
import os
import re
import zipfile
import subprocess
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from num2words import num2words

st.title("Salary Payslip Generator")

excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Word Template", type=["docx"])

if excel_file and template_file:

    if st.button("Generate Payslips"):

        os.makedirs("payslips", exist_ok=True)

        # Save uploaded files
        with open("data.xlsx", "wb") as f:
            f.write(excel_file.read())

        with open("template.docx", "wb") as f:
            f.write(template_file.read())

        # Extract Month from Excel title
        wb = load_workbook("data.xlsx")
        ws = wb.active

        title = str(ws["A1"].value)

        match = re.search(r"month of (.*)", title, re.IGNORECASE)
        month = match.group(1) if match else "Unknown"

        st.write("Detected Month:", month)

        # Read Excel
        df = pd.read_excel("data.xlsx", header=2)

        # Clean column names
        df.columns = df.columns.str.strip().str.upper()
        df = df.fillna(0)

        pdf_files = []

        for _, row in df.iterrows():

            doc = DocxTemplate("template.docx")

            gross = float(row.get("GROSS", 0))
            gross_dedns = float(row.get("GROSS DEDNS.", 0))
            total_deduction = float(row.get("TOTAL DEDUCTIONS", 0))
            net_pay = float(row.get("NET PAY", 0))

            net_words = num2words(int(net_pay), lang="en_IN").title()

            context = {
                "Month": month,
                "Employee_Name": row.get("NAME OF EMPLOYEE", ""),
                "EmpNo": row.get("EMP.NO.", ""),
                "Basic": row.get("BASIC", 0),
                "HRA": row.get("HRA", 0),
                "Spl_All": row.get("SPL. ALL.", 0),
                "LTA": row.get("LTA", 0),
                "PF": row.get("PF", 0),
                "PT": row.get("PT", 0),
                "Employer_Contribution": row.get("EMPLOYER CONTRIBUTION", 0),
                "GROSS": gross,
                "Gross_Dedns": gross_dedns,
                "Total_Deduction": total_deduction,
                "NetPay": net_pay,
                "NetPayWords": net_words
            }

            doc.render(context)

            emp = str(row.get("NAME OF EMPLOYEE", "Employee")).replace(" ", "_")

            temp_docx = f"{emp}.docx"
            doc.save(temp_docx)

            # Convert DOCX to PDF
            subprocess.run([
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                temp_docx,
                "--outdir",
                "payslips"
            ])

            generated_pdf = f"payslips/{emp}.pdf"
            final_pdf = f"payslips/payslip_{emp}_{month.replace(' ','_')}.pdf"

            if os.path.exists(generated_pdf):
                os.rename(generated_pdf, final_pdf)
                pdf_files.append(final_pdf)

            os.remove(temp_docx)

        st.success("PDF Payslips Generated")

        # Create ZIP file
        zip_name = "payslips_pdf_only.zip"

        with zipfile.ZipFile(zip_name, 'w') as z:
            for pdf in pdf_files:
                z.write(pdf, os.path.basename(pdf))

        with open(zip_name, "rb") as f:
            st.download_button(
                label="Download ZIP (PDF Payslips)",
                data=f,
                file_name="payslips_pdf_only.zip",
                mime="application/zip"
            )
