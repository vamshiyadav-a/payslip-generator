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

        # Extract Month
        wb = load_workbook("data.xlsx")
        ws = wb.active
        title = str(ws["A1"].value)

        match = re.search(r"month of (.*)", title, re.IGNORECASE)
        month = match.group(1) if match else "Unknown"

        df = pd.read_excel("data.xlsx", header=2)

        pdf_files = []

        for _, row in df.iterrows():

            doc = DocxTemplate("template.docx")

            net_pay = row["Net Pay"]
            net_pay_words = num2words(net_pay, lang="en_IN").title()

            context = {
                "Month": month,
                "Employee_Name": row["Name of Employee"],
                "EmpNo": row["Emp.No."],
                "Basic": row["Basic"],
                "HRA": row["HRA"],
                "Spl_All": row["Spl. All."],
                "LTA": row["LTA"],
                "PF": row["PF"],
                "PT": row["PT"],
                "Employer_Contribution": row["Employer Contribution"],
                "GROSS": row["GROSS"],
                "Gross_Dedns": row["Gross Dedns."],
                "NetPay": net_pay,
                "NetPayWords": net_pay_words
            }

            doc.render(context)

            emp_name = str(row["Name of Employee"]).replace(" ", "_")

            temp_docx = f"{emp_name}.docx"
            doc.save(temp_docx)

            subprocess.run([
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                temp_docx,
                "--outdir",
                "payslips"
            ])

            generated_pdf = f"payslips/{emp_name}.pdf"
            final_pdf = f"payslips/payslip_{emp_name}_{month}.pdf"

            if os.path.exists(generated_pdf):
                os.rename(generated_pdf, final_pdf)
                pdf_files.append(final_pdf)

            os.remove(temp_docx)

        zip_name = "payslips.zip"

        with zipfile.ZipFile(zip_name, "w") as z:
            for pdf in pdf_files:
                z.write(pdf, os.path.basename(pdf))

        with open(zip_name, "rb") as f:
            st.download_button(
                "Download Payslips ZIP",
                f,
                file_name="payslips.zip"
            )

        st.success("Payslips Generated Successfully")
