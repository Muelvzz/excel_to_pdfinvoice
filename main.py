import pandas as pd
import glob
import openpyxl
from fpdf import FPDF
from pathlib import Path

filepath = glob.glob("invoices/*.xlsx")

pdf = FPDF(orientation="P", unit="mm", format="A4")
pdf.set_auto_page_break(auto=False, margin=0)

for file in filepath:
    df = pd.read_excel(file, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")

    pdf.add_page()

    filename = Path(file).stem
    invoice_nr = filename.split("-")

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice Nr. {invoice_nr[0]}", align="L", ln=1)
    pdf.cell(w=0, h=5, txt=invoice_nr[1], align="L", ln=1)

    pdf.output(f"pdf's\{filename}.pdf")
