import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# This extracts all the excel files and turns them into a list
filepath = glob.glob("invoices/*.xlsx")

# This iterates over the list in the filepath variable one by one
for file in filepath:

    # This creates a pdf with standard format shown
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False, margin=0)

    # Creates a page
    pdf.add_page()

    """
    This is the process to strip the unnecessary names included in the file, without this -  
    it would output into a invoices/10001-2023.1.18.xlsx and not 10001-2023.1.18.xlsx 
    """

    filename = Path(file).stem
    invoice_nr, date = filename.split("-")

    # This formats the header
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice Nr. {invoice_nr}", align="L", ln=1)
    pdf.cell(w=0, h=5, txt=f"Date {date}", align="L", ln=1)

    df = pd.read_excel(file, sheet_name="Sheet 1")

    # Creating columns
    columns = list(df.columns)
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times",style="B", size=10)

    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=50, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=40, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Creating rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)

        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Variable for calculating the total price
    total = df["total_price"].sum()

    # This creates a new row underneath the finished row after iterating in the for-loop
    pdf.set_font(family="Times", size=10)

    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total), border=1, ln=1)

    # Creating two lines under the cell
    pdf.set_font(family="Times", size=14, style="B")

    pdf.cell(w=50, h=8, txt=f"You're total sum is {total}", ln=1)
    pdf.cell(w=30, h=8, txt="By Muelvin Lopez")
    pdf.image("pythonhow.png", w=10)


    pdf.output(f"pdf's\{filename}.pdf")
