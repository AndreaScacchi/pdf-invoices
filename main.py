# import libraries
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# using the glob library
filepaths = glob.glob("invoices/*.xlsx")

# I've used the for loop to loop through all the file in the invoices folder and generated the same number of PDF files
for filepath in filepaths:
    
    # with the fpdf library I setted the orientation, unit and format of the files
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
    
    # here I've setted the family font, the size font and style for the number of all pdf files
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)
    
    # here instead I've done the same for the date
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)
    
    # here I've used the pandas library to read the files
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
    # add header
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=9, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)
    
    # add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=8)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)
    
    # this line generates the PDFs
    pdf.output(f"PDFs/{filename}.pdf")