# import libraries
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# using the glob library
filepaths = glob.glob("invoices/*.xlsx")

# I've used the for loop to loop through all the file in the invoices folder and generated the same number of PDF files
for filepath in filepaths:
    # here I've used the pandas library to read the files
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
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
    pdf.cell(w=50, h=8, txt=f"Date: {date}")
    
    # this line generates the PDFs
    pdf.output(f"PDFs/{filename}.pdf")