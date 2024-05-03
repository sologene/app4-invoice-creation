import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="p",unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")
    filename = Path(filepath).stem
    invoice_name=filename.split("-")[0]
    pdf.cell(w=50,h=8, txt=f"Invoice nr.{invoice_name}")
    pdf.output(f"Output/{filename}.pdf")