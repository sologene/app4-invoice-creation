import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)
for filepath in filepaths:
    pdf = FPDF(orientation="p",unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False, margin=0)
    pdf.add_page()


    filename = Path(filepath).stem
    invoice_name, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50,h=8, txt=f"Invoice nr.{invoice_name}", ln=1)

    pdf.set_font(family="Times", size=15, style="B")
    pdf.cell(w=50, h=8, txt=f"Date:{date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf.set_font(family="Times", size=10, style="B")
    col = list(df.columns)
    col = [item.replace("_", " ") for item in col]
    pdf.cell(w=30, h=8, txt=col[0], border=1, align="C")
    pdf.cell(w=50, h=8, txt=col[1], border=1, align="C")
    pdf.cell(w=30, h=8, txt=col[2], border=1, align="C")
    pdf.cell(w=50, h=8, txt=col[3], border=1, align="C")
    pdf.cell(w=50, h=8, txt=col[4], border=1, align="C", ln=1)

    for index,item in df.iterrows():
        pdf.set_font(family = "Times", size = 10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(item["product_id"]), border=1, align="C")
        pdf.cell(w=50, h=8, txt=str(item["product_name"]),border=1, align="C")
        pdf.cell(w=30, h=8, txt=str(item["amount_purchased"]),border=1, align="C")
        pdf.cell(w=50, h=8, txt=str(item["price_per_unit"]),border=1, align="C")
        pdf.cell(w=50, h=8, txt=str(item["total_price"]),border=1, align="C",ln=1)
    total = str(df["total_price"].sum())


    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1, align="C")
    pdf.cell(w=50, h=8, txt="", border=1, align="C")
    pdf.cell(w=30, h=8, txt="", border=1, align="C")
    pdf.cell(w=50, h=8, txt="", border=1, align="C")
    pdf.cell(w=50, h=8, txt=total, border=1, align="C", ln=1)

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=30, h=8, txt=f"The total price is {total}", align="L", ln=1)
    pdf.image(r"C:\Users\Sologene\Downloads\download (1).jpg",w=10)
    pdf.output(f"Output/{filename}.pdf")