from fpdf import FPDF
import pandas as pd
from glob import glob
from pathlib import Path

files = glob("invoices/*.xlsx")

for file in files:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(file).stem
    inv_num, inv_date = filename.split("-")
    pdf_name = "invoice-" + inv_num

    pdf.set_font(family="Helvetica", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice #: {inv_num}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {inv_date}", ln=1)
    pdf.ln(h=8)

    df = pd.read_excel(file, "Sheet 1")

    # Add a header
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=65, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)
    pdf.set_text_color(80, 80, 80)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=65, h=8, txt=row["product_name"], border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Display total price
    total_sum = df["total_price"].sum()
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=65, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Add total sum sentence
    pdf.ln(h=8)
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)
    pdf.ln(h=8)

    # Add company name and logo
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=14, txt=f"CompanyName")
    pdf.image("logo.png", w=10)

    pdf.output(f"PDFs/{pdf_name}.pdf")
