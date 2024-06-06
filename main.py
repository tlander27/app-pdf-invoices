from fpdf import FPDF
import pandas as pd
from glob import glob
from pathlib import Path

files = glob("invoices/*.xlsx")

for file in files:
    df = pd.read_excel(file, "Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Helvetica", size=16, style="B")
    filename = Path(file).stem
    inv_num, inv_date = filename.split("-")
    pdf_name = "invoice-" + inv_num
    pdf.cell(w=50, h=8, txt=f"Invoice #: {inv_num}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {inv_date}", ln=1)
    pdf.ln(h=8)
    pdf.set_font(family="Helvetica", size=12, style="B")

    for title in df.columns:
        pdf.cell(w=30, h=8, txt=title, border=1)

    pdf.output(f"PDFs/{pdf_name}.pdf")
