from fpdf import FPDF
import pandas as pd
from glob import glob
import os

files = glob("invoices/*.xlsx")

for file in files:
    df = pd.read_excel(file, "Sheet 1")
    print(df)