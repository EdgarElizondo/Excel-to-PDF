import glob
import pandas as pd
from fpdf import FPDF
from pathlib import Path

# Get EXCEL filepaths
filepaths = glob.glob("src/*xlsx")
for filepath in filepaths:
    # Load Excel into dataframes
    data = pd.read_excel(filepath, sheet_name="Sheet 1")
    # Creates PDF
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    document_num = filename.split("-")[0]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice num. {document_num}", align="L", ln=1)

    pdf.output(f"PDFs/{filename}.pdf")