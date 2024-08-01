import glob
import pandas as pd
from fpdf import FPDF
from pathlib import Path

# Get EXCEL filepaths
filepaths = glob.glob("src/*xlsx")

for filepath in filepaths:
    # Creates PDF
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    # Get file name
    filename = Path(filepath).stem
    document_num, date = filename.split("-")
    # Document number
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice num. {document_num}", align="L", ln=1)
    # Date
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", align="L", ln=1)
    
    # Load Excel into dataframes
    data = pd.read_excel(filepath, sheet_name="Sheet 1")
    # Table
    columns_width = [30,50,50,30,30]
    # Header
    for index, column in enumerate(data.columns):
        pdf.set_font(family="Times", size=12, style="B")
        pdf.set_text_color(0, 0, 0)    
        col_name = column.replace("_"," ").title()
        pdf.cell(w=columns_width[index], h=10, txt=f"{col_name}", border=1)
    pdf.cell(w=50, h=10,txt="",ln=1)

    # Values
    for index, row in data.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(50, 50, 50)
        for index, value in enumerate(row.values):
            pdf.cell(w=columns_width[index], h=10, txt=f"{value}", border=1)
        pdf.cell(w=50, h=10,txt="",ln=1)


    pdf.output(f"PDFs/{filename}.pdf")