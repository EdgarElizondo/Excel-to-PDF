import glob
import pandas as pd
from fpdf import FPDF
from pathlib import Path

# Get EXCEL filepaths
filepaths = glob.glob("src/Data/*xlsx")

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
    columns_width = [30,70,30,30,30]
    # Header
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(0, 0, 0)    
    for index, column in enumerate(data.columns):
        col_name = column.replace("_"," ").title()
        pdf.cell(w=columns_width[index], h=10, txt=f"{col_name}", border=1)
    pdf.cell(w=50, h=10,txt="",ln=1)

    # Values
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(50, 50, 50)
    for index, row in data.iterrows():
        for index, value in enumerate(row.values):
            pdf.cell(w=columns_width[index], h=10, txt=f"{value}", border=1)
        pdf.cell(w=50, h=10,txt="",ln=1)
    
    # Total Values
    total_sum = data[column].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(50, 50, 50)    
    for index, column in enumerate(data.columns):
        if column == "total_price":
            pdf.cell(w=columns_width[index], h=10, txt=f"{total_sum}", border=1)
        else:
            pdf.cell(w=columns_width[index], h=10, txt="", border=1)
    pdf.cell(w=50, h=10,txt="",ln=1)

    # Total sum sentences
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"Total price is {total_sum}", ln=1)
    # Campany Name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=35, h=8, txt=f"2-Ring Python")
    pdf.image("src/Images/pythonlogo.png", w=10)
    
    pdf.output(f"PDFs/{filename}.pdf") 