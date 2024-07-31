import glob
import pandas as pd


filepaths = glob.glob("src/*xlsx")

for filepath in filepaths:
    data = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(data)