import pandas as pd
import os

files = [
    r"D:\Antigravity\Stock Cal\MSKU Sheets.xlsx",
    r"D:\Antigravity\Stock Cal\Tiktok Template.xlsx",
    r"D:\Antigravity\Stock Cal\Warehouse sheet.xlsx"
]

for file in files:
    print(f"--- {os.path.basename(file)} ---")
    try:
        df = pd.read_excel(file, nrows=5)
        print(df.columns.tolist())
        print(df.head(2))
    except Exception as e:
        print(f"Error reading {file}: {e}")
    print("\n")
