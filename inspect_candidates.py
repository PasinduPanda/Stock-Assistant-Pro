import pandas as pd
import os
from openpyxl import load_workbook

directory = r"D:\Antigravity\Stock Cal"

files = {
    "Candidate_MSKU (Sheet format)": os.path.join(directory, "Sheet format.xlsx"),
    "Candidate_Template (Yesterday_Stock_Template)": os.path.join(directory, "Yesterday_Stock_Template.xlsx"),
    "Candidate_Warehouse (Today Stocks)": os.path.join(directory, "Today Stocks.xlsx")
}

print("\n--- Inspecting Candidate Files ---")

for key, path in files.items():
    print(f"\nChecking {key}...")
    if not os.path.exists(path):
        print("  -> File NOT found!")
        continue
    
    try:
        # Load first few rows
        df = pd.read_excel(path, nrows=5, engine='openpyxl')
        print(f"  [Columns]: {df.columns.tolist()}")
        print(df.head(2).to_string())
    except Exception as e:
        print(f"  -> Error reading file: {e}")
