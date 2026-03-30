import pandas as pd
import os

files = {
    "Today": r"D:\Antigravity\Stock Cal\Today_Stocks.xlsx",
    "Yesterday": r"D:\Antigravity\Stock Cal\Yesterday_Stocks.xlsx"
}

sku = "RS24A17070-apricot-9"

for label, file_path in files.items():
    print(f"\n--- Inspecting {label} ({os.path.basename(file_path)}) ---")
    try:
        df = pd.read_excel(file_path)
        print("Columns:", df.columns.tolist())
        
        # Normalize headers (strip)
        df.columns = df.columns.astype(str).str.strip()
        
        # Find SKU
        # Assuming header 'SKU' exists or column 0
        if 'SKU' in df.columns:
            row = df[df['SKU'].astype(str).str.strip() == sku]
        else:
            # Fallback to col 0
            row = df[df.iloc[:,0].astype(str).str.strip() == sku]
            
        if not row.empty:
            print(f"SKU {sku} Found:")
            print(row.to_string())
        else:
            print(f"SKU {sku} NOT FOUND.")
            
    except Exception as e:
        print(f"Error reading {label}: {e}")
