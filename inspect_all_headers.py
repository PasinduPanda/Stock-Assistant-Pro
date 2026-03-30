import openpyxl
import os
import glob

# Search in root and check dir
files = glob.glob(r"D:\Antigravity\Stock Cal\*.xlsx") + glob.glob(r"D:\Antigravity\Stock Cal\check\*.xlsx")

print(f"Found {len(files)} files to inspect.\n")

for path in files:
    print(f"--- FIlE: {os.path.basename(path)} ---")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active # Inspect first sheet
        
        # Print first 10 rows to see layout
        found_data = False
        for i, row in enumerate(ws.iter_rows(max_row=10, values_only=True)):
            # Filter out completely empty rows for cleaner output
            if any(row):
                print(f"Row {i+1}: {row}")
                found_data = True
        
        if not found_data:
            print("  (Empty or blank sheet)")
            
    except Exception as e:
        print(f"  Error reading file: {e}")
    print("\n")
