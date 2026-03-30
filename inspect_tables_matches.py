import openpyxl

path = r"D:\Antigravity\Stock Cal\check\My\Yesterday_Stocks_Template.xlsx"
print(f"Inspecting Deeply: {path}")

try:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    
    # 1. Check Tables
    if ws.tables:
        print(f"Tables Found: {len(ws.tables)}")
        for name, table in ws.tables.items():
            print(f"  Table: {name} Range: {table.ref} Style: {table.tableStyleInfo}")
    else:
        print("No Excel Tables found.")

    # 2. Deep Scan Column B for ANY fill
    print("\nScanning Column B for any fill...")
    found_fill = False
    for i, row in enumerate(ws.iter_rows(min_col=2, max_col=2)):
        cell = row[0]
        if cell.fill and (cell.fill.patternType or cell.fill.fgColor):
            # Check if it's "real" fill (not default None/00000000)
            fg = cell.fill.fgColor.rgb if cell.fill.fgColor else "None"
            if fg != "00000000" and fg != "None":
                print(f"  Row {i+1}: Fill Found! FG={fg}")
                found_fill = True
                if i > 20: break # Stop after found

    if not found_fill:
        print("No explicit fill found in first 20+ rows.")

except Exception as e:
    print(f"Error: {e}")
