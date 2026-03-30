import openpyxl

path = r"D:\Antigravity\Stock Cal\check\Updated_Yesterday_Stocks_Template.xlsx"
print(f"Inspecting Output File: {path}")

try:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    
    print("\n--- Inspecting Column B Cells (Output) ---")
    for i, row in enumerate(ws.iter_rows(min_row=2, max_row=6, min_col=2, max_col=2)):
        cell = row[0]
        fill = cell.fill
        fg = fill.fgColor.rgb if fill.fgColor else 'None'
        print(f"Cell {cell.coordinate}: Pattern={fill.patternType}, FG={fg}")
        
        # Check if it matches Yellow (FFFFFF00)
        if fg == "FFFFFF00":
            print("  -> DETECTED YELLOW!")

except Exception as e:
    print(f"Error: {e}")
