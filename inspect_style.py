import openpyxl
import os

file_path = r"D:\Antigravity\Stock Cal\Yesterday_Stocks.xlsx"

print(f"--- Inspecting Styles in {os.path.basename(file_path)} ---")
try:
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    
    # Check Column B (index 'B')
    col_b = ws.column_dimensions['B']
    print(f"Column B Fill: {col_b.fill}")
    
    # Check a specific cell in Col B (e.g. B2)
    cell_b2 = ws['B2']
    print(f"Cell B2 Fill: {cell_b2.fill}")
    if cell_b2.fill and hasattr(cell_b2.fill, 'fgColor'):
        print(f"Cell B2 fgColor: {cell_b2.fill.fgColor.rgb}")
        
except Exception as e:
    print(f"Error: {e}")
