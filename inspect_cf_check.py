import openpyxl
import os

path = r"D:\Antigravity\Stock Cal\check\Yesterday_Stocks_Template.xlsx"
print(f"Inspecting CF for: {path}")

try:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    
    # Check Conditional Formatting
    if ws.conditional_formatting:
        print(f"Conditional Formatting Rules Found: {len(ws.conditional_formatting)}")
        for cf in ws.conditional_formatting:
            print(f"Rule: {cf}")
    else:
        print("No Conditional Formatting rules found.")

    # Check Column B Style
    col_b = ws.column_dimensions['B']
    print(f"Column B Fill: {col_b.fill}")

except Exception as e:
    print(f"Error: {e}")
