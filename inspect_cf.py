import openpyxl
import os

file_path = r"D:\Antigravity\Stock Cal\Yesterday_Stocks.xlsx"

print(f"--- Inspecting Conditional Formatting in {os.path.basename(file_path)} ---")
try:
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    
    cf_rules = ws.conditional_formatting
    print(f"Conditional Formatting Rules Found: {len(cf_rules)}")
    for range_string, rules in cf_rules.items():
        print(f"Range: {range_string}")
        for rule in rules:
            print(f"  - Type: {rule.type}, Priority: {rule.priority}, DxfId: {rule.dxfId}")
            
except Exception as e:
    print(f"Error: {e}")
