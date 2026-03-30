import pandas as pd

file = r"D:\Antigravity\Stock Cal\Today Stocks.xlsx"
sku = "RS24A17070-apricot-9"

print(f"--- Inspecting {file} ---")
df = pd.read_excel(file)
print("Columns:", df.columns.tolist())

# Find the row
row = df[df['SKU'].astype(str).str.strip() == sku]

if not row.empty:
    print("\nRow Data:")
    print(row.T)
else:
    print(f"\nSKU {sku} NOT FOUND in {file}")

print("\n--- Inspecting First 5 Rows ---")
print(df.head(5))
