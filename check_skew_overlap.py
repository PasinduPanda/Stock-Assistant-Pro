import pandas as pd
import os

today_file = r"D:\Antigravity\Stock Cal\Today Stocks.xlsx"
yesterday_file = r"D:\Antigravity\Stock Cal\Yesterday_Stock_Template.xlsx"

try:
    df_today = pd.read_excel(today_file)
    df_yesterday = pd.read_excel(yesterday_file)

    today_skus = set(df_today['SKU'].dropna().astype(str))
    yesterday_skus = set(df_yesterday['SKU'].dropna().astype(str))

    intersection = today_skus.intersection(yesterday_skus)
    
    print(f"Total Today SKUs: {len(today_skus)}")
    print(f"Total Yesterday SKUs: {len(yesterday_skus)}")
    print(f"Intersection Count: {len(intersection)}")
    
    if len(intersection) > 0:
        print("Example Matches:", list(intersection)[:5])
    else:
        print("No matches found. Mapping definitely needed.")
        
except Exception as e:
    print(f"Error: {e}")
