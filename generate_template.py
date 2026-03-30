import pandas as pd

# Create sample data
data = {
    'SKU': ['RS24A17028-black-6', 'RS24A17028-black-7', 'RS24A17028-gold-6', 'Example-SKU-1'],
    'Availability': [15, 8, 5, 20]
}

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel
df.to_excel('Yesterday_Stock_Template.xlsx', index=False)
print("Template created successfully.")
