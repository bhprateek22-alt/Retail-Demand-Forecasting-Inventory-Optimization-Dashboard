import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# 1. Generate Historical Sales Data (Fact_Sales)
np.random.seed(42)
start_date = datetime.now().date() - timedelta(days=90)
dates = [start_date + timedelta(days=i) for i in range(90)]

skus = ['SKU-101 (T-Shirt)', 'SKU-202 (Headphones)', 'SKU-303 (Shoes)']
sales_data = []

for sku in skus:
    base = 50 if '101' in sku else (20 if '202' in sku else 35)
    for d in dates:
        sales = max(5, int(base + np.random.normal(0, 8)))
        sales_data.append({'Date': d, 'SKU': sku, 'Sales_Units': sales})

df_sales = pd.DataFrame(sales_data)

# 2. Generate Editable Parameters (Dim_ProductParameters)
df_params = pd.DataFrame([
    {'SKU': 'SKU-101 (T-Shirt)', 'Current_Stock': 250, 'Lead_Time_Days': 7, 'Service_Level': 0.95, 'Ordering_Cost': 50, 'Holding_Cost': 2.50},
    {'SKU': 'SKU-202 (Headphones)', 'Current_Stock': 80, 'Lead_Time_Days': 14, 'Service_Level': 0.98, 'Ordering_Cost': 120, 'Holding_Cost': 8.00},
    {'SKU': 'SKU-303 (Shoes)', 'Current_Stock': 190, 'Lead_Time_Days': 10, 'Service_Level': 0.95, 'Ordering_Cost': 80, 'Holding_Cost': 4.50}
])

# Write to Excel formatted as Official Tables
with pd.ExcelWriter('data_source.xlsx', engine='openpyxl') as writer:
    df_sales.to_excel(writer, sheet_name='Sales_Data', index=False)
    df_params.to_excel(writer, sheet_name='Parameters', index=False)

print("Created 'data_source.xlsx' successfully!")
