import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import os
import math

print("🎯 Creating DEMO Excel file...")

# Create demo data
stock_data = {
    'Product': ['Gold', 'Silver', 'Platinum', 'Diamond', 'TRY YOURS ➡️'],
    'Quantity': [100, 500, 50, 25, 0],
    'Cost Price ($)': [75000, 800, 25000, 12000, 0],
    'Selling Price ($)': [85000, 950, 30000, 15000, 0],
    'Total Cost ($)': [0, 0, 0, 0, 0],
    'Total Value ($)': [0, 0, 0, 0, 0],
    'Profit ($)': [0, 0, 0, 0, 0],
    'Margin %': [0, 0, 0, 0, 0]
}
df_stock = pd.DataFrame(stock_data)

# Calculate
for i in range(len(df_stock)):
    qty = df_stock.at[i, 'Quantity']
    cost = df_stock.at[i, 'Cost Price ($)']
    sell = df_stock.at[i, 'Selling Price ($)']
    df_stock.at[i, 'Total Cost ($)'] = qty * cost
    df_stock.at[i, 'Total Value ($)'] = qty * sell
    df_stock.at[i, 'Profit ($)'] = (qty * sell) - (qty * cost)
    if qty * cost > 0:
        df_stock.at[i, 'Margin %'] = round(((qty * sell) - (qty * cost)) / (qty * cost) * 100, 1)

# Percentage data
pct_data = {
    'Type': ['Discount', 'Tax', 'Tip', 'TRY YOURS ➡️'],
    'Amount ($)': [1000, 500, 150, 0],
    'Rate %': [15, 10, 18, 0],
    'Calculated': [0, 0, 0, 0],
    'Final': [0, 0, 0, 0]
}
df_pct = pd.DataFrame(pct_data)
for i in range(len(df_pct)):
    amt = df_pct.at[i, 'Amount ($)']
    rate = df_pct.at[i, 'Rate %']
    df_pct.at[i, 'Calculated'] = amt * (rate / 100)
    df_pct.at[i, 'Final'] = amt + df_pct.at[i, 'Calculated']

# Write to Excel
with pd.ExcelWriter('demo_calculator.xlsx', engine='openpyxl') as writer:
    df_stock.to_excel(writer, sheet_name='💰 STOCK DEMO', index=False)
    df_pct.to_excel(writer, sheet_name='🎯 PERCENTAGE DEMO', index=False)

print("✅ demo_calculator.xlsx created!")
print("📁 File ready for upload to GitHub!")
