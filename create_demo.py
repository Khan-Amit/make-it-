import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# Create Stock Demo Sheet
stock_data = {
    'Product': ['Gold', 'Silver', 'TRY YOURSELF'],
    'Quantity': [100, 500, 0],
    'Cost Price ($)': [75000, 800, 0],
    'Selling Price ($)': [85000, 950, 0],
    'Total Cost ($)': ['=B2*C2', '=B3*C3', '=B4*C4'],
    'Total Value ($)': ['=B2*D2', '=B3*D3', '=B4*D4'],
    'Profit ($)': ['=F2-E2', '=F3-E3', '=F4-E4'],
    'Margin %': ['=G2/E2*100', '=G3/E3*100', '=G4/E4*100']
}

df_stock = pd.DataFrame(stock_data)

# Create Percentage Demo Sheet
pct_data = {
    'Type': ['Discount', 'Sales Tax', 'TRY YOURSELF'],
    'Original Amount ($)': [1000, 500, 0],
    'Rate (%)': [15, 10, 0],
    'Calculated Amount ($)': ['=B2*C2/100', '=B3*C3/100', '=B4*C4/100'],
    'Final Amount ($)': ['=B2+D2', '=B3+D3', '=B4+D4']
}

df_pct = pd.DataFrame(pct_data)

# Create Math Demo Sheet
math_data = {
    'Operation': ['Addition', 'Subtraction', 'Multiplication', 'Division', 'TRY YOURSELF'],
    'Number A': [25, 50, 12, 100, 0],
    'Number B': [15, 25, 8, 20, 0],
    'Result': ['=B2+C2', '=B3-C3', '=B4*C4', '=B5/C5', '=B6+C6']
}

df_math = pd.DataFrame(math_data)

# Write to Excel
with pd.ExcelWriter('demo_calculator.xlsx', engine='openpyxl') as writer:
    df_stock.to_excel(writer, sheet_name='💰 STOCK DEMO', index=False)
    df_pct.to_excel(writer, sheet_name='🎯 PERCENTAGE DEMO', index=False)
    df_math.to_excel(writer, sheet_name='🧮 MATH DEMO', index=False)

print("✅ demo_calculator.xlsx CREATED SUCCESSFULLY!")
print("📁 Location: Same folder as this script")
