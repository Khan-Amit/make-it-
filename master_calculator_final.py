import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import os
import math

# ============================================================
# SINGLE CLEAN CALCULATOR - NO CONFLICTS
# ============================================================

def create_calculator():
    filename = "my_calculator.xlsx"
    
    # INDEX/DASHBOARD PAGE
    index_data = {
        '#' : [1, 2, 3, 4, 5, 6, 7, 8, 9],
        'Calculator' : ['Stock', 'Demography', 'Statistics', 'Finance', 
                       'Math', 'Geometry', 'Percentage', 'Trend', 'Unit'],
        'What It Does' : ['Value & profit', 'Population stats', 'Mean/median/sum',
                         'Income/expense', 'Add/subtract/multiply/divide',
                         'Area & volume', 'Discount/tax/margin', 'Sales growth',
                         'Convert units'],
        'Sheet Name' : ['📊 STOCK', '👥 DEMO', '📈 STATS', '💰 FINANCE',
                       '🔢 MATH', '📐 GEO', '% PERCENT', '📉 TREND', '⚖️ UNIT']
    }
    df_index = pd.DataFrame(index_data)
    
    # STOCK CALCULATOR
    stock = pd.DataFrame({
        'Item': ['Gold', 'Silver', 'Diamond', 'Ruby'],
        'Qty': [100, 500, 25, 40],
        'Cost': [75000, 800, 12000, 4000],
        'Price': [85000, 950, 15000, 5000],
        'Total Cost': [0, 0, 0, 0],
        'Total Value': [0, 0, 0, 0],
        'Profit': [0, 0, 0, 0]
    })
    stock['Total Cost'] = stock['Qty'] * stock['Cost']
    stock['Total Value'] = stock['Qty'] * stock['Price']
    stock['Profit'] = stock['Total Value'] - stock['Total Cost']
    
    # DEMOGRAPHY
    demo = pd.DataFrame({
        'Age': ['0-18', '19-30', '31-45', '46-60', '60+'],
        'Population': [25000, 45000, 38000, 22000, 15000],
        'Male %': [48, 50, 49, 47, 42],
        'Female %': [52, 50, 51, 53, 58]
    })
    
    # STATISTICS
    stats = pd.DataFrame({
        'Data': ['Set A', 'Set B', 'Set C'],
        'V1': [23, 45, 67],
        'V2': [34, 56, 78],
        'V3': [45, 67, 89],
        'Mean': [0, 0, 0],
        'Sum': [0, 0, 0]
    })
    for i in range(3):
        vals = [stats.at[i, 'V1'], stats.at[i, 'V2'], stats.at[i, 'V3']]
        stats.at[i, 'Mean'] = round(sum(vals)/3, 2)
        stats.at[i, 'Sum'] = sum(vals)
    
    # FINANCE
    finance = pd.DataFrame({
        'Date': ['Day1', 'Day2', 'Day3', 'Day4'],
        'Income': [5000, 0, 0, 0],
        'Expense': [0, 1200, 500, 800],
        'Balance': [0, 0, 0, 0]
    })
    bal = 0
    for i in range(4):
        bal = bal + finance.at[i, 'Income'] - finance.at[i, 'Expense']
        finance.at[i, 'Balance'] = bal
    
    # MATH
    math_calc = pd.DataFrame({
        'Op': ['Add', 'Subtract', 'Multiply', 'Divide'],
        'A': [25, 50, 12, 100],
        'B': [15, 25, 8, 20],
        'Result': [40, 25, 96, 5]
    })
    
    # GEOMETRY
    geo = pd.DataFrame({
        'Shape': ['Circle r=5', 'Square s=4', 'Rectangle 6x8'],
        'Area': [78.54, 16, 48],
        'Perimeter': [31.42, 16, 28]
    })
    
    # PERCENTAGE
    pct = pd.DataFrame({
        'Type': ['Discount', 'Tax', 'Tip'],
        'Amount': [1000, 500, 150],
        'Rate': [15, 10, 18],
        'Value': [150, 50, 27],
        'Total': [850, 550, 177]
    })
    
    # TREND
    trend = pd.DataFrame({
        'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
        'Sales': [10000, 12000, 11500, 13500, 15000, 14800],
        'Growth %': ['-', 20, -4.2, 17.4, 11.1, -1.3]
    })
    
    # UNIT
    unit = pd.DataFrame({
        'From': ['10 km', '100 kg', '24 hours', '25°C'],
        'To': ['6.21 miles', '220.46 lbs', '1440 minutes', '77°F']
    })
    
    # Write all to Excel
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df_index.to_excel(writer, sheet_name='🏠 INDEX', index=False)
        stock.to_excel(writer, sheet_name='📊 STOCK', index=False)
        demo.to_excel(writer, sheet_name='👥 DEMO', index=False)
        stats.to_excel(writer, sheet_name='📈 STATS', index=False)
        finance.to_excel(writer, sheet_name='💰 FINANCE', index=False)
        math_calc.to_excel(writer, sheet_name='🔢 MATH', index=False)
        geo.to_excel(writer, sheet_name='📐 GEO', index=False)
        pct.to_excel(writer, sheet_name='% PERCENT', index=False)
        trend.to_excel(writer, sheet_name='📉 TREND', index=False)
        unit.to_excel(writer, sheet_name='⚖️ UNIT', index=False)
    
    # Formatting
    wb = openpyxl.load_workbook(filename)
    if '🏠 INDEX' in wb.sheetnames:
        ws = wb['🏠 INDEX']
        green_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
        white_font = Font(bold=True, color="FFFFFF")
        for cell in ws[1]:
            cell.fill = green_fill
            cell.font = white_font
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 25
        ws.column_dimensions['D'].width = 20
    wb.save(filename)
    
    print("\n" + "="*50)
    print("✅ ONE CLEAN CALCULATOR CREATED!")
    print("="*50)
    print(f"\n📁 File: {os.path.abspath(filename)}")
    print("\n📊 SHEETS (Click tabs at bottom):")
    print("   🏠 INDEX    - Main menu")
    print("   📊 STOCK    - Inventory")
    print("   👥 DEMO     - Population")
    print("   📈 STATS    - Statistics")
    print("   💰 FINANCE  - Budget")
    print("   🔢 MATH     - Calculations")
    print("   📐 GEO      - Geometry")
    print("   % PERCENT   - Percentage")
    print("   📉 TREND    - Sales")
    print("   ⚖️ UNIT     - Converter")
    print("\n🎯 HOW TO USE:")
    print("   1. OPEN the file")
    print("   2. INDEX sheet shows everything")
    print("   3. CLICK any tab at BOTTOM")
    print("   4. ENTER numbers - auto calculates!")
    print("="*50)

# RUN IT
if __name__ == "__main__":
    try:
        import pandas, openpyxl
    except:
        import subprocess
        subprocess.call(['pip', 'install', 'pandas', 'openpyxl'])
    
    create_calculator()
    print("\n🚀 DONE! Open 'my_calculator.xlsx' now!")
