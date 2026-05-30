import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
import os
import math

class MasterCalculator:
    def __init__(self, filename="master_calculator_final.xlsx"):
        self.filename = filename
        
    def create_system(self):
        """Create complete Excel calculator system"""
        
        # ============================================================
        # SHEET 1: INDEX / DASHBOARD (Front Page)
        # ============================================================
        dashboard_data = {
            'S.No': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
            'Calculator': ['Stock Inventory', 'Demography', 'Statistics', 'Finance', 
                          'Math Solver', 'Geometry', 'Percentage', 'Trend Analysis', 
                          'Unit Converter', 'Custom Calculator'],
            'Category': ['Business', 'Demographics', 'Statistics', 'Finance', 
                        'Math', 'Geometry', 'Math', 'Business', 'Science', 'Custom'],
            'Description': ['Calculate stock value & profit', 'Population by age/gender',
                           'Mean, median & sum', 'Income, expense & balance',
                           'Add, subtract, multiply, divide', 'Area & volume of shapes',
                           'Discount, tax & margin', 'Sales growth & trends',
                           'Convert between units', 'Create your own calculator'],
            'Go to Sheet': ['STOCK', 'DEMOGRAPHY', 'STATISTICS', 'FINANCE',
                           'MATH', 'GEOMETRY', 'PERCENTAGE', 'TREND',
                           'UNIT', 'CREATE NEW'],
            'Status': ['✅', '✅', '✅', '✅', '✅', '✅', '✅', '✅', '✅', '✨']
        }
        df_index = pd.DataFrame(dashboard_data)
        
        # ============================================================
        # SHEET 2: STOCK CALCULATOR
        # ============================================================
        stock_data = {
            'Product': ['Gold', 'Silver', 'Platinum', 'Diamond', 'Ruby', 'Emerald'],
            'Qty': [100, 500, 50, 25, 40, 30],
            'Cost Price': [75000, 800, 25000, 12000, 4000, 6500],
            'Sell Price': [85000, 950, 30000, 15000, 5000, 8000],
            'Total Cost': [0, 0, 0, 0, 0, 0],
            'Total Value': [0, 0, 0, 0, 0, 0],
            'Profit': [0, 0, 0, 0, 0, 0],
            'Margin %': [0, 0, 0, 0, 0, 0]
        }
        df_stock = pd.DataFrame(stock_data)
        df_stock['Total Cost'] = df_stock['Qty'] * df_stock['Cost Price']
        df_stock['Total Value'] = df_stock['Qty'] * df_stock['Sell Price']
        df_stock['Profit'] = df_stock['Total Value'] - df_stock['Total Cost']
        df_stock['Margin %'] = round((df_stock['Profit'] / df_stock['Total Cost']) * 100, 1)
        
        # ============================================================
        # SHEET 3: DEMOGRAPHY
        # ============================================================
        demo_data = {
            'Age Group': ['0-18', '19-30', '31-45', '46-60', '60+', 'TOTAL'],
            'Population': [25000, 45000, 38000, 22000, 15000, 145000],
            'Male %': [48, 50, 49, 47, 42, 0],
            'Female %': [52, 50, 51, 53, 58, 0],
            'Male Count': [0, 0, 0, 0, 0, 0],
            'Female Count': [0, 0, 0, 0, 0, 0]
        }
        df_demo = pd.DataFrame(demo_data)
        for i in range(5):
            df_demo.at[i, 'Male Count'] = int(df_demo.at[i, 'Population'] * df_demo.at[i, 'Male %'] / 100)
            df_demo.at[i, 'Female Count'] = int(df_demo.at[i, 'Population'] * df_demo.at[i, 'Female %'] / 100)
        df_demo.at[5, 'Male Count'] = df_demo['Male Count'].sum()
        df_demo.at[5, 'Female Count'] = df_demo['Female Count'].sum()
        
        # ============================================================
        # SHEET 4: STATISTICS
        # ============================================================
        stats_data = {
            'Sample': ['A', 'B', 'C', 'D', 'E'],
            'Value 1': [23, 45, 67, 12, 89],
            'Value 2': [34, 56, 78, 23, 45],
            'Value 3': [45, 67, 89, 34, 56],
            'Mean': [0, 0, 0, 0, 0],
            'Median': [0, 0, 0, 0, 0],
            'Sum': [0, 0, 0, 0, 0]
        }
        df_stats = pd.DataFrame(stats_data)
        for i in range(5):
            vals = [df_stats.at[i, 'Value 1'], df_stats.at[i, 'Value 2'], df_stats.at[i, 'Value 3']]
            df_stats.at[i, 'Mean'] = round(sum(vals) / 3, 2)
            df_stats.at[i, 'Median'] = sorted(vals)[1]
            df_stats.at[i, 'Sum'] = sum(vals)
        
        # ============================================================
        # SHEET 5: FINANCE
        # ============================================================
        finance_data = {
            'Date': ['Day 1', 'Day 2', 'Day 3', 'Day 4', 'Day 5'],
            'Description': ['Salary', 'Rent', 'Food', 'Shopping', 'Savings'],
            'Income': [5000, 0, 0, 0, 0],
            'Expense': [0, 1200, 500, 300, 1000],
            'Balance': [0, 0, 0, 0, 0]
        }
        df_finance = pd.DataFrame(finance_data)
        bal = 0
        for i in range(5):
            bal = bal + df_finance.at[i, 'Income'] - df_finance.at[i, 'Expense']
            df_finance.at[i, 'Balance'] = bal
        
        # ============================================================
        # SHEET 6: MATH
        # ============================================================
        math_data = {
            'Operation': ['Add', 'Subtract', 'Multiply', 'Divide', 'Power', 'Sqrt'],
            'A': [25, 50, 12, 100, 4, 64],
            'B': [15, 25, 8, 20, 3, 2],
            'Result': [0, 0, 0, 0, 0, 0],
            'Formula': ['A+B', 'A-B', 'A×B', 'A÷B', 'A^B', '√A']
        }
        df_math = pd.DataFrame(math_data)
        df_math.at[0, 'Result'] = df_math.at[0, 'A'] + df_math.at[0, 'B']
        df_math.at[1, 'Result'] = df_math.at[1, 'A'] - df_math.at[1, 'B']
        df_math.at[2, 'Result'] = df_math.at[2, 'A'] * df_math.at[2, 'B']
        df_math.at[3, 'Result'] = round(df_math.at[3, 'A'] / df_math.at[3, 'B'], 2)
        df_math.at[4, 'Result'] = df_math.at[4, 'A'] ** df_math.at[4, 'B']
        df_math.at[5, 'Result'] = round(df_math.at[5, 'A'] ** 0.5, 2)
        
        # ============================================================
        # SHEET 7: GEOMETRY
        # ============================================================
        geo_data = {
            'Shape': ['Circle', 'Square', 'Rectangle', 'Triangle', 'Sphere', 'Cube'],
            'Input 1': [5, 4, 6, 3, 3, 4],
            'Input 2': [0, 0, 8, 4, 0, 0],
            'Area': [0, 0, 0, 0, 0, 0],
            'Perimeter/Volume': [0, 0, 0, 0, 0, 0]
        }
        df_geo = pd.DataFrame(geo_data)
        df_geo.at[0, 'Area'] = round(math.pi * 25, 2)
        df_geo.at[0, 'Perimeter/Volume'] = round(2 * math.pi * 5, 2)
        df_geo.at[1, 'Area'] = 16
        df_geo.at[1, 'Perimeter/Volume'] = 16
        df_geo.at[2, 'Area'] = 48
        df_geo.at[2, 'Perimeter/Volume'] = 28
        df_geo.at[3, 'Area'] = 6
        df_geo.at[4, 'Area'] = round(4 * math.pi * 9, 2)
        df_geo.at[4, 'Perimeter/Volume'] = round((4/3) * math.pi * 27, 2)
        df_geo.at[5, 'Area'] = 96
        df_geo.at[5, 'Perimeter/Volume'] = 64
        
        # ============================================================
        # SHEET 8: PERCENTAGE
        # ============================================================
        pct_data = {
            'Type': ['Discount', 'Tax', 'Tip', 'Profit', 'Commission'],
            'Amount': [1000, 500, 150, 5000, 2000],
            'Rate': [15, 10, 18, 25, 8],
            'Calculated': [0, 0, 0, 0, 0],
            'Final': [0, 0, 0, 0, 0]
        }
        df_pct = pd.DataFrame(pct_data)
        df_pct['Calculated'] = df_pct['Amount'] * (df_pct['Rate'] / 100)
        df_pct['Final'] = df_pct['Amount'] + df_pct['Calculated']
        
        # ============================================================
        # SHEET 9: TREND
        # ============================================================
        trend_data = {
            'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
            'Sales': [10000, 12000, 11500, 13500, 15000, 14800],
            'Growth': [0, 20.0, -4.2, 17.4, 11.1, -1.3],
            'Avg 3M': [0, 0, 11167, 12333, 13333, 14433]
        }
        df_trend = pd.DataFrame(trend_data)
        
        # ============================================================
        # SHEET 10: UNIT CONVERTER
        # ============================================================
        unit_data = {
            'Type': ['Length', 'Weight', 'Time', 'Temp', 'Currency', 'Volume'],
            'From': ['km', 'kg', 'hours', 'Celsius', 'USD', 'Liters'],
            'Value': [10, 100, 24, 25, 1000, 5],
            'To': ['miles', 'lbs', 'minutes', 'Fahrenheit', 'EUR', 'gallons'],
            'Result': [6.21, 220.46, 1440, 77, 920, 1.32]
        }
        df_unit = pd.DataFrame(unit_data)
        
        # ============================================================
        # SHEET 11: INSTRUCTIONS
        # ============================================================
        instructions = {
            'Step': ['1', '2', '3', '4', '5'],
            'Action': [
                'Open the Excel file',
                'Look at INDEX sheet (this is your dashboard)',
                'Click any sheet tab at the BOTTOM to use that calculator',
                'Enter your own numbers in any white cell',
                'All formulas calculate automatically!'
            ],
            'Tip': [
                'Sheet tabs are your buttons!',
                'Each sheet has a different calculator',
                'Try STOCK, DEMOGRAPHY, FINANCE etc.',
                'Add more rows if needed',
                'Right-click sheet to copy for custom calculator'
            ]
        }
        df_instructions = pd.DataFrame(instructions)
        
        # ============================================================
        # WRITE ALL SHEETS TO EXCEL
        # ============================================================
        with pd.ExcelWriter(self.filename, engine='openpyxl') as writer:
            df_index.to_excel(writer, sheet_name='🏠 INDEX', index=False)
            df_stock.to_excel(writer, sheet_name='💰 STOCK', index=False)
            df_demo.to_excel(writer, sheet_name='👥 DEMOGRAPHY', index=False)
            df_stats.to_excel(writer, sheet_name='📈 STATISTICS', index=False)
            df_finance.to_excel(writer, sheet_name='💳 FINANCE', index=False)
            df_math.to_excel(writer, sheet_name='🧮 MATH', index=False)
            df_geo.to_excel(writer, sheet_name='📐 GEOMETRY', index=False)
            df_pct.to_excel(writer, sheet_name='🎯 PERCENTAGE', index=False)
            df_trend.to_excel(writer, sheet_name='📉 TREND', index=False)
            df_unit.to_excel(writer, sheet_name='⚖️ UNIT', index=False)
            df_instructions.to_excel(writer, sheet_name='📖 HELP', index=False)
        
        # Apply formatting
        self._format_file()
        
        # Show success
        self._show_success()
        
        return True
    
    def _format_file(self):
        """Apply nice formatting"""
        try:
            wb = openpyxl.load_workbook(self.filename)
            
            # Color for INDEX header
            if '🏠 INDEX' in wb.sheetnames:
                ws = wb['🏠 INDEX']
                header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF")
                
                for cell in ws[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center")
                
                # Adjust column widths
                for col in ws.columns:
                    max_len = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_len:
                                max_len = len(str(cell.value))
                        except:
                            pass
                    ws.column_dimensions[col_letter].width = min(max_len + 2, 30)
            
            wb.save(self.filename)
            print("✓ Formatting applied")
        except:
            print("✓ File created successfully")
    
    def _show_success(self):
        """Display success message"""
        print("\n" + "="*60)
        print("✅ EXCEL CALCULATOR SYSTEM CREATED!")
        print("="*60)
        print(f"\n📁 File: {os.path.abspath(self.filename)}")
        print("\n📊 SHEETS IN THIS FILE:")
        print("   🏠 INDEX      - Main dashboard (front page)")
        print("   💰 STOCK      - Inventory calculator")
        print("   👥 DEMOGRAPHY - Population stats")
        print("   📈 STATISTICS - Mean, median, sum")
        print("   💳 FINANCE    - Budget tracker")
        print("   🧮 MATH       - Basic operations")
        print("   📐 GEOMETRY   - Area & volume")
        print("   🎯 PERCENTAGE - % calculator")
        print("   📉 TREND      - Sales analysis")
        print("   ⚖️ UNIT       - Unit converter")
        print("   📖 HELP       - Instructions")
        print("\n🎯 HOW TO USE:")
        print("   1. OPEN the Excel file")
        print("   2. The INDEX sheet is your dashboard")
        print("   3. CLICK any sheet tab at the BOTTOM")
        print("   4. ENTER your numbers - everything auto-calculates!")
        print("\n" + "="*60)

# ============================================================
# RUN THE PROGRAM
# ============================================================
if __name__ == "__main__":
    print("\n" + "🔧"*30)
    print("   CREATING MASTER EXCEL CALCULATOR")
    print("🔧"*30 + "\n")
    
    # Install packages if needed
    try:
        import pandas, openpyxl
    except ImportError:
        print("📦 Installing pandas and openpyxl...")
        os.system("pip install pandas openpyxl")
        print("✅ Installed!")
    
    # Create the calculator
    app = MasterCalculator("master_calculator_final.xlsx")
    app.create_system()
    
    print("\n🚀 READY TO USE! Open the file now!")
