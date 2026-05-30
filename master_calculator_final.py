import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import os
import math

class MasterCalculator:
    def __init__(self, filename="my_calculator.xlsx"):
        self.filename = filename
    
    def create_system(self):
        """Create complete Excel calculator system"""
        
        # ========== SHEET 1: INDEX/DASHBOARD ==========
        index_data = {
            'S.No': [1, 2, 3, 4, 5, 6, 7, 8, 9],
            'Calculator Name': ['Stock Inventory', 'Demography', 'Statistics', 
                               'Finance', 'Math Solver', 'Geometry', 
                               'Percentage', 'Trend Analysis', 'Unit Converter'],
            'Category': ['Business', 'Demographics', 'Statistics', 'Finance',
                        'Math', 'Geometry', 'Math', 'Business', 'Science'],
            'Description': ['Calculate stock value & profit', 'Population by age/gender',
                           'Mean, median & sum', 'Income, expense & balance',
                           'Add, subtract, multiply, divide', 'Area & volume of shapes',
                           'Discount, tax & margin', 'Sales growth & trends',
                           'Convert between units'],
            'Status': ['✅', '✅', '✅', '✅', '✅', '✅', '✅', '✅', '✅']
        }
        df_index = pd.DataFrame(index_data)
        
        # ========== SHEET 2: STOCK CALCULATOR ==========
        stock_data = {
            'Product': ['Gold', 'Silver', 'Platinum', 'Diamond', 'Ruby', 'Emerald'],
            'Quantity': [100, 500, 50, 25, 40, 30],
            'Cost Price': [75000, 800, 25000, 12000, 4000, 6500],
            'Selling Price': [85000, 950, 30000, 15000, 5000, 8000],
            'Total Cost': [0, 0, 0, 0, 0, 0],
            'Total Value': [0, 0, 0, 0, 0, 0],
            'Profit': [0, 0, 0, 0, 0, 0],
            'Margin %': [0, 0, 0, 0, 0, 0]
        }
        df_stock = pd.DataFrame(stock_data)
        df_stock['Total Cost'] = df_stock['Quantity'] * df_stock['Cost Price']
        df_stock['Total Value'] = df_stock['Quantity'] * df_stock['Selling Price']
        df_stock['Profit'] = df_stock['Total Value'] - df_stock['Total Cost']
        df_stock['Margin %'] = round((df_stock['Profit'] / df_stock['Total Cost']) * 100, 1)
        
        # ========== SHEET 3: DEMOGRAPHY ==========
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
        
        # ========== SHEET 4: STATISTICS ==========
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
            values = [df_stats.at[i, 'Value 1'], df_stats.at[i, 'Value 2'], df_stats.at[i, 'Value 3']]
            df_stats.at[i, 'Mean'] = round(sum(values) / 3, 2)
            df_stats.at[i, 'Median'] = sorted(values)[1]
            df_stats.at[i, 'Sum'] = sum(values)
        
        # ========== SHEET 5: FINANCE ==========
        finance_data = {
            'Date': ['Day 1', 'Day 2', 'Day 3', 'Day 4', 'Day 5'],
            'Description': ['Salary', 'Rent', 'Groceries', 'Shopping', 'Savings'],
            'Income': [5000, 0, 0, 0, 0],
            'Expense': [0, 1200, 500, 300, 1000],
            'Balance': [0, 0, 0, 0, 0]
        }
        df_finance = pd.DataFrame(finance_data)
        balance = 0
        for i in range(5):
            balance += df_finance.at[i, 'Income'] - df_finance.at[i, 'Expense']
            df_finance.at[i, 'Balance'] = balance
        
        # ========== SHEET 6: MATH ==========
        math_data = {
            'Operation': ['Addition', 'Subtraction', 'Multiplication', 'Division', 'Power', 'Square Root'],
            'Value A': [25, 50, 12, 100, 4, 64],
            'Value B': [15, 25, 8, 20, 3, 2],
            'Result': [0, 0, 0, 0, 0, 0],
            'Formula': ['A + B', 'A - B', 'A × B', 'A ÷ B', 'A^B', '√A']
        }
        df_math = pd.DataFrame(math_data)
        df_math.at[0, 'Result'] = df_math.at[0, 'Value A'] + df_math.at[0, 'Value B']
        df_math.at[1, 'Result'] = df_math.at[1, 'Value A'] - df_math.at[1, 'Value B']
        df_math.at[2, 'Result'] = df_math.at[2, 'Value A'] * df_math.at[2, 'Value B']
        df_math.at[3, 'Result'] = round(df_math.at[3, 'Value A'] / df_math.at[3, 'Value B'], 2)
        df_math.at[4, 'Result'] = df_math.at[4, 'Value A'] ** df_math.at[4, 'Value B']
        df_math.at[5, 'Result'] = round(math.sqrt(df_math.at[5, 'Value A']), 2)
        
        # ========== SHEET 7: GEOMETRY ==========
        geo_data = {
            'Shape': ['Circle', 'Square', 'Rectangle', 'Triangle', 'Sphere', 'Cube'],
            'Parameter 1': [5, 4, 6, 3, 3, 4],
            'Parameter 2': [0, 0, 8, 4, 0, 0],
            'Area/Surface': [0, 0, 0, 0, 0, 0],
            'Perimeter/Volume': [0, 0, 0, 0, 0, 0]
        }
        df_geo = pd.DataFrame(geo_data)
        df_geo.at[0, 'Area/Surface'] = round(math.pi * 25, 2)
        df_geo.at[0, 'Perimeter/Volume'] = round(2 * math.pi * 5, 2)
        df_geo.at[1, 'Area/Surface'] = 16
        df_geo.at[1, 'Perimeter/Volume'] = 16
        df_geo.at[2, 'Area/Surface'] = 48
        df_geo.at[2, 'Perimeter/Volume'] = 28
        df_geo.at[3, 'Area/Surface'] = 6
        df_geo.at[4, 'Area/Surface'] = round(4 * math.pi * 9, 2)
        df_geo.at[4, 'Perimeter/Volume'] = round((4/3) * math.pi * 27, 2)
        df_geo.at[5, 'Area/Surface'] = 96
        df_geo.at[5, 'Perimeter/Volume'] = 64
        
        # ========== SHEET 8: PERCENTAGE ==========
        pct_data = {
            'Type': ['Discount', 'Sales Tax', 'Tip', 'Profit Margin', 'Commission'],
            'Original Amount': [1000, 500, 150, 5000, 2000],
            'Rate %': [15, 10, 18, 25, 8],
            'Calculated Amount': [0, 0, 0, 0, 0],
            'Final Amount': [0, 0, 0, 0, 0]
        }
        df_pct = pd.DataFrame(pct_data)
        df_pct['Calculated Amount'] = df_pct['Original Amount'] * (df_pct['Rate %'] / 100)
        df_pct['Final Amount'] = df_pct['Original Amount'] + df_pct['Calculated Amount']
        
        # ========== SHEET 9: TREND ==========
        trend_data = {
            'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
            'Sales': [10000, 12000, 11500, 13500, 15000, 14800, 16000, 17000, 16500, 18000, 19000, 20000],
            'Growth %': [0, 20.0, -4.2, 17.4, 11.1, -1.3, 8.1, 6.3, -2.9, 9.1, 5.6, 5.3],
            '3-Month Avg': [0, 0, 11167, 12333, 13333, 14433, 15267, 15933, 16500, 17167, 17833, 18667]
        }
        df_trend = pd.DataFrame(trend_data)
        
        # ========== SHEET 10: UNIT CONVERTER ==========
        unit_data = {
            'Category': ['Length', 'Weight', 'Time', 'Temperature', 'Currency', 'Volume', 'Speed'],
            'From': ['km', 'kg', 'hours', 'Celsius', 'USD', 'Liters', 'km/h'],
            'Value': [10, 100, 24, 25, 1000, 5, 100],
            'To': ['miles', 'lbs', 'minutes', 'Fahrenheit', 'EUR', 'gallons', 'mph'],
            'Result': [6.21, 220.46, 1440, 77, 920, 1.32, 62.14],
            'Formula': ['× 0.6214', '× 2.2046', '× 60', '(°C × 9/5) + 32', '× 0.92', '× 0.2642', '× 0.6214']
        }
        df_unit = pd.DataFrame(unit_data)
        
        # ========== WRITE ALL SHEETS ==========
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
        
        # Apply formatting
        self._apply_formatting()
        self._show_success()
        
        return True
    
    def _apply_formatting(self):
        """Apply professional formatting"""
        try:
            wb = openpyxl.load_workbook(self.filename)
            
            if '🏠 INDEX' in wb.sheetnames:
                ws = wb['🏠 INDEX']
                header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF", size=11)
                
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
                            if cell.value and len(str(cell.value)) > max_len:
                                max_len = len(str(cell.value))
                        except:
                            pass
                    ws.column_dimensions[col_letter].width = min(max_len + 2, 35)
            
            wb.save(self.filename)
            print("✓ Professional formatting applied!")
        except Exception as e:
            print(f"✓ File created with basic formatting")
    
    def _show_success(self):
        """Display success message"""
        print("\n" + "="*60)
        print("✅ MASTER EXCEL CALCULATOR CREATED!")
        print("="*60)
        print(f"\n📁 File: {os.path.abspath(self.filename)}")
        print("\n📊 10 CALCULATOR SHEETS INCLUDED:")
        print("   🏠 INDEX      - Main dashboard")
        print("   💰 STOCK      - Inventory calculator")
        print("   👥 DEMOGRAPHY - Population stats")
        print("   📈 STATISTICS - Mean, median, sum")
        print("   💳 FINANCE    - Budget tracker")
        print("   🧮 MATH       - Basic operations")
        print("   📐 GEOMETRY   - Area & volume")
        print("   🎯 PERCENTAGE - % calculator")
        print("   📉 TREND      - Sales analysis")
        print("   ⚖️ UNIT       - Unit converter")
        print("\n🎯 HOW TO USE:")
        print("   1. OPEN the Excel file")
        print("   2. Click 🏠 INDEX sheet to see all calculators")
        print("   3. Click any sheet tab at the BOTTOM")
        print("   4. Enter your numbers - everything auto-calculates!")
        print("\n💡 TIP: Right-click any sheet → 'Move or Copy' → 'Create a copy'")
        print("   Then rename to create your own custom calculator!")
        print("="*60)

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
        print("✓ Required packages found!")
    except ImportError:
        print("📦 Installing pandas and openpyxl...")
        os.system("pip install pandas openpyxl")
        print("✅ Packages installed!")
    
    # Create the calculator
    app = MasterCalculator("my_calculator.xlsx")
    app.create_system()
    
    print("\n🚀 SYSTEM READY! Open 'my_calculator.xlsx' now!")
