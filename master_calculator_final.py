× B', 'A ÷ B', 'A^B', '√A', '(A × B)/100', 'A % B']
        }
        df_math = pd.DataFrame(math_data)
        df_math.at[0, 'Result'] = df_math.at[0, 'Value A'] + df_math.at[0, 'Value B']
        df_math.at[1, 'Result'] = df_math.at[1, 'Value A'] - df_math.at[1, 'Value B']
        df_math.at[2, 'Result'] = df_math.at[2, 'Value A'] * df_math.at[2, 'Value B']
        df_math.at[3, 'Result'] = round(df_math.at[3, 'Value A'] / df_math.at[3, 'Value B'], 2)
        df_math.at[4, 'Result'] = df_math.at[4, 'Value A'] ** df_math.at[4, 'Value B']
        df_math.at[5, 'Result'] = round(df_math.at[5, 'Value A'] ** 0.5, 2)
        df_math.at[6, 'Result'] = (df_math.at[6, 'Value A'] * df_math.at[6, 'Value B']) / 100
        df_math.at[7, 'Result'] = df_math.at[7, 'Value A'] % df_math.at[7, 'Value B']
        
        # ========== SHEET 7: GEOMETRY CALCULATOR ==========
        geo_data = {
            'Shape': ['Circle', 'Square', 'Rectangle', 'Triangle', 'Sphere', 'Cube', 'Cylinder'],
            'Parameter 1 (r/s/l)': [5, 4, 6, 3, 3, 4, 5],
            'Parameter 2 (w/h)': [0, 0, 8, 4, 0, 0, 8],
            'Area / Surface Area': [0, 0, 0, 0, 0, 0, 0],
            'Perimeter / Volume': [0, 0, 0, 0, 0, 0, 0],
            'Formula Used': ['πr²', 's²', 'l×w', '½×b×h', '4πr²', '6s²', '2πrh']
        }
        df_geo = pd.DataFrame(geo_data)
        
        # Circle
        df_geo.at[0, 'Area / Surface Area'] = round(math.pi * (df_geo.at[0, 'Parameter 1 (r/s/l)'] ** 2), 2)
        df_geo.at[0, 'Perimeter / Volume'] = round(2 * math.pi * df_geo.at[0, 'Parameter 1 (r/s/l)'], 2)
        # Square
        df_geo.at[1, 'Area / Surface Area'] = df_geo.at[1, 'Parameter 1 (r/s/l)'] ** 2
        df_geo.at[1, 'Perimeter / Volume'] = 4 * df_geo.at[1, 'Parameter 1 (r/s/l)']
        # Rectangle
        df_geo.at[2, 'Area / Surface Area'] = df_geo.at[2, 'Parameter 1 (r/s/l)'] * df_geo.at[2, 'Parameter 2 (w/h)']
        df_geo.at[2, 'Perimeter / Volume'] = 2 * (df_geo.at[2, 'Parameter 1 (r/s/l)'] + df_geo.at[2, 'Parameter 2 (w/h)'])
        # Triangle
        df_geo.at[3, 'Area / Surface Area'] = 0.5 * df_geo.at[3, 'Parameter 1 (r/s/l)'] * df_geo.at[3, 'Parameter 2 (w/h)']
        df_geo.at[3, 'Perimeter / Volume'] = df_geo.at[3, 'Parameter 1 (r/s/l)'] + df_geo.at[3, 'Parameter 2 (w/h)'] + math.sqrt(df_geo.at[3, 'Parameter 1 (r/s/l)']**2 + df_geo.at[3, 'Parameter 2 (w/h)']**2)
        # Sphere
        df_geo.at[4, 'Area / Surface Area'] = round(4 * math.pi * (df_geo.at[4, 'Parameter 1 (r/s/l)'] ** 2), 2)
        df_geo.at[4, 'Perimeter / Volume'] = round((4/3) * math.pi * (df_geo.at[4, 'Parameter 1 (r/s/l)'] ** 3), 2)
        # Cube
        df_geo.at[5, 'Area / Surface Area'] = 6 * (df_geo.at[5, 'Parameter 1 (r/s/l)'] ** 2)
        df_geo.at[5, 'Perimeter / Volume'] = df_geo.at[5, 'Parameter 1 (r/s/l)'] ** 3
        # Cylinder
        df_geo.at[6, 'Area / Surface Area'] = round(2 * math.pi * df_geo.at[6, 'Parameter 1 (r/s/l)'] * df_geo.at[6, 'Parameter 2 (w/h)'], 2)
        df_geo.at[6, 'Perimeter / Volume'] = round(math.pi * (df_geo.at[6, 'Parameter 1 (r/s/l)'] ** 2) * df_geo.at[6, 'Parameter 2 (w/h)'], 2)
        
        # ========== SHEET 8: PERCENTAGE CALCULATOR ==========
        pct_data = {
            'Calculation Type': ['🏷️ Discount', '💰 Sales Tax', '💵 Tip/Gratuity', '📈 Profit Margin', 
                                '📊 Markup', '🎯 Commission', '📉 Loss', '⭐ Grade'],
            'Original         # ========== WRITE ALL SHEETS ==========
        print("📝 Writing Excel sheets...")
        
        with pd.ExcelWriter(self.filename, engine='openpyxl') as writer:
            df_index.to_excel(writer, sheet_name='🏠 INDEX', index=False)
            df_stock.to_excel(writer, sheet_name='💰 STOCK', index=False)
            df_demo.to_excel(writer, sheet_name='👥 DEMOGRAPHY', index=False)
            df_stats.to_excel(writer, sheet_name='📈 STATISTICS', index=False)
            df_finance.to_excel(writer, sheet_name='💳 FINANCE', index=False)
            df_math.to_excel(writer, sheet_name='🔢 MATH', index=False)
            df_geo.to_excel(writer, sheet_name='📐 GEOMETRY', index=False)
            df_pct.to_excel(writer, sheet_name='🎯 PERCENTAGE', index=False)
            df_trend.to_excel(writer, sheet_name='📉 TREND', index=False)
            df_unit.to_excel(writer, sheet_name='⚖️ UNIT', index=False)
            df_instructions.to_excel(writer, sheet_name='📖 HELP', index=False)
        
        # Apply formatting
        self._apply_formatting()
        
        return True
    
    def _apply_formatting(self):
        """Apply professional formatting to all sheets"""
        try:
            wb = openpyxl.load_workbook(self.filename)
            
            # Color scheme for different sheets
            colors = {
                '🏠 INDEX': '2E7D32',
                '💰 STOCK': 'C62828',
                '👥 DEMOGRAPHY': '1565C0',
                '📈 STATISTICS': '6A1B9A',
                '💳 FINANCE': 'E65100',
                '🔢 MATH': '00695C',
                '📐 GEOMETRY': '283593',
                '🎯 PERCENTAGE': '827717',
                '📉 TREND': 'BF360C',
                '⚖️ UNIT': '00838F',
                '📖 HELP': '455A64'
            }
            
            header_font = Font(bold=True, color="FFFFFF", size=11)
            border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                           top=Side(style='thin'), bottom=Side(style='thin'))
            
            for sheet_name, color in colors.items():
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    header_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                    
                    # Format header row
                    for cell in ws[1]:
                        cell.font = header_font
                        cell.fill = header_fill
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.border = border
                    
                    # Auto-fit columns
                    for column in ws.columns:
                        max_length = 0
                        col_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if cell.value and len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        ws.column_dimensions[col
