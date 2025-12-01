# import streamlit as st
# import pandas as pd
# import numpy as np
# from io import BytesIO
# from openpyxl import load_workbook
# from openpyxl.styles import PatternFill, Font
# from openpyxl.utils import get_column_letter
# import warnings
# warnings.filterwarnings('ignore')

# # Page configuration
# st.set_page_config(page_title="OOS Flipkart Sales Analysis", page_icon="📊", layout="wide")

# # Title and description
# st.title("📊 OOS Flipkart Sales Analysis Dashboard")
# st.markdown("""
# This application analyzes Flipkart sales data and calculates important metrics like:
# - **DRR (Daily Run Rate)**: Average units sold per day
# - **DOC (Days of Coverage)**: How many days current stock will last based on sales velocity
# """)

# st.markdown("---")

# # Sidebar for inputs
# with st.sidebar:
#     st.header("⚙️ Configuration")
#     st.markdown("### Enter Number of Days")
#     st.info("📅 Enter the number of days covered in your sales report (e.g., 27 for Nov 1-27)")
#     no_of_days = st.number_input(
#         "Number of Days in Sales Period:",
#         min_value=1,
#         max_value=365,
#         value=27,
#         step=1,
#         help="This is used to calculate Daily Run Rate (DRR)"
#     )
    
#     st.markdown("---")
#     st.markdown("""
#     ### 📈 DOC Color Legend:
#     - 🔴 **Red (0-7 days)**: Critical - Immediate action needed
#     - 🟠 **Orange (7-15 days)**: Low - Reorder soon
#     - 🟢 **Green (15-30 days)**: Optimal - Good stock level
#     - 🟡 **Yellow (30-45 days)**: Monitor sales
#     - 🔵 **Sky Blue (45-60 days)**: High - Monitor closely
#     - 🟤 **Brown (60-90 days)**: Excess - Stop ordering
#     - ⬛ **Black (>90 days)**: Overstocked - Clearance needed
#     """)

# # Center column for file upload
# col1, col2, col3 = st.columns([1, 2, 1])

# with col2:
#     st.markdown("### 📁 Upload Your Files")
    
#     # File uploaders
#     sales_file = st.file_uploader(
#         "Upload Flipkart Sales Report (Excel)",
#         type=['xlsx', 'xls'],
#         help="Upload the sales report downloaded from Flipkart"
#     )
    
#     inventory_file = st.file_uploader(
#         "Upload Inventory Listing (Excel/XLS)",
#         type=['xlsx', 'xls'],
#         help="Upload the inventory/listing report from Flipkart"
#     )
    
#     pm_file = st.file_uploader(
#         "Upload Product Master (Excel)",
#         type=['xlsx', 'xls'],
#         help="Upload the product master file with brand and manager details"
#     )

# st.markdown("---")

# # 🔴🟠🟢 DOC styling for Streamlit table
# def style_doc_column(s):
#     styles = []
#     for v in s:
#         try:
#             value = float(v)
#         except (TypeError, ValueError):
#             value = 0
#         if 0 <= value < 7:
#             styles.append('background-color: #FF0000; color: white;')
#         elif 7 <= value < 15:
#             styles.append('background-color: #FFA500; color: white;')  # orange
#         elif 15 <= value < 30:
#             styles.append('background-color: #008000; color: white;')  # green
#         elif 30 <= value < 45:
#             styles.append('background-color: #FFFF00; color: black;')  # yellow (black text better)
#         elif 45 <= value < 60:
#             styles.append('background-color: #87CEEB; color: black;')  # sky blue
#         elif 60 <= value < 90:
#             styles.append('background-color: #8B4513; color: white;')  # brown
#         else:  # >= 90
#             styles.append('background-color: #000000; color: white;')  # black
#     return styles

# # Process button
# if sales_file and inventory_file and pm_file:
#     if st.button("🚀 Process Data", type="primary", width='stretch'):
#         with st.spinner("Processing your data..."):
#             try:
#                 # Read files
#                 F_Sales = pd.read_excel(sales_file)
#                 Inventory = pd.read_excel(inventory_file)
#                 PM = pd.read_excel(pm_file)
                
#                 # Remove header row from Inventory if present
#                 if Inventory.iloc[0].astype(str).str.contains('Title of your product').any():
#                     Inventory = Inventory.iloc[1:].reset_index(drop=True)
                
#                 # Merge with PM
#                 F_Sales = F_Sales.merge(
#                     PM[['FNS', 'Brand Manager', 'Brand']],
#                     left_on='Product Id',
#                     right_on='FNS',
#                     how='left'
#                 )
#                 F_Sales = F_Sales.drop(columns=['FNS'])
                
#                 # Rename and reorganize columns
#                 if 'Brand_y' in F_Sales.columns:
#                     F_Sales.rename(columns={'Brand_y': 'Brand'}, inplace=True)
#                     if 'Brand_x' in F_Sales.columns:
#                         F_Sales.drop(columns=['Brand_x'], inplace=True)
                
#                 # Reorder columns
#                 cols = F_Sales.columns.tolist()
#                 if 'Brand Manager' in cols and 'Brand' in cols and 'SKU ID' in cols:
#                     bm = cols.pop(cols.index('Brand Manager'))
#                     br = cols.pop(cols.index('Brand'))
#                     sku_pos = cols.index('SKU ID')
#                     cols.insert(sku_pos, br)
#                     cols.insert(sku_pos, bm)
#                     F_Sales = F_Sales[cols]
                
#                 # Calculate DRR
#                 F_Sales["DRR"] = (F_Sales["Final Sale Units"] / no_of_days).round(2)
                
#                 # Merge with Inventory
#                 F_Sales = F_Sales.merge(
#                     Inventory[['Flipkart Serial Number', 'System Stock count']],
#                     left_on='Product Id',
#                     right_on='Flipkart Serial Number',
#                     how='left'
#                 )
#                 F_Sales.rename(columns={'System Stock count': 'Flipkart Stock'}, inplace=True)
#                 F_Sales.drop(columns=['Flipkart Serial Number'], inplace=True)
                
#                 # Calculate DOC
#                 F_Sales["Flipkart Stock"] = pd.to_numeric(F_Sales["Flipkart Stock"], errors="coerce")
#                 F_Sales["DRR"] = pd.to_numeric(F_Sales["DRR"], errors="coerce")
                
#                 F_Sales["DOC"] = np.where(
#                     F_Sales["DRR"] > 0,
#                     F_Sales["Flipkart Stock"] / F_Sales["DRR"],
#                     np.nan
#                 )
#                 F_Sales["DOC"] = F_Sales["DOC"].round(2)
#                 F_Sales["DOC"] = F_Sales["DOC"].fillna(0)
                
#                 st.success("✅ Data processed successfully!")
                
#                 # Summary metrics
#                 col1, col2, col3, col4 = st.columns(4)
#                 with col1:
#                     st.metric("Total Products", len(F_Sales))
#                 with col2:
#                     st.metric("Total Final Sale Units", int(F_Sales["Final Sale Units"].sum()))
#                 with col3:
#                     st.metric("Total GMV", f"₹{F_Sales['Final Sale Amount'].sum():,.0f}")
#                 with col4:
#                     st.metric("Avg DOC", f"{F_Sales['DOC'].mean():.1f} days")
                
#                 # 🔥 Styled DataFrame in app (DOC column colored)
#                 st.markdown("### 📊 Processed Data Preview")
#                 if "DOC" in F_Sales.columns:
#                     styled_df = F_Sales.style.apply(style_doc_column, subset=['DOC'])
#                     st.dataframe(styled_df, width='stretch', height=400)
#                 else:
#                     st.dataframe(F_Sales,width='stretch', height=400)
                
#                 # Excel with same conditional formatting
#                 def create_formatted_excel(df):
#                     output = BytesIO()
                    
#                     # Write to Excel
#                     with pd.ExcelWriter(output, engine='openpyxl') as writer:
#                         df.to_excel(writer, index=False, sheet_name='Sales Analysis')
                    
#                     # Load workbook for formatting
#                     output.seek(0)
#                     wb = load_workbook(output)
#                     ws = wb['Sales Analysis']
                    
#                     # Freeze header row
#                     ws.freeze_panes = "A2"

#                     # Auto-filter on header row
#                     ws.auto_filter.ref = ws.dimensions

#                     # Auto column width
#                     for col in ws.columns:
#                         max_length = 0
#                         col_idx = col[0].column
#                         for cell in col:
#                             try:
#                                 if cell.value is not None:
#                                     max_length = max(max_length, len(str(cell.value)))
#                             except:
#                                 pass
#                         adjusted_width = max_length + 2
#                         ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width
                    
#                     # Find DOC column
#                     doc_col = None
#                     for col in ws[1]:
#                         if col.value == 'DOC':
#                             doc_col = col.column
#                             break
                    
#                     if doc_col:
#                         # Dark fills for strong visibility
#                         red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
#                         orange_fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
#                         green_fill = PatternFill(start_color='008000', end_color='008000', fill_type='solid')
#                         yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
#                         skyblue_fill = PatternFill(start_color='87CEEB', end_color='87CEEB', fill_type='solid')
#                         brown_fill = PatternFill(start_color='8B4513', end_color='8B4513', fill_type='solid')
#                         black_fill = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
                        
#                         white_font = Font(color="FFFFFF")  # white text for dark backgrounds
#                         black_font = Font(color="000000")  # black text for light backgrounds
                        
#                         from openpyxl.utils import column_index_from_string
#                         if isinstance(doc_col, str):
#                             doc_col_idx = column_index_from_string(doc_col)
#                         else:
#                             doc_col_idx = doc_col
                        
#                         for row in range(2, ws.max_row + 1):
#                             cell = ws.cell(row=row, column=doc_col_idx)
#                             try:
#                                 value = float(cell.value) if cell.value is not None else 0
                                
#                                 if 0 <= value < 7:
#                                     cell.fill = red_fill
#                                     cell.font = white_font
#                                 elif 7 <= value < 15:
#                                     cell.fill = orange_fill
#                                     cell.font = white_font
#                                 elif 15 <= value < 30:
#                                     cell.fill = green_fill
#                                     cell.font = white_font
#                                 elif 30 <= value < 45:
#                                     cell.fill = yellow_fill
#                                     cell.font = black_font
#                                 elif 45 <= value < 60:
#                                     cell.fill = skyblue_fill
#                                     cell.font = black_font
#                                 elif 60 <= value < 90:
#                                     cell.fill = brown_fill
#                                     cell.font = white_font
#                                 elif value >= 90:
#                                     cell.fill = black_fill
#                                     cell.font = white_font
#                             except:
#                                 pass
                    
#                     # Save to BytesIO
#                     output_final = BytesIO()
#                     wb.save(output_final)
#                     output_final.seek(0)
#                     return output_final
                
#                 # Download button
#                 excel_data = create_formatted_excel(F_Sales)
                
#                 st.markdown("### 💾 Download Processed File")
#                 st.download_button(
#                     label="📥 Download Excel with DOC Conditional Formatting",
#                     data=excel_data,
#                     file_name="Flipkart_Sales_Analysis_Formatted.xlsx",
#                     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                     width='stretch'
#                 )
                
#                 # Simple count-based insights (same ranges, no charts)
#                 st.markdown("### 📈 Stock Status Insights")
#                 col1, col2 = st.columns(2)
                
#                 with col1:
#                     critical_stock = len(F_Sales[F_Sales['DOC'] < 7])
#                     low_stock = len(F_Sales[(F_Sales['DOC'] >= 7) & (F_Sales['DOC'] < 15)])
#                     optimal_stock = len(F_Sales[(F_Sales['DOC'] >= 15) & (F_Sales['DOC'] < 30)])
                    
#                     st.markdown(f"""
#                     **Stock Alerts:**
#                     - 🔴 Critical (0-7 days): **{critical_stock} products**
#                     - 🟠 Low (7-15 days): **{low_stock} products**
#                     - 🟢 Optimal (15-30 days): **{optimal_stock} products**
#                     """)
                
#                 with col2:
#                     high_stock = len(F_Sales[(F_Sales['DOC'] >= 30) & (F_Sales['DOC'] < 45)])
#                     very_high_stock = len(F_Sales[(F_Sales['DOC'] >= 45) & (F_Sales['DOC'] < 60)])
#                     excess_stock = len(F_Sales[F_Sales['DOC'] >= 60])
                    
#                     st.markdown(f"""
#                     **Excess Stock:**
#                     - 🟡 Monitor (30-45 days): **{high_stock} products**
#                     - 🔵 High (45-60 days): **{very_high_stock} products**
#                     - 🟤 Excess (60+ days): **{excess_stock} products**
#                     """)
                
#             except Exception as e:
#                 st.error(f"❌ Error processing files: {str(e)}")
#                 st.info("Please ensure all files are in the correct format and contain the required columns.")
# else:
#     st.info("👆 Please upload all three required files to begin analysis")
    
#     with st.expander("ℹ️ Required Files Information"):
#         st.markdown("""
#         ### File Requirements:
        
#         1. **Flipkart Sales Report**: 
#            - Should contain: Product Id, SKU ID, Category, Brand, Vertical, Order Date, etc.
           
#         2. **Inventory Listing**: 
#            - Should contain: Flipkart Serial Number, System Stock count, etc.
           
#         3. **Product Master**: 
#            - Should contain: FNS, Brand Manager, Brand, etc.
        
#         All files should be in Excel format (.xlsx or .xls)
#         """)

# # Footer
# st.markdown("---")
# st.markdown("""
# <div style='text-align: center; color: gray;'>
#     <p>Flipkart Sales Analysis Dashboard | Built with Streamlit</p>
# </div>
# """, unsafe_allow_html=True)
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, Reference
import os
import traceback
import warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(page_title="OOS Flipkart Sales Analysis", page_icon="📊", layout="wide")

# Title and description
st.title("📊 OOS Flipkart Sales Analysis Dashboard")
st.markdown("""
This application analyzes Flipkart sales data and calculates important metrics like:
- **DRR (Daily Run Rate)**: Average units sold per day
- **DOC (Days of Coverage)**: How many days current stock will last based on sales velocity
""")

st.markdown("---")

# Sidebar for inputs
with st.sidebar:
    st.header("⚙️ Configuration")
    st.markdown("### Enter Number of Days")
    st.info("📅 Enter the number of days covered in your sales report (e.g., 27 for Nov 1-27)")
    no_of_days = st.number_input(
        "Number of Days in Sales Period:",
        min_value=1,
        max_value=365,
        value=27,
        step=1,
        help="This is used to calculate Daily Run Rate (DRR)"
    )
    
    st.markdown("---")
    st.markdown("""
    ### 📈 DOC Color Legend:
    - 🔴 **Red (0-7 days)**: Critical - Immediate action needed
    - 🟠 **Orange (7-15 days)**: Low - Reorder soon
    - 🟢 **Green (15-30 days)**: Optimal - Good stock level
    - 🟡 **Yellow (30-45 days)**: Monitor sales
    - 🔵 **Sky Blue (45-60 days)**: High - Monitor closely
    - 🟤 **Brown (60-90 days)**: Excess - Stop ordering
    - ⬛ **Black (>90 days)**: Overstocked - Clearance needed
    """)

# Center column for file upload
col1, col2, col3 = st.columns([1, 2, 1])

with col2:
    st.markdown("### 📁 Upload Your Files")
    
    # File uploaders
    sales_file = st.file_uploader(
        "Upload Flipkart Sales Report (Excel)",
        type=['xlsx', 'xls'],
        help="Upload the sales report downloaded from Flipkart"
    )
    
    inventory_file = st.file_uploader(
        "Upload Inventory Listing (Excel/XLS)",
        type=['xlsx', 'xls'],
        help="Upload the inventory/listing report from Flipkart"
    )
    
    pm_file = st.file_uploader(
        "Upload Product Master (Excel)",
        type=['xlsx', 'xls'],
        help="Upload the product master file with brand and manager details"
    )

st.markdown("---")

# 🔴🟠🟢 DOC styling for Streamlit table
def style_doc_column(s):
    styles = []
    for v in s:
        try:
            value = float(v)
        except (TypeError, ValueError):
            value = 0
        if 0 <= value < 7:
            styles.append('background-color: #FF0000; color: white;')
        elif 7 <= value < 15:
            styles.append('background-color: #FFA500; color: white;')  # orange
        elif 15 <= value < 30:
            styles.append('background-color: #008000; color: white;')  # green
        elif 30 <= value < 45:
            styles.append('background-color: #FFFF00; color: black;')  # yellow
        elif 45 <= value < 60:
            styles.append('background-color: #87CEEB; color: black;')  # sky blue
        elif 60 <= value < 90:
            styles.append('background-color: #8B4513; color: white;')  # brown
        else:  # >= 90
            styles.append('background-color: #000000; color: white;')  # black
    return styles

# ----- shared color fills for Excel DOC -----
DOC_FILLS = {
    "red": PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid'),
    "orange": PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid'),
    "green": PatternFill(start_color='008000', end_color='008000', fill_type='solid'),
    "yellow": PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid'),
    "sky": PatternFill(start_color='87CEEB', end_color='87CEEB', fill_type='solid'),
    "brown": PatternFill(start_color='8B4513', end_color='8B4513', fill_type='solid'),
    "black": PatternFill(start_color='000000', end_color='000000', fill_type='solid'),
}
WHITE_FONT = Font(color="FFFFFF")
BLACK_FONT = Font(color="000000")

def apply_doc_color_to_column(ws, header_row_idx: int, col_name: str = "DOC"):
    """
    Loop DOC column cells in worksheet ws and apply same color logic as app.
    """
    # Find DOC column index
    header_cells = list(ws.iter_rows(min_row=header_row_idx, max_row=header_row_idx, values_only=False))[0]
    doc_col_idx = None
    for idx, cell in enumerate(header_cells, start=1):
        if cell.value and str(cell.value).strip().lower() == col_name.lower():
            doc_col_idx = idx
            break
    if doc_col_idx is None:
        return

    max_row = ws.max_row
    for row in range(header_row_idx + 1, max_row + 1):
        cell = ws.cell(row=row, column=doc_col_idx)
        try:
            v = float(cell.value) if cell.value is not None else 0
        except (TypeError, ValueError):
            v = 0
        if 0 <= v < 7:
            cell.fill = DOC_FILLS["red"]
            cell.font = WHITE_FONT
        elif 7 <= v < 15:
            cell.fill = DOC_FILLS["orange"]
            cell.font = WHITE_FONT
        elif 15 <= v < 30:
            cell.fill = DOC_FILLS["green"]
            cell.font = WHITE_FONT
        elif 30 <= v < 45:
            cell.fill = DOC_FILLS["yellow"]
            cell.font = BLACK_FONT
        elif 45 <= v < 60:
            cell.fill = DOC_FILLS["sky"]
            cell.font = BLACK_FONT
        elif 60 <= v < 90:
            cell.fill = DOC_FILLS["brown"]
            cell.font = WHITE_FONT
        elif v >= 90:
            cell.fill = DOC_FILLS["black"]
            cell.font = WHITE_FONT

# ---------- Helper 1: main formatted Excel ----------
def create_formatted_excel(df):
    output = BytesIO()
    
    # Write to Excel
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sales Analysis')
    
    # Load workbook for formatting
    output.seek(0)
    wb = load_workbook(output)
    ws = wb['Sales Analysis']
    
    # Freeze header row
    ws.freeze_panes = "A2"

    # Auto-filter on header row
    ws.auto_filter.ref = ws.dimensions

    # Auto column width
    for col in ws.columns:
        max_length = 0
        col_idx = col[0].column
        for cell in col:
            try:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width
    
    # Apply DOC color formatting
    apply_doc_color_to_column(ws, header_row_idx=1, col_name="DOC")
    
    # Save to BytesIO
    output_final = BytesIO()
    wb.save(output_final)
    output_final.seek(0)
    return output_final

# ---------- Helper 2: fill Excel template (DataTable) + DOC formatting ----------
def fill_template_and_get_bytes(template_path: str, df: pd.DataFrame, table_name: str = "DataTable") -> BytesIO:
    """
    Load an Excel template (xlsx/xlsm) with a table named `table_name` (e.g. DataTable).
    Replace its header + rows with df and resize the table.
    Also apply DOC color formatting on that sheet.
    Returns BytesIO of modified workbook.
    """
    import re

    wb = load_workbook(template_path, keep_vba=True)
    table_sheet = None
    table_obj = None

    # Robustly find the table named table_name
    for ws in wb.worksheets:
        tables = getattr(ws, "_tables", None)
        if not tables:
            continue

        # _tables might be dict or list
        if isinstance(tables, dict):
            iter_tables = list(tables.values())
        else:
            iter_tables = list(tables)

        for tbl in iter_tables:
            name = None
            try:
                # Table object
                name = getattr(tbl, "displayName", None) or getattr(tbl, "name", None)
            except Exception:
                pass

            if name is None and isinstance(tbl, str):
                name = tbl  # just a name string

            if name == table_name:
                table_sheet = ws
                table_obj = tbl
                break

        if table_obj is not None:
            break

    if table_obj is None or table_sheet is None:
        raise RuntimeError(f"Table '{table_name}' not found in template '{template_path}'")

    # Helper: convert A1 -> (row, col)
    def cell_to_rowcol(cell_ref: str):
        m = re.match(r"([A-Z]+)(\d+)$", cell_ref)
        if not m:
            raise RuntimeError(f"Unexpected table ref format: {cell_ref}")
        col_letters, row = m.groups()
        col = 0
        for ch in col_letters:
            col = col * 26 + (ord(ch) - ord("A") + 1)
        return int(row), col

    # Clear existing table body, write new df
    ref = table_obj.ref  # e.g. "A1:H200"
    start_cell, end_cell = ref.split(":")
    start_row, start_col = cell_to_rowcol(start_cell)
    end_row, end_col = cell_to_rowcol(end_cell)

    # clear old data rows
    for r in range(start_row + 1, end_row + 1):
        for c in range(start_col, end_col + 1):
            table_sheet.cell(row=r, column=c).value = None

    # write header
    header = list(df.columns)
    for idx, col_name in enumerate(header):
        table_sheet.cell(row=start_row, column=start_col + idx, value=col_name)

    # write rows
    for r_idx, row in enumerate(df.itertuples(index=False, name=None), start=start_row + 1):
        for c_idx, v in enumerate(row, start=start_col):
            table_sheet.cell(row=r_idx, column=c_idx, value=v)

    # update table ref to new size
    new_end_row = start_row + len(df)
    new_end_col = start_col + len(header) - 1
    new_ref = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(new_end_col)}{new_end_row}"
    table_obj.ref = new_ref

    # Apply DOC formatting on this sheet's DOC column
    apply_doc_color_to_column(table_sheet, header_row_idx=start_row, col_name="DOC")

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ---------- Helper 3: fallback workbook with PivotSummary + Chart ----------
def create_pivot_fallback_workbook(df: pd.DataFrame, sheet_name: str) -> BytesIO:
    """
    Fallback workbook:
      - Data sheet with df (DOC colored)
      - DataTable
      - PivotSummary (Brand + Product Id → sum DOC & DRR, DOC colored)
      - ChartData + Chart
      - HowToPivot instructions
    """
    working = df.copy()

    if "DOC" in working.columns:
        working["DOC"] = pd.to_numeric(working["DOC"], errors="coerce")
        working = working.sort_values(by="DOC", ascending=False)

    if "DRR" in working.columns:
        working["DRR"] = pd.to_numeric(working["DRR"], errors="coerce")

    # Aggregate for summary
    if "Brand" in working.columns and "Product Id" in working.columns:
        agg = (
            working.groupby(["Brand", "Product Id"], dropna=False)[["DOC", "DRR"]]
            .sum()
            .reset_index()
        )
        agg["Brand_Parent"] = agg["Brand"].astype(str) + " | " + agg["Product Id"].astype(str)
    elif "Brand" in working.columns:
        agg = (
            working.groupby(["Brand"], dropna=False)[["DOC", "DRR"]]
            .sum()
            .reset_index()
        )
        agg["Brand_Parent"] = agg["Brand"].astype(str)
    else:
        agg = pd.DataFrame(columns=["Brand_Parent", "DOC", "DRR"])

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Data sheet
    for r in dataframe_to_rows(working, index=False, header=True):
        ws.append(r)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Auto column width
    for col in ws.columns:
        max_len = 0
        col_idx = col[0].column
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2

    # Color DOC in Data sheet
    apply_doc_color_to_column(ws, header_row_idx=1, col_name="DOC")

    # Add DataTable
    try:
        max_row = ws.max_row
        max_col = ws.max_column
        ref = f"A1:{get_column_letter(max_col)}{max_row}"
        table = Table(displayName="DataTable", ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        ws.add_table(table)
    except Exception:
        pass

    # PivotSummary sheet
    ws_pivot = wb.create_sheet("PivotSummary")
    if not agg.empty:
        for r in dataframe_to_rows(agg, index=False, header=True):
            ws_pivot.append(r)

        # Color DOC in PivotSummary
        apply_doc_color_to_column(ws_pivot, header_row_idx=1, col_name="DOC")

    # ChartData + Chart
    ws_chartdata = wb.create_sheet("ChartData")
    if not agg.empty:
        for r in dataframe_to_rows(agg[["Brand_Parent", "DOC", "DRR"]], index=False, header=True):
            ws_chartdata.append(r)
        if ws_chartdata.max_row > 1:
            chart = BarChart()
            cats = Reference(ws_chartdata, min_col=1, min_row=2, max_row=ws_chartdata.max_row)
            vals1 = Reference(ws_chartdata, min_col=2, min_row=2, max_row=ws_chartdata.max_row)
            vals2 = Reference(ws_chartdata, min_col=3, min_row=2, max_row=ws_chartdata.max_row)
            chart.add_data(vals1, titles_from_data=False)
            chart.add_data(vals2, titles_from_data=False)
            chart.set_categories(cats)
            chart.title = f"Sum DOC and DRR by Brand + Product ({sheet_name})"
            ws_chart = wb.create_sheet("Chart")
            ws_chart.add_chart(chart, "A1")

    # HowToPivot sheet
    ws_how = wb.create_sheet("HowToPivot")
    ws_how.append([f"How to create PivotTable + PivotChart ({sheet_name})"])
    ws_how.append([])
    ws_how.append(["1) Insert → PivotTable using table 'DataTable' on sheet " + sheet_name])
    ws_how.append(["2) Put 'Brand' in Rows, then 'Product Id' under it."])
    ws_how.append(["3) Put 'DOC' and 'DRR' into Values as Sum."])
    ws_how.append(["4) Insert → PivotChart from PivotTable."])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ------------------------- MAIN PROCESS BUTTON -------------------------
if sales_file and inventory_file and pm_file:
    if st.button("🚀 Process Data", type="primary", width='stretch'):
        with st.spinner("Processing your data..."):
            try:
                # Read files
                F_Sales = pd.read_excel(sales_file)
                Inventory = pd.read_excel(inventory_file)
                PM = pd.read_excel(pm_file)
                
                # Remove header row from Inventory if present
                if Inventory.iloc[0].astype(str).str.contains('Title of your product').any():
                    Inventory = Inventory.iloc[1:].reset_index(drop=True)
                
                # Merge with PM
                F_Sales = F_Sales.merge(
                    PM[['FNS', 'Brand Manager', 'Brand']],
                    left_on='Product Id',
                    right_on='FNS',
                    how='left'
                )
                F_Sales = F_Sales.drop(columns=['FNS'])
                
                # Rename and reorganize columns
                if 'Brand_y' in F_Sales.columns:
                    F_Sales.rename(columns={'Brand_y': 'Brand'}, inplace=True)
                    if 'Brand_x' in F_Sales.columns:
                        F_Sales.drop(columns=['Brand_x'], inplace=True)
                
                # Reorder columns
                cols = F_Sales.columns.tolist()
                if 'Brand Manager' in cols and 'Brand' in cols and 'SKU ID' in cols:
                    bm = cols.pop(cols.index('Brand Manager'))
                    br = cols.pop(cols.index('Brand'))
                    sku_pos = cols.index('SKU ID')
                    cols.insert(sku_pos, br)
                    cols.insert(sku_pos, bm)
                    F_Sales = F_Sales[cols]

                # CLEAN Final Sale Units: negative → 0
                if "Final Sale Units" in F_Sales.columns:
                    F_Sales["Final Sale Units"] = pd.to_numeric(F_Sales["Final Sale Units"], errors="coerce").fillna(0)
                    F_Sales.loc[F_Sales["Final Sale Units"] < 0, "Final Sale Units"] = 0
                
                # Calculate DRR using cleaned Final Sale Units
                if "Final Sale Units" in F_Sales.columns:
                    F_Sales["DRR"] = (F_Sales["Final Sale Units"] / no_of_days).round(2)
                else:
                    F_Sales["DRR"] = 0
                
                # Merge with Inventory
                F_Sales = F_Sales.merge(
                    Inventory[['Flipkart Serial Number', 'System Stock count']],
                    left_on='Product Id',
                    right_on='Flipkart Serial Number',
                    how='left'
                )
                F_Sales.rename(columns={'System Stock count': 'Flipkart Stock'}, inplace=True)
                F_Sales.drop(columns=['Flipkart Serial Number'], inplace=True)
                
                # Calculate DOC
                F_Sales["Flipkart Stock"] = pd.to_numeric(F_Sales["Flipkart Stock"], errors="coerce")
                F_Sales["DRR"] = pd.to_numeric(F_Sales["DRR"], errors="coerce")
                
                F_Sales["DOC"] = np.where(
                    F_Sales["DRR"] > 0,
                    F_Sales["Flipkart Stock"] / F_Sales["DRR"],
                    np.nan
                )
                F_Sales["DOC"] = F_Sales["DOC"].round(2)
                F_Sales["DOC"] = F_Sales["DOC"].fillna(0)
                
                st.success("✅ Data processed successfully!")
                
                # Summary metrics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Products", len(F_Sales))
                with col2:
                    if "Final Sale Units" in F_Sales.columns:
                        st.metric("Total Final Sale Units", int(F_Sales["Final Sale Units"].sum()))
                    else:
                        st.metric("Total Final Sale Units", "N/A")
                with col3:
                    if "Final Sale Amount" in F_Sales.columns:
                        st.metric("Total GMV", f"₹{F_Sales['Final Sale Amount'].sum():,.0f}")
                    else:
                        st.metric("Total GMV", "N/A")
                with col4:
                    st.metric("Avg DOC", f"{F_Sales['DOC'].mean():.1f} days")
                
                # Styled DataFrame in app (DOC column colored)
                st.markdown("### 📊 Processed Data Preview")
                if "DOC" in F_Sales.columns:
                    styled_df = F_Sales.style.apply(style_doc_column, subset=['DOC'])
                    st.dataframe(styled_df, width='stretch', height=400)
                else:
                    st.dataframe(F_Sales, width='stretch', height=400)
                
                # MAIN processed full Excel (no filters)
                excel_data = create_formatted_excel(F_Sales)
                
                st.markdown("### 💾 Download Processed File")
                st.download_button(
                    label="📥 Download Excel with DOC Conditional Formatting",
                    data=excel_data,
                    file_name="Flipkart_Sales_Analysis_Formatted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width='stretch'
                )

                # ----------------- OOS & Overstock Pivot Exports -----------------
                st.markdown("### 📂 OOS & Overstock Excel (with PivotTable + Chart)")

                # OOS: Flipkart Stock == 0
                oos_df = F_Sales.copy()
                oos_df["Flipkart Stock"] = pd.to_numeric(oos_df["Flipkart Stock"], errors="coerce").fillna(0)
                oos_df = oos_df[oos_df["Flipkart Stock"] == 0].copy()

                # Overstock: DOC >= 90
                over_df = F_Sales.copy()
                over_df["DOC"] = pd.to_numeric(over_df["DOC"], errors="coerce")
                over_df = over_df[over_df["DOC"] >= 90].copy()

                base_dir = os.path.dirname(__file__)
                tmpl_xlsm = os.path.join(base_dir, "pivot_template.xlsm")
                tmpl_xlsx = os.path.join(base_dir, "pivot_template.xlsx")
                template_path = tmpl_xlsm if os.path.exists(tmpl_xlsm) else (tmpl_xlsx if os.path.exists(tmpl_xlsx) else None)

                # ---------- OOS export ----------
                oos_bytes = None
                oos_ext = ".xlsx"
                oos_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                if template_path:
                    try:
                        buf_oos = fill_template_and_get_bytes(template_path, oos_df, table_name="DataTable")
                        oos_bytes = buf_oos.getvalue()
                        if template_path.lower().endswith(".xlsm"):
                            oos_ext = ".xlsm"
                            oos_mime = "application/vnd.ms-excel.sheet.macroEnabled.12"
                        st.success("✅ OOS: Used pivot template — PivotTable & PivotChart (your VBA macro) will build on open.")
                    except Exception:
                        st.warning("⚠️ OOS: Template fill failed, using fallback workbook.")
                        st.code(traceback.format_exc())

                if oos_bytes is None:
                    fb_oos = create_pivot_fallback_workbook(oos_df, sheet_name="OOS")
                    oos_bytes = fb_oos.getvalue()
                    oos_ext = ".xlsx"
                    oos_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    st.info("ℹ️ OOS fallback workbook generated (DataTable + PivotSummary + ChartData + HowToPivot).")

                st.download_button(
                    label="📥 Download OOS Excel (Flipkart Stock = 0)",
                    data=oos_bytes,
                    file_name=f"Flipkart_OOS_with_Pivot{oos_ext}",
                    mime=oos_mime,
                    key="oos_download",
                )

                # ---------- Overstock export ----------
                over_bytes = None
                over_ext = ".xlsx"
                over_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                if template_path:
                    try:
                        buf_over = fill_template_and_get_bytes(template_path, over_df, table_name="DataTable")
                        over_bytes = buf_over.getvalue()
                        if template_path.lower().endswith(".xlsm"):
                            over_ext = ".xlsm"
                            over_mime = "application/vnd.ms-excel.sheet.macroEnabled.12"
                        st.success("✅ Overstock: Used pivot template — PivotTable & PivotChart (your VBA macro) will build on open.")
                    except Exception:
                        st.warning("⚠️ Overstock: Template fill failed, using fallback workbook.")
                        st.code(traceback.format_exc())

                if over_bytes is None:
                    fb_over = create_pivot_fallback_workbook(over_df, sheet_name="Overstock")
                    over_bytes = fb_over.getvalue()
                    over_ext = ".xlsx"
                    over_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    st.info("ℹ️ Overstock fallback workbook generated (DataTable + PivotSummary + ChartData + HowToPivot).")

                st.download_button(
                    label="📥 Download Overstock Excel (DOC ≥ 90)",
                    data=over_bytes,
                    file_name=f"Flipkart_Overstock_with_Pivot{over_ext}",
                    mime=over_mime,
                    key="over_download",
                )
                
                # Simple count-based insights (same ranges, no charts)
                st.markdown("### 📈 Stock Status Insights")
                col1, col2 = st.columns(2)
                
                with col1:
                    critical_stock = len(F_Sales[F_Sales['DOC'] < 7])
                    low_stock = len(F_Sales[(F_Sales['DOC'] >= 7) & (F_Sales['DOC'] < 15)])
                    optimal_stock = len(F_Sales[(F_Sales['DOC'] >= 15) & (F_Sales['DOC'] < 30)])
                    
                    st.markdown(f"""
                    **Stock Alerts:**
                    - 🔴 Critical (0-7 days): **{critical_stock} products**
                    - 🟠 Low (7-15 days): **{low_stock} products**
                    - 🟢 Optimal (15-30 days): **{optimal_stock} products**
                    """)
                
                with col2:
                    high_stock = len(F_Sales[(F_Sales['DOC'] >= 30) & (F_Sales['DOC'] < 45)])
                    very_high_stock = len(F_Sales[(F_Sales['DOC'] >= 45) & (F_Sales['DOC'] < 60)])
                    excess_stock = len(F_Sales[F_Sales['DOC'] >= 60])
                    
                    st.markdown(f"""
                    **Excess Stock:**
                    - 🟡 Monitor (30-45 days): **{high_stock} products**
                    - 🔵 High (45-60 days): **{very_high_stock} products**
                    - 🟤 Excess (60+ days): **{excess_stock} products**
                    """)
                
            except Exception as e:
                st.error(f"❌ Error processing files: {str(e)}")
                st.info("Please ensure all files are in the correct format and contain the required columns.")
else:
    st.info("👆 Please upload all three required files to begin analysis")
    
    with st.expander("ℹ️ Required Files Information"):
        st.markdown("""
        ### File Requirements:
        
        1. **Flipkart Sales Report**: 
           - Should contain: Product Id, SKU ID, Category, Brand, Vertical, Order Date, etc.
           
        2. **Inventory Listing**: 
           - Should contain: Flipkart Serial Number, System Stock count, etc.
           
        3. **Product Master**: 
           - Should contain: FNS, Brand Manager, Brand, etc.
        
        All files should be in Excel format (.xlsx or .xls)
        """)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: gray;'>
    <p>Flipkart Sales Analysis Dashboard | Built with Streamlit</p>
</div>
""", unsafe_allow_html=True)








