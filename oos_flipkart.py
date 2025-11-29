import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
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
            styles.append('background-color: #FFFF00; color: black;')  # yellow (black text better)
        elif 45 <= value < 60:
            styles.append('background-color: #87CEEB; color: black;')  # sky blue
        elif 60 <= value < 90:
            styles.append('background-color: #8B4513; color: white;')  # brown
        else:  # >= 90
            styles.append('background-color: #000000; color: white;')  # black
    return styles

# Process button
if sales_file and inventory_file and pm_file:
    if st.button("🚀 Process Data", type="primary", use_container_width=True):
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
                
                # Calculate DRR
                F_Sales["DRR"] = (F_Sales["Final Sale Units"] / no_of_days).round(2)
                
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
                    st.metric("Total Final Sale Units", int(F_Sales["Final Sale Units"].sum()))
                with col3:
                    st.metric("Total GMV", f"₹{F_Sales['Final Sale Amount'].sum():,.0f}")
                with col4:
                    st.metric("Avg DOC", f"{F_Sales['DOC'].mean():.1f} days")
                
                # 🔥 Styled DataFrame in app (DOC column colored)
                st.markdown("### 📊 Processed Data Preview")
                if "DOC" in F_Sales.columns:
                    styled_df = F_Sales.style.apply(style_doc_column, subset=['DOC'])
                    st.dataframe(styled_df, use_container_width=True, height=400)
                else:
                    st.dataframe(F_Sales, use_container_width=True, height=400)
                
                # Excel with same conditional formatting
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
                    
                    # Find DOC column
                    doc_col = None
                    for col in ws[1]:
                        if col.value == 'DOC':
                            doc_col = col.column
                            break
                    
                    if doc_col:
                        # Dark fills for strong visibility
                        red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                        orange_fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
                        green_fill = PatternFill(start_color='008000', end_color='008000', fill_type='solid')
                        yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                        skyblue_fill = PatternFill(start_color='87CEEB', end_color='87CEEB', fill_type='solid')
                        brown_fill = PatternFill(start_color='8B4513', end_color='8B4513', fill_type='solid')
                        black_fill = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
                        
                        white_font = Font(color="FFFFFF")  # white text for dark backgrounds
                        black_font = Font(color="000000")  # black text for light backgrounds
                        
                        from openpyxl.utils import column_index_from_string
                        if isinstance(doc_col, str):
                            doc_col_idx = column_index_from_string(doc_col)
                        else:
                            doc_col_idx = doc_col
                        
                        for row in range(2, ws.max_row + 1):
                            cell = ws.cell(row=row, column=doc_col_idx)
                            try:
                                value = float(cell.value) if cell.value is not None else 0
                                
                                if 0 <= value < 7:
                                    cell.fill = red_fill
                                    cell.font = white_font
                                elif 7 <= value < 15:
                                    cell.fill = orange_fill
                                    cell.font = white_font
                                elif 15 <= value < 30:
                                    cell.fill = green_fill
                                    cell.font = white_font
                                elif 30 <= value < 45:
                                    cell.fill = yellow_fill
                                    cell.font = black_font
                                elif 45 <= value < 60:
                                    cell.fill = skyblue_fill
                                    cell.font = black_font
                                elif 60 <= value < 90:
                                    cell.fill = brown_fill
                                    cell.font = white_font
                                elif value >= 90:
                                    cell.fill = black_fill
                                    cell.font = white_font
                            except:
                                pass
                    
                    # Save to BytesIO
                    output_final = BytesIO()
                    wb.save(output_final)
                    output_final.seek(0)
                    return output_final
                
                # Download button
                excel_data = create_formatted_excel(F_Sales)
                
                st.markdown("### 💾 Download Processed File")
                st.download_button(
                    label="📥 Download Excel with DOC Conditional Formatting",
                    data=excel_data,
                    file_name="Flipkart_Sales_Analysis_Formatted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
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
