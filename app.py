import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(page_title="Sales & Inventory Dashboard", layout="wide", page_icon="📊")

# Custom CSS for better styling
st.markdown("""
    <style>
    .main {
        padding: 0rem 1rem;
    }
    .metric-container {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
    }
    .stMetric {
        background-color: white;
        padding: 15px;
        border-radius: 8px;
    }
    .dataframe {
        font-size: 12px;
    }
    </style>
    """, unsafe_allow_html=True)

# Title
st.markdown("<h1 style='text-align: center; color: #1f77b4;'>📊 Sales & Inventory Analytics Dashboard</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center; color: #666;'>Comprehensive Overview of Balance & Sales Performance</h3>", unsafe_allow_html=True)
st.markdown("---")

# File uploader
uploaded_file = st.file_uploader("Upload Excel File (with 'Balance' and 'Sales' sheets)", type=['xlsx', 'xls'])

@st.cache_data(ttl=3600)
def load_and_process_data(uploaded_file):
    """Load and process the Excel file with your specific column structure"""
    try:
        # Read both sheets
        sales_df = pd.read_excel(uploaded_file, sheet_name='Sales')
        balance_df = pd.read_excel(uploaded_file, sheet_name='Balance')
        
        # Display raw data info
        st.sidebar.info(f"""
        **Raw Data Loaded:**
        - 📊 Sales Records: {len(sales_df):,}
        - 📦 Balance Records: {len(balance_df):,}
        """)
        
        # Function to clean column names - preserve original case but strip spaces
        def clean_columns(df):
            df.columns = df.columns.astype(str).str.strip()
            return df
        
        # Clean column names
        sales_df = clean_columns(sales_df)
        balance_df = clean_columns(balance_df)
        
        # Show column names for verification
        with st.sidebar.expander("🔍 Verify Column Names"):
            st.write("**Balance Sheet Columns:**", list(balance_df.columns))
            st.write("**Sales Sheet Columns:**", list(sales_df.columns))
        
        # Function to find column with case-insensitive matching
        def find_column(df, possible_names):
            df_cols_upper = {col.upper(): col for col in df.columns}
            for name in possible_names:
                if name.upper() in df_cols_upper:
                    return df_cols_upper[name.upper()]
            return None
        
        # Identify key columns for BALANCE sheet (based on your structure)
        balance_style_col = find_column(balance_df, ['Style_ID', 'STYLE_ID', 'StyleID'])
        balance_year_col = find_column(balance_df, ['YEAR', 'Year'])
        balance_month_col = find_column(balance_df, ['MONTH', 'Month'])
        balance_qty_col = find_column(balance_df, ['Balance_QTY', 'BALANCE_QTY', 'Balance', 'Qty'])
        balance_date_col = find_column(balance_df, ['Date', 'DATE'])
        
        # Identify key columns for SALES sheet (based on your structure)
        sales_style_col = find_column(sales_df, ['Style_ID', 'STYLE_ID', 'StyleID'])
        sales_sku_col = find_column(sales_df, ['SKU', 'Sku'])
        sales_year_col = find_column(sales_df, ['YEAR', 'Year'])
        sales_month_col = find_column(sales_df, ['MONTH', 'Month'])
        sales_qty_col = find_column(sales_df, ['Qty', 'QTY', 'Quantity', 'Sales_QTY'])
        sales_date_col = find_column(sales_df, ['Date', 'DATE'])
        
        # Verify required columns exist for Balance sheet
        required_balance_cols = {
            'Style_ID': balance_style_col,
            'YEAR': balance_year_col,
            'MONTH': balance_month_col,
            'Balance_QTY': balance_qty_col
        }
        
        # Verify required columns exist for Sales sheet
        required_sales_cols = {
            'Style_ID': sales_style_col,
            'YEAR': sales_year_col,
            'MONTH': sales_month_col,
            'Qty': sales_qty_col
        }
        
        missing_balance = [k for k, v in required_balance_cols.items() if v is None]
        missing_sales = [k for k, v in required_sales_cols.items() if v is None]
        
        if missing_balance:
            st.error(f"❌ Missing columns in Balance sheet: {', '.join(missing_balance)}")
            st.info("Available columns: " + ", ".join(balance_df.columns))
            st.stop()
        
        if missing_sales:
            st.error(f"❌ Missing columns in Sales sheet: {', '.join(missing_sales)}")
            st.info("Available columns: " + ", ".join(sales_df.columns))
            st.stop()
        
        # Create clean BALANCE dataframe with standardized names
        balance_clean = pd.DataFrame({
            'STYLE_ID': balance_df[balance_style_col].astype(str).str.strip(),
            'YEAR': pd.to_numeric(balance_df[balance_year_col], errors='coerce'),
            'MONTH': pd.to_numeric(balance_df[balance_month_col], errors='coerce'),
            'BALANCE_QTY': pd.to_numeric(balance_df[balance_qty_col], errors='coerce').fillna(0)
        })
        
        # Add Date column if it exists
        if balance_date_col:
            try:
                balance_clean['BALANCE_DATE'] = pd.to_datetime(balance_df[balance_date_col], errors='coerce')
            except:
                balance_clean['BALANCE_DATE'] = None
        
        # Create clean SALES dataframe with standardized names
        sales_clean = pd.DataFrame({
            'STYLE_ID': sales_df[sales_style_col].astype(str).str.strip(),
            'YEAR': pd.to_numeric(sales_df[sales_year_col], errors='coerce'),
            'MONTH': pd.to_numeric(sales_df[sales_month_col], errors='coerce'),
            'SALES_QTY': pd.to_numeric(sales_df[sales_qty_col], errors='coerce').fillna(0)
        })
        
        # Add SKU column if it exists
        if sales_sku_col:
            sales_clean['SKU'] = sales_df[sales_sku_col].astype(str).str.strip()
        
        # Add Date column if it exists
        if sales_date_col:
            try:
                sales_clean['SALES_DATE'] = pd.to_datetime(sales_df[sales_date_col], errors='coerce')
            except:
                sales_clean['SALES_DATE'] = None
        
        # Add additional columns from sales if they exist (based on your structure)
        additional_cols_mapping = {
            'Subcategory': ['Subcategory', 'SUBCATEGORY', 'Sub_Category'],
            'Season': ['Season', 'SEASON'],
            'Brand': ['Brand', 'BRAND'],
            'Color': ['Color', 'COLOR'],
            'Heel_Type_1': ['Heel_Type 1', 'Heel Type 1', 'HEEL_TYPE_1', 'Heel_Type_1'],
            'Maketplace': ['Maketplace', 'MAKETPLACE', 'Marketplace', 'MARKETPLACE'],
            'MRP': ['MRP', 'Mrp'],
            'SP': ['SP', 'Sp', 'Selling_Price'],
            'Size': ['Size', 'SIZE'],
            'FOB': ['FOB', 'Fob']
        }
        
        for standard_name, possible_names in additional_cols_mapping.items():
            found_col = find_column(sales_df, possible_names)
            if found_col:
                sales_clean[standard_name] = sales_df[found_col]
        
        # Handle duplicate sales records (same style, year, month)
        duplicate_check = sales_clean.duplicated(subset=['STYLE_ID', 'YEAR', 'MONTH'], keep=False).sum()
        if duplicate_check > 0:
            st.sidebar.warning(f"⚠️ Found {duplicate_check} duplicate sales records. Aggregating...")
            
            # Get list of columns to aggregate
            agg_dict = {'SALES_QTY': 'sum'}
            for col in sales_clean.columns:
                if col not in ['STYLE_ID', 'YEAR', 'MONTH', 'SALES_QTY']:
                    # For categorical columns, take first non-null value
                    agg_dict[col] = 'first'
            
            sales_clean = sales_clean.groupby(['STYLE_ID', 'YEAR', 'MONTH'], as_index=False).agg(agg_dict)
        
        # MERGE DATA on STYLE_ID, YEAR, MONTH (as you requested)
        st.sidebar.info("🔗 Merging data on: STYLE_ID, YEAR, MONTH")
        
        merged_df = pd.merge(
            balance_clean,
            sales_clean,
            on=['STYLE_ID', 'YEAR', 'MONTH'],
            how='left',  # Keep all balance records (inventory) even if no sales
            suffixes=('_BALANCE', '_SALES')
        )
        
        # Fill missing sales with 0 (products with balance but no sales)
        merged_df['SALES_QTY'] = merged_df['SALES_QTY'].fillna(0)
        
        # Calculate % sold (handle division by zero)
        merged_df['PCT_SOLD'] = np.where(
            merged_df['BALANCE_QTY'] > 0,
            (merged_df['SALES_QTY'] / merged_df['BALANCE_QTY']) * 100,
            0
        )
        
        # Add month name for display
        month_names = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 
                      6: 'June', 7: 'July', 8: 'August', 9: 'September', 
                      10: 'October', 11: 'November', 12: 'December'}
        merged_df['MONTH_NAME'] = merged_df['MONTH'].map(month_names)
        
        # Add Year-Month column for time series
        merged_df['YEAR_MONTH'] = merged_df['YEAR'].astype(str) + '-' + merged_df['MONTH'].astype(str).str.zfill(2)
        
        # Data validation summary
        total_balance = merged_df['BALANCE_QTY'].sum()
        total_sales = merged_df['SALES_QTY'].sum()
        matched_records = len(merged_df[merged_df['SALES_QTY'] > 0])
        unmatched_records = len(merged_df[merged_df['SALES_QTY'] == 0])
        
        st.sidebar.success(f"""
        **Data Processing Complete:**
        - ✅ Total Balance Qty: {total_balance:,.0f}
        - ✅ Total Sales Qty: {total_sales:,.0f}
        - ✅ Matched Records: {matched_records:,}
        - ✅ Products with no sales: {unmatched_records:,}
        - ✅ Overall % Sold: {(total_sales/total_balance*100 if total_balance > 0 else 0):.1f}%
        """)
        
        # Show merge statistics
        with st.sidebar.expander("🔍 Merge Statistics"):
            st.write(f"**Join Details:**")
            st.write(f"- Balance records before merge: {len(balance_clean):,}")
            st.write(f"- Sales records before merge: {len(sales_clean):,}")
            st.write(f"- Merged records: {len(merged_df):,}")
            st.write(f"- Successful matches: {matched_records:,} ({matched_records/len(merged_df)*100:.1f}%)")
            
            # Check for Style_ID mismatches
            balance_styles = set(balance_clean['STYLE_ID'].unique())
            sales_styles = set(sales_clean['STYLE_ID'].unique())
            
            common_styles = balance_styles.intersection(sales_styles)
            only_in_balance = balance_styles - sales_styles
            only_in_sales = sales_styles - balance_styles
            
            st.write(f"\n**Style_ID Match Analysis:**")
            st.write(f"- Common Style_IDs: {len(common_styles):,}")
            st.write(f"- Style_IDs only in Balance: {len(only_in_balance):,}")
            st.write(f"- Style_IDs only in Sales: {len(only_in_sales):,}")
            
            if len(only_in_balance) > 0:
                st.warning(f"⚠️ {len(only_in_balance):,} products in Balance have no matching sales")
            if len(only_in_sales) > 0:
                st.info(f"ℹ️ {len(only_in_sales):,} products in Sales have no matching balance")
        
        return merged_df
        
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        import traceback
        st.write("Detailed error:", traceback.format_exc())
        st.stop()

if uploaded_file is not None:
    try:
        # Load and process data
        with st.spinner('🔍 Loading and processing data...'):
            merged_df = load_and_process_data(uploaded_file)
        
        # Display success message
        st.success(f"✅ Data loaded successfully! {len(merged_df):,} records processed")
        
        # Data summary metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Balance", f"{merged_df['BALANCE_QTY'].sum():,.0f}")
        with col2:
            st.metric("Total Sales", f"{merged_df['SALES_QTY'].sum():,.0f}")
        with col3:
            st.metric("Unique Products", f"{merged_df['STYLE_ID'].nunique():,.0f}")
        with col4:
            pct_sold = (merged_df['SALES_QTY'].sum() / merged_df['BALANCE_QTY'].sum() * 100) if merged_df['BALANCE_QTY'].sum() > 0 else 0
            st.metric("% Sold", f"{pct_sold:.1f}%")
        
        st.markdown("---")
        
        # Sidebar filters
        st.sidebar.header("🔍 Filter Options")
        
        # Get unique years and months
        years = sorted([int(y) for y in merged_df['YEAR'].dropna().unique() if not pd.isna(y)])
        months = sorted([int(m) for m in merged_df['MONTH'].dropna().unique() if not pd.isna(m)])
        month_names = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 
                      6: 'June', 7: 'July', 8: 'August', 9: 'September', 
                      10: 'October', 11: 'November', 12: 'December'}
        
        selected_year = st.sidebar.selectbox("Select Year", ['All'] + years)
        selected_month = st.sidebar.selectbox("Select Month", ['All'] + [month_names[m] for m in months if m in month_names])
        
        # Filter data
        filtered_df = merged_df.copy()
        
        if selected_year != 'All':
            filtered_df = filtered_df[filtered_df['YEAR'] == selected_year]
        
        if selected_month != 'All':
            month_num = [k for k, v in month_names.items() if v == selected_month][0]
            filtered_df = filtered_df[filtered_df['MONTH'] == month_num]
        
        # Display filter summary
        st.sidebar.markdown("---")
        st.sidebar.info(f"""
        **Filter Applied:**
        - 📅 Year: {selected_year}
        - 📆 Month: {selected_month}
        - 📊 Records: {len(filtered_df):,}
        - 📦 Filtered Balance: {filtered_df['BALANCE_QTY'].sum():,.0f}
        - 💰 Filtered Sales: {filtered_df['SALES_QTY'].sum():,.0f}
        - 📈 Filtered % Sold: {(filtered_df['SALES_QTY'].sum()/filtered_df['BALANCE_QTY'].sum()*100 if filtered_df['BALANCE_QTY'].sum() > 0 else 0):.1f}%
        """)
        
        if len(filtered_df) == 0:
            st.warning("⚠️ No data available for the selected filters.")
        else:
            # Helper function for grouped analysis
            def analyze_by_category(df, category_col, category_name):
                if category_col not in df.columns:
                    return pd.DataFrame()
                
                # Group by category
                grouped = df.groupby(category_col, observed=True).agg({
                    'BALANCE_QTY': 'sum',
                    'SALES_QTY': 'sum',
                    'STYLE_ID': 'nunique'  # Count unique products
                }).reset_index()
                
                # Calculate percentage sold and average per product
                grouped['PCT_SOLD'] = np.where(
                    grouped['BALANCE_QTY'] > 0,
                    (grouped['SALES_QTY'] / grouped['BALANCE_QTY']) * 100,
                    0
                )
                
                grouped['AVG_SALES_PER_PRODUCT'] = np.where(
                    grouped['STYLE_ID'] > 0,
                    grouped['SALES_QTY'] / grouped['STYLE_ID'],
                    0
                )
                
                # Sort by sales descending
                grouped = grouped.sort_values('SALES_QTY', ascending=False)
                grouped.rename(columns={category_col: category_name, 'STYLE_ID': 'PRODUCT_COUNT'}, inplace=True)
                
                return grouped
            
            # Time Series Analysis
            st.markdown("### 📈 Sales Trend Over Time")
            
            # Group by Year-Month
            time_series = filtered_df.groupby('YEAR_MONTH').agg({
                'BALANCE_QTY': 'sum',
                'SALES_QTY': 'sum',
                'STYLE_ID': 'nunique'
            }).reset_index()
            
            # Calculate percentage
            time_series['PCT_SOLD'] = np.where(
                time_series['BALANCE_QTY'] > 0,
                (time_series['SALES_QTY'] / time_series['BALANCE_QTY']) * 100,
                0
            )
            
            # Create time series chart
            fig_time = go.Figure()
            
            # Add Sales line
            fig_time.add_trace(go.Scatter(
                x=time_series['YEAR_MONTH'],
                y=time_series['SALES_QTY'],
                mode='lines+markers',
                name='Sales Quantity',
                line=dict(color='#1f77b4', width=3),
                marker=dict(size=8)
            ))
            
            # Add Balance line
            fig_time.add_trace(go.Scatter(
                x=time_series['YEAR_MONTH'],
                y=time_series['BALANCE_QTY'],
                mode='lines+markers',
                name='Inventory Balance',
                line=dict(color='#ff7f0e', width=3, dash='dash'),
                marker=dict(size=8)
            ))
            
            fig_time.update_layout(
                height=400,
                xaxis_title="Time Period",
                yaxis_title="Quantity",
                hovermode='x unified',
                template='plotly_white',
                legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01)
            )
            
            st.plotly_chart(fig_time, use_container_width=True)
            
            # Identify which categorical columns are available
            available_categories = []
            category_options = {
                'Season': 'Season',
                'Subcategory': 'Subcategory', 
                'Color': 'Color',
                'Brand': 'Brand',
                'Heel_Type_1': 'Heel Type',
                'Maketplace': 'Marketplace',
                'Size': 'Size'
            }
            
            for col, name in category_options.items():
                if col in filtered_df.columns:
                    available_categories.append((col, name))
            
            # Create charts for available categories
            if available_categories:
                st.markdown("### 📊 Performance by Category")
                
                # Create columns for charts (max 3 per row)
                num_charts = len(available_categories)
                cols_per_row = 3
                num_rows = (num_charts + cols_per_row - 1) // cols_per_row
                
                for row in range(num_rows):
                    cols = st.columns(cols_per_row)
                    start_idx = row * cols_per_row
                    end_idx = min(start_idx + cols_per_row, num_charts)
                    
                    for idx, (col_name, display_name) in enumerate(available_categories[start_idx:end_idx]):
                        with cols[idx]:
                            category_data = analyze_by_category(filtered_df, col_name, display_name)
                            if not category_data.empty:
                                st.markdown(f"#### {display_name}")
                                
                                # Create bar chart for top 10 categories
                                top_data = category_data.head(10)
                                
                                fig = px.bar(top_data, x=display_name, y='SALES_QTY',
                                           color='SALES_QTY', 
                                           color_continuous_scale='viridis',
                                           text='SALES_QTY')
                                
                                fig.update_traces(
                                    texttemplate='%{text:.0f}', 
                                    textposition='outside'
                                )
                                fig.update_layout(
                                    height=350,
                                    showlegend=False,
                                    xaxis_title=display_name,
                                    yaxis_title="Sales Quantity",
                                    xaxis={'categoryorder': 'total descending'}
                                )
                                
                                st.plotly_chart(fig, use_container_width=True)
                
                st.markdown("---")
            
            # Detailed Data View
            st.markdown("### 📋 Detailed Data View")
            
            # Column selector for display
            all_columns = list(filtered_df.columns)
            default_cols = ['STYLE_ID', 'YEAR', 'MONTH_NAME', 'BALANCE_QTY', 'SALES_QTY', 'PCT_SOLD']
            
            # Add available additional columns
            for col in ['SKU', 'Subcategory', 'Brand', 'Color', 'Maketplace', 'Season']:
                if col in all_columns and col not in default_cols:
                    default_cols.append(col)
            
            selected_cols = st.multiselect(
                "Select columns to display:",
                all_columns,
                default=default_cols
            )
            
            if selected_cols:
                # Display data with formatting
                display_df = filtered_df[selected_cols].copy()
                
                # Apply formatting
                def format_row(row):
                    if 'PCT_SOLD' in display_df.columns:
                        if row['PCT_SOLD'] > 100:
                            return ['background-color: #ffcccc'] * len(row)
                        elif row['PCT_SOLD'] > 80:
                            return ['background-color: #ccffcc'] * len(row)
                    return [''] * len(row)
                
                st.dataframe(
                    display_df.style.apply(format_row, axis=1).format({
                        'BALANCE_QTY': '{:,.0f}',
                        'SALES_QTY': '{:,.0f}',
                        'PCT_SOLD': '{:.1f}%'
                    }),
                    hide_index=True,
                    use_container_width=True,
                    height=400
                )
            
            # Data validation section
            with st.expander("🔍 Data Validation & Statistics"):
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**Data Quality Check:**")
                    quality_df = pd.DataFrame({
                        'Metric': [
                            'Total Records',
                            'Unique Style IDs', 
                            'Records with Sales > 0',
                            'Records with Sales = 0',
                            'Average Balance per Record',
                            'Average Sales per Record',
                            'Max % Sold',
                            'Min % Sold'
                        ],
                        'Value': [
                            len(filtered_df),
                            filtered_df['STYLE_ID'].nunique(),
                            (filtered_df['SALES_QTY'] > 0).sum(),
                            (filtered_df['SALES_QTY'] == 0).sum(),
                            filtered_df['BALANCE_QTY'].mean(),
                            filtered_df['SALES_QTY'].mean(),
                            filtered_df['PCT_SOLD'].max(),
                            filtered_df['PCT_SOLD'].min()
                        ]
                    })
                    st.dataframe(quality_df, use_container_width=True)
                
                with col2:
                    st.write("**Top 10 Performing Products:**")
                    top_products = filtered_df.sort_values('SALES_QTY', ascending=False).head(10)
                    top_display = top_products[['STYLE_ID', 'BALANCE_QTY', 'SALES_QTY', 'PCT_SOLD']].copy()
                    if 'SKU' in filtered_df.columns:
                        top_display['SKU'] = top_products['SKU']
                    st.dataframe(
                        top_display.style.format({
                            'BALANCE_QTY': '{:,.0f}',
                            'SALES_QTY': '{:,.0f}',
                            'PCT_SOLD': '{:.1f}%'
                        }),
                        hide_index=True,
                        use_container_width=True
                    )
                
                st.write("**Data Sample (First 20 Rows):**")
                sample_cols = ['STYLE_ID', 'YEAR', 'MONTH', 'BALANCE_QTY', 'SALES_QTY', 'PCT_SOLD']
                if 'SKU' in filtered_df.columns:
                    sample_cols.insert(1, 'SKU')
                st.dataframe(filtered_df[sample_cols].head(20), use_container_width=True)
                
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        import traceback
        st.write("Detailed error:", traceback.format_exc())
else:
    st.info("👆 Please upload an Excel file to begin analyzing your data.")
    
    # Instructions for Excel preparation
    with st.expander("📋 Your Excel File Structure"):
        st.markdown("""
        ### **Balance Sheet Columns (Required):**
        - `Style_ID` - Product identifier
        - `YEAR` - Balance year (e.g., 2024)
        - `MONTH` - Balance month (1-12)
        - `Balance_QTY` - Available stock quantity
        
        ### **Sales Sheet Columns (Required):**
        - `Style_ID` - Product identifier (must match Balance sheet Style_ID)
        - `YEAR` - Sales year (e.g., 2024)
        - `MONTH` - Sales month (1-12)
        - `Qty` - Quantity sold
        
        ### **Sales Sheet Columns (Optional but useful):**
        - `SKU` - Additional product identifier
        - `Subcategory` - Product subcategory
        - `Heel_Type 1` - Type of heel
        - `Maketplace` - Selling marketplace
        - `Season` - Product season
        - `Brand` - Product brand
        - `Color` - Product color
        - `MRP` - Maximum Retail Price
        - `SP` - Selling Price
        - `Size` - Product size
        - `FOB` - Free on Board price
        
        ### **How the Data is Joined:**
        1. **Primary Join Keys:** `Style_ID`, `YEAR`, `MONTH`
        2. **Join Type:** LEFT JOIN (keeps all balance records)
        3. **Matching Logic:** Exact match on all three columns
        4. **Missing Sales:** Products with balance but no sales will show 0 sales
        
        ### **Important Notes:**
        1. The Style_ID in both sheets must match exactly (case and whitespace handled)
        2. Year and Month columns must be numeric (1-12 for months)
        3. The dashboard shows detailed merge statistics in the sidebar
        4. Products appearing in Sales but not in Balance will not appear in merged results
        """)
