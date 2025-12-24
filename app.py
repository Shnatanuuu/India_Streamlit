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
            st.write("**Sales Columns:**", list(sales_df.columns))
            st.write("**Balance Columns:**", list(balance_df.columns))
        
        # Function to find column with case-insensitive matching
        def find_column(df, possible_names):
            df_cols_upper = {col.upper(): col for col in df.columns}
            for name in possible_names:
                if name.upper() in df_cols_upper:
                    return df_cols_upper[name.upper()]
            return None
        
        # Identify key columns for BALANCE sheet
        balance_style_col = find_column(balance_df, ['Style_ID', 'STYLE_ID', 'StyleID'])
        balance_year_col = find_column(balance_df, ['YEAR', 'Year'])
        balance_month_col = find_column(balance_df, ['MONTH', 'Month'])
        balance_qty_col = find_column(balance_df, ['Balance_QTY', 'BALANCE_QTY', 'Balance', 'Qty'])
        
        # Identify key columns for SALES sheet
        sales_style_col = find_column(sales_df, ['Style_ID', 'STYLE_ID', 'StyleID', 'SKU'])
        sales_year_col = find_column(sales_df, ['YEAR', 'Year'])
        sales_month_col = find_column(sales_df, ['MONTH', 'Month'])
        sales_qty_col = find_column(sales_df, ['Qty', 'QTY', 'Quantity', 'Sales_QTY'])
        
        # Verify required columns exist
        required_cols_balance = {
            'Balance Style': balance_style_col,
            'Balance Year': balance_year_col,
            'Balance Month': balance_month_col,
            'Balance Qty': balance_qty_col
        }
        
        required_cols_sales = {
            'Sales Style': sales_style_col,
            'Sales Year': sales_year_col,
            'Sales Month': sales_month_col,
            'Sales Qty': sales_qty_col
        }
        
        missing_balance = [k for k, v in required_cols_balance.items() if v is None]
        missing_sales = [k for k, v in required_cols_sales.items() if v is None]
        
        if missing_balance:
            st.error(f"❌ Missing columns in Balance sheet: {', '.join(missing_balance)}")
            st.info("Available columns: " + ", ".join(balance_df.columns))
            st.stop()
        
        if missing_sales:
            st.error(f"❌ Missing columns in Sales sheet: {', '.join(missing_sales)}")
            st.info("Available columns: " + ", ".join(sales_df.columns))
            st.stop()
        
        # Create clean dataframes with standardized names
        balance_clean = pd.DataFrame({
            'STYLE_ID': balance_df[balance_style_col].astype(str).str.strip(),
            'YEAR': pd.to_numeric(balance_df[balance_year_col], errors='coerce'),
            'MONTH': pd.to_numeric(balance_df[balance_month_col], errors='coerce'),
            'BALANCE_QTY': pd.to_numeric(balance_df[balance_qty_col], errors='coerce').fillna(0)
        })
        
        # Create sales dataframe - use Style_ID if available, otherwise SKU
        sales_clean = pd.DataFrame({
            'STYLE_ID': sales_df[sales_style_col].astype(str).str.strip(),
            'YEAR': pd.to_numeric(sales_df[sales_year_col], errors='coerce'),
            'MONTH': pd.to_numeric(sales_df[sales_month_col], errors='coerce'),
            'SALES_QTY': pd.to_numeric(sales_df[sales_qty_col], errors='coerce').fillna(0)
        })
        
        # Add additional columns from sales if they exist
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
            'FOB': ['FOB', 'Fob'],
            'SKU': ['SKU', 'Sku']
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
                    agg_dict[col] = 'first'  # Take first value for categorical columns
            
            sales_clean = sales_clean.groupby(['STYLE_ID', 'YEAR', 'MONTH'], as_index=False).agg(agg_dict)
        
        # Merge data on STYLE_ID, YEAR, MONTH
        merged_df = pd.merge(
            balance_clean,
            sales_clean,
            on=['STYLE_ID', 'YEAR', 'MONTH'],
            how='left',  # Keep all balance records
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
        
        # Data validation summary
        total_balance = merged_df['BALANCE_QTY'].sum()
        total_sales = merged_df['SALES_QTY'].sum()
        matched_records = len(merged_df[merged_df['SALES_QTY'] > 0])
        
        st.sidebar.success(f"""
        **Data Processing Complete:**
        - ✅ Total Balance Qty: {total_balance:,.0f}
        - ✅ Total Sales Qty: {total_sales:,.0f}
        - ✅ Matched Records: {matched_records:,} of {len(merged_df):,}
        - ✅ % Sold: {(total_sales/total_balance*100 if total_balance > 0 else 0):.1f}%
        """)
        
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
                    'SALES_QTY': 'sum'
                }).reset_index()
                
                # Calculate percentage
                grouped['PCT_SOLD'] = np.where(
                    grouped['BALANCE_QTY'] > 0,
                    (grouped['SALES_QTY'] / grouped['BALANCE_QTY']) * 100,
                    0
                )
                
                # Sort by sales descending
                grouped = grouped.sort_values('SALES_QTY', ascending=False)
                grouped.rename(columns={category_col: category_name}, inplace=True)
                
                return grouped
            
            # Identify which categorical columns are available
            available_categories = []
            category_options = {
                'Season': 'Season',
                'Subcategory': 'Subcategory', 
                'Color': 'Color',
                'Brand': 'Brand',
                'Heel_Type_1': 'Heel Type',
                'Maketplace': 'Marketplace'
            }
            
            for col, name in category_options.items():
                if col in filtered_df.columns:
                    available_categories.append((col, name))
            
            # Create charts for available categories
            if available_categories:
                st.markdown("### 📈 Sales Analysis by Category")
                
                # Create columns for charts
                num_charts = len(available_categories)
                if num_charts <= 3:
                    cols = st.columns(num_charts)
                else:
                    # First row with 3 columns
                    cols1 = st.columns(3)
                    # Second row with remaining columns
                    remaining = num_charts - 3
                    cols2 = st.columns(remaining) if remaining > 0 else []
                    cols = list(cols1) + list(cols2)
                
                # Create charts
                for idx, (col_name, display_name) in enumerate(available_categories):
                    if idx < len(cols):
                        with cols[idx]:
                            category_data = analyze_by_category(filtered_df, col_name, display_name)
                            if not category_data.empty:
                                st.markdown(f"#### {display_name}")
                                
                                # Create bar chart
                                fig = px.bar(category_data, x=display_name, y='SALES_QTY',
                                           color='SALES_QTY', 
                                           color_continuous_scale=['#FF6B6B', '#4ECDC4', '#45B7D1'][idx % 3])
                                fig.update_traces(
                                    texttemplate='%{y:.0f}', 
                                    textposition='outside'
                                )
                                fig.update_layout(
                                    height=400,
                                    showlegend=False,
                                    xaxis_title=display_name,
                                    yaxis_title="Sales Quantity",
                                    xaxis={'categoryorder': 'total descending'}
                                )
                                
                                # Add rangeslider if many categories
                                if len(category_data) > 15:
                                    fig.update_layout(xaxis_rangeslider_visible=True)
                                
                                st.plotly_chart(fig, use_container_width=True)
                                
                                # Data table
                                with st.expander(f"📋 {display_name} Data ({len(category_data)} items)"):
                                    display_df = category_data[[display_name, 'BALANCE_QTY', 'SALES_QTY', 'PCT_SOLD']].copy()
                                    display_df.columns = [display_name, 'Inventory Balance', 'Sales Quantity', '% Sold']
                                    
                                    st.dataframe(
                                        display_df.style.format({
                                            'Inventory Balance': '{:,.0f}',
                                            'Sales Quantity': '{:,.0f}',
                                            '% Sold': '{:.1f}%'
                                        }),
                                        hide_index=True,
                                        use_container_width=True
                                    )
                
                st.markdown("---")
            
            # Special Marketplace Trend Chart (if available)
            if 'Maketplace' in filtered_df.columns:
                st.markdown("### 📊 Marketplace Performance Trend")
                
                marketplace_data = analyze_by_category(filtered_df, 'Maketplace', 'Marketplace')
                if not marketplace_data.empty:
                    # Line chart
                    fig_market = go.Figure()
                    fig_market.add_trace(go.Scatter(
                        x=marketplace_data['Marketplace'],
                        y=marketplace_data['SALES_QTY'],
                        mode='lines+markers+text',
                        text=marketplace_data['SALES_QTY'],
                        textposition="top center",
                        line=dict(color='#1f77b4', width=3),
                        marker=dict(size=10, color='#ff7f0e')
                    ))
                    
                    fig_market.update_layout(
                        height=400,
                        xaxis_title="Marketplace",
                        yaxis_title="Total Sales Quantity",
                        hovermode='x unified',
                        template='plotly_white',
                        showlegend=False,
                        xaxis={'categoryorder': 'total descending'}
                    )
                    
                    fig_market.update_traces(
                        texttemplate='%{text:.0f}',
                        textfont=dict(size=12, color='black')
                    )
                    
                    st.plotly_chart(fig_market, use_container_width=True)
                    
                    # Marketplace data table
                    with st.expander(f"📋 Marketplace Data ({len(marketplace_data)} items)"):
                        display_df = marketplace_data[['Marketplace', 'BALANCE_QTY', 'SALES_QTY', 'PCT_SOLD']].copy()
                        display_df.columns = ['Marketplace', 'Inventory Balance', 'Sales Quantity', '% Sold']
                        st.dataframe(
                            display_df.style.format({
                                'Inventory Balance': '{:,.0f}',
                                'Sales Quantity': '{:,.0f}',
                                '% Sold': '{:.1f}%'
                            }),
                            hide_index=True,
                            use_container_width=True
                        )
            
            # Data validation section
            with st.expander("🔍 Data Validation Details"):
                st.write("**Sample Data (First 10 Rows):**")
                display_cols = ['STYLE_ID', 'YEAR', 'MONTH', 'BALANCE_QTY', 'SALES_QTY', 'PCT_SOLD']
                if 'SKU' in filtered_df.columns:
                    display_cols.insert(1, 'SKU')
                st.dataframe(filtered_df[display_cols].head(10), use_container_width=True)
                
                st.write("**Data Quality Check:**")
                quality_df = pd.DataFrame({
                    'Metric': [
                        'Total Records',
                        'Unique Style IDs', 
                        'Records with Sales > 0',
                        'Average Balance per Product',
                        'Average Sales per Product',
                        'Overall % Sold'
                    ],
                    'Value': [
                        len(filtered_df),
                        filtered_df['STYLE_ID'].nunique(),
                        (filtered_df['SALES_QTY'] > 0).sum(),
                        filtered_df['BALANCE_QTY'].mean(),
                        filtered_df['SALES_QTY'].mean(),
                        (filtered_df['SALES_QTY'].sum() / filtered_df['BALANCE_QTY'].sum() * 100) if filtered_df['BALANCE_QTY'].sum() > 0 else 0
                    ]
                })
                st.dataframe(quality_df, use_container_width=True)
                
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        import traceback
        st.write("Detailed error:", traceback.format_exc())
else:
    st.info("👆 Please upload an Excel file to begin analyzing your data.")
    
    # Instructions for Excel preparation
    with st.expander("📋 Your Excel File Structure"):
        st.markdown("""
        ### **Balance Sheet Columns:**
        - `Style_ID` - Product identifier
        - `YEAR` - Balance year (e.g., 2024)
        - `MONTH` - Balance month (1-12)
        - `Balance_QTY` - Available stock quantity
        
        ### **Sales Sheet Columns:**
        - `Style_ID` or `SKU` - Product identifier (must match Balance sheet Style_ID)
        - `YEAR` - Sales year (e.g., 2024)
        - `MONTH` - Sales month (1-12)
        - `Qty` - Quantity sold
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
        
        ### **Important Notes:**
        1. The Style_ID in Balance sheet must match Style_ID in Sales sheet
        2. If you have SKU in Sales sheet, it will be kept as additional information
        3. Year and Month columns are used to match sales with balance
        4. The dashboard will automatically handle any variations in capitalization
        """)
