import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(page_title="Sales Analytics Dashboard", layout="wide", page_icon="📊")

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
    .dataframe th {
        cursor: pointer;
        background-color: #f0f2f6;
    }
    .dataframe th:hover {
        background-color: #e0e2e6;
    }
    .table-container {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

# Title
st.markdown("<h1 style='text-align: center; color: #1f77b4;'>📊 Sales Analytics Dashboard</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center; color: #666;'>Sales Performance Analysis</h3>", unsafe_allow_html=True)
st.markdown("---")

# File uploader
uploaded_file = st.file_uploader("Upload Excel File with 'Sales' sheet", type=['xlsx', 'xls'])

@st.cache_data(ttl=3600)
def load_and_process_data(uploaded_file):
    """Load and process the Excel file with sales data only"""
    try:
        # Read only sales sheet
        sales_df = pd.read_excel(uploaded_file, sheet_name='Sales')
        
        # Display raw data info
        st.sidebar.info(f"""
        **Raw Data Loaded:**
        - 📊 Sales Records: {len(sales_df):,}
        """)
        
        # Clean column names - preserve original case but strip spaces
        def clean_columns(df):
            df.columns = df.columns.astype(str).str.strip()
            return df
        
        # Clean column names
        sales_df = clean_columns(sales_df)
        
        # Show column names for verification
        with st.sidebar.expander("🔍 Verify Column Names"):
            st.write("**Sales Columns:**", list(sales_df.columns))
        
        # Function to find column with case-insensitive matching
        def find_column(df, possible_names):
            df_cols_upper = {col.upper(): col for col in df.columns}
            for name in possible_names:
                if name.upper() in df_cols_upper:
                    return df_cols_upper[name.upper()]
            return None
        
        # Identify key columns for SALES sheet
        sales_style_col = find_column(sales_df, ['Style_ID', 'STYLE_ID', 'StyleID', 'SKU'])
        sales_year_col = find_column(sales_df, ['YEAR', 'Year'])
        sales_month_col = find_column(sales_df, ['MONTH', 'Month'])
        sales_qty_col = find_column(sales_df, ['Qty', 'QTY', 'sales Qty', 'sales_Qty', 'Sales_Qty', 'Sales Qty', 'Quantity', 'Sales_QTY'])
        
        # Verify required columns exist
        required_cols_sales = {
            'Style': sales_style_col,
            'Year': sales_year_col,
            'Month': sales_month_col,
            'Sales Qty': sales_qty_col
        }
        
        missing_sales = [k for k, v in required_cols_sales.items() if v is None]
        
        if missing_sales:
            st.error(f"❌ Missing required columns in Sales sheet: {', '.join(missing_sales)}")
            st.info("Available columns: " + ", ".join(sales_df.columns))
            st.stop()
        
        # Create clean dataframe with standardized names
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
            'Closing_stock': ['Closing_stock', 'Closing Stock', 'CLOSING_STOCK'],
            'Date': ['Date', 'DATE']
        }
        
        for standard_name, possible_names in additional_cols_mapping.items():
            found_col = find_column(sales_df, possible_names)
            if found_col:
                sales_clean[standard_name] = sales_df[found_col]
        
        # Handle duplicate sales records - FIXED VERSION
        # First, build the subset for duplicate checking
        duplicate_subset = ['STYLE_ID', 'YEAR', 'MONTH']
        
        # Check if Maketplace column exists and add it to subset
        if 'Maketplace' in sales_clean.columns:
            duplicate_subset.append('Maketplace')
        
        # Now check for duplicates with the correct subset
        duplicate_check = sales_clean.duplicated(subset=duplicate_subset, keep=False).sum()
        
        if duplicate_check > 0:
            st.sidebar.warning(f"⚠️ Found {duplicate_check} duplicate sales records. Aggregating...")
            
            # Get list of columns to aggregate
            agg_dict = {'SALES_QTY': 'sum'}
            for col in sales_clean.columns:
                if col not in duplicate_subset + ['SALES_QTY']:
                    agg_dict[col] = 'first'  # Take first value for categorical columns
            
            sales_clean = sales_clean.groupby(duplicate_subset, as_index=False).agg(agg_dict)
        
        # Add month name for display
        month_names = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 
                      6: 'June', 7: 'July', 8: 'August', 9: 'September', 
                      10: 'October', 11: 'November', 12: 'December'}
        sales_clean['MONTH_NAME'] = sales_clean['MONTH'].map(month_names)
        
        # Data validation summary
        total_sales = sales_clean['SALES_QTY'].sum()
        
        st.sidebar.success(f"""
        **Data Processing Complete:**
        - ✅ Total Sales Qty: {total_sales:,.0f}
        - ✅ Unique Products: {sales_clean['STYLE_ID'].nunique():,}
        - ✅ Time Period: {sales_clean['YEAR'].min()} - {sales_clean['YEAR'].max()}
        """)
        
        return sales_clean
        
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        import traceback
        st.write("Detailed error:", traceback.format_exc())
        st.stop()

if uploaded_file is not None:
    try:
        # Load and process data
        with st.spinner('🔍 Loading and processing data...'):
            df = load_and_process_data(uploaded_file)
        
        # Display success message
        st.success(f"✅ Data loaded successfully! {len(df):,} records processed")
        
        # Data summary metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            total_sales = df['SALES_QTY'].sum()
            st.metric("Total Sales", f"{total_sales:,.0f}")
        with col2:
            st.metric("Unique Products", f"{df['STYLE_ID'].nunique():,.0f}")
        with col3:
            st.metric("Time Period", f"{df['YEAR'].min()} - {df['YEAR'].max()}")
        with col4:
            avg_monthly_sales = df.groupby(['YEAR', 'MONTH'])['SALES_QTY'].sum().mean()
            st.metric("Avg Monthly Sales", f"{avg_monthly_sales:,.0f}")
        
        st.markdown("---")
        
        # Sidebar filters - ONLY Year and Month
        st.sidebar.header("🔍 Filter Options")
        
        # Get unique years and months
        years = sorted([int(y) for y in df['YEAR'].dropna().unique() if not pd.isna(y)])
        months = sorted([int(m) for m in df['MONTH'].dropna().unique() if not pd.isna(m)])
        month_names = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 
                      6: 'June', 7: 'July', 8: 'August', 9: 'September', 
                      10: 'October', 11: 'November', 12: 'December'}
        
        selected_year = st.sidebar.selectbox("Select Year", ['All'] + years)
        selected_month = st.sidebar.selectbox("Select Month", ['All'] + [month_names[m] for m in months if m in month_names])
        
        # Filter data
        filtered_df = df.copy()
        
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
        - 💰 Sales: {filtered_df['SALES_QTY'].sum():,.0f}
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
                    'SALES_QTY': 'sum'
                }).reset_index()
                
                # Sort by sales descending by default
                grouped = grouped.sort_values('SALES_QTY', ascending=False)
                grouped.rename(columns={category_col: category_name}, inplace=True)
                
                return grouped
            
            # Marketplace Bar Chart
            if 'Maketplace' in filtered_df.columns:
                st.markdown("### 📊 Marketplace Performance")
                
                # Group by marketplace
                marketplace_data = analyze_by_category(filtered_df, 'Maketplace', 'Marketplace')
                
                if not marketplace_data.empty:
                    # Create bar chart
                    fig_marketplace = px.bar(
                        marketplace_data, 
                        x='Marketplace', 
                        y='SALES_QTY',
                        color='SALES_QTY',
                        color_continuous_scale='viridis',
                        title=f"Sales by Marketplace (Filtered: {selected_year}, {selected_month})",
                        text='SALES_QTY'
                    )
                    
                    fig_marketplace.update_traces(
                        texttemplate='%{text:,.0f}',
                        textposition='outside'
                    )
                    
                    fig_marketplace.update_layout(
                        height=500,
                        showlegend=False,
                        xaxis_title="Marketplace",
                        yaxis_title="Sales Quantity",
                        xaxis={'categoryorder': 'total descending', 'tickangle': 45},
                        title_x=0.5,
                        hovermode='x unified'
                    )
                    
                    st.plotly_chart(fig_marketplace, use_container_width=True)
                    
                    # Marketplace data table
                    with st.expander("📋 Marketplace Data Table"):
                        market_table = marketplace_data.copy()
                        market_table.columns = ['Marketplace', 'Sales Quantity']
                        market_table['Sales Quantity'] = market_table['Sales Quantity'].apply(lambda x: f"{x:,.0f}")
                        st.dataframe(market_table, hide_index=True, use_container_width=True)
                
                st.markdown("---")
            
            # Category Analysis Tables
            # Identify which categorical columns are available
            available_categories = []
            category_options = {
                'Season': 'Season',
                'Subcategory': 'Subcategory', 
                'Color': 'Color',
                'Brand': 'Brand',
                'Heel_Type_1': 'Heel Type'
            }
            
            for col, name in category_options.items():
                if col in filtered_df.columns:
                    available_categories.append((col, name))
            
            if available_categories:
                st.markdown("### 📈 Sales Analysis by Category")
                
                # Display tables in rows of 2
                for i in range(0, len(available_categories), 2):
                    cols = st.columns(2)
                    
                    for j in range(2):
                        if i + j < len(available_categories):
                            col_name, display_name = available_categories[i + j]
                            
                            with cols[j]:
                                st.markdown(f"<div class='table-container'>", unsafe_allow_html=True)
                                st.markdown(f"#### {display_name}")
                                
                                category_data = analyze_by_category(filtered_df, col_name, display_name)
                                
                                if not category_data.empty:
                                    # Create table
                                    category_table = category_data.copy()
                                    category_table.columns = [display_name, 'Sales Quantity']
                                    
                                    # Display table with sorting
                                    st.dataframe(
                                        category_table,
                                        hide_index=True,
                                        use_container_width=True,
                                        height=300
                                    )
                                    
                                    # Show summary below table
                                    total_sales_cat = category_table['Sales Quantity'].sum()
                                    st.caption(f"Total: {total_sales_cat:,.0f} units")
                                else:
                                    st.info(f"No data available for {display_name}")
                                
                                st.markdown(f"</div>", unsafe_allow_html=True)
                
                st.markdown("---")
            
            # Monthly Trend Chart
            st.markdown("### 📅 Monthly Sales Trend")
            
            # Group by month for trend analysis
            monthly_data = filtered_df.groupby(['YEAR', 'MONTH', 'MONTH_NAME']).agg({
                'SALES_QTY': 'sum'
            }).reset_index()
            
            # Sort by year and month
            monthly_data = monthly_data.sort_values(['YEAR', 'MONTH'])
            
            # Create X-axis labels
            monthly_data['Period'] = monthly_data['MONTH_NAME'] + ' ' + monthly_data['YEAR'].astype(str)
            
            # Line chart for trend visualization
            fig_monthly = go.Figure()
            fig_monthly.add_trace(go.Scatter(
                x=monthly_data['Period'],
                y=monthly_data['SALES_QTY'],
                mode='lines+markers+text',
                text=monthly_data['SALES_QTY'],
                textposition="top center",
                line=dict(color='#1f77b4', width=3),
                marker=dict(size=10, color='#ff7f0e'),
                name='Sales Quantity'
            ))
            
            fig_monthly.update_layout(
                height=500,
                xaxis_title="Month",
                yaxis_title="Sales Quantity",
                hovermode='x unified',
                template='plotly_white',
                showlegend=False,
                xaxis={'tickangle': 45}
            )
            
            fig_monthly.update_traces(
                texttemplate='%{text:,.0f}',
                textfont=dict(size=10, color='black')
            )
            
            st.plotly_chart(fig_monthly, use_container_width=True)
            
            # Monthly data table with sorting
            with st.expander("📋 Monthly Trend Data Table"):
                trend_table = monthly_data[['Period', 'SALES_QTY']].copy()
                trend_table.columns = ['Period', 'Sales Quantity']
                st.dataframe(trend_table, hide_index=True, use_container_width=True)
            
            st.markdown("---")
            
            # Top Products Analysis
            st.markdown("### 🏆 Top Products Analysis")
            
            # Group by product
            product_data = filtered_df.groupby('STYLE_ID').agg({
                'SALES_QTY': 'sum'
            }).reset_index()
            
            # Sort by sales descending by default
            product_data = product_data.sort_values('SALES_QTY', ascending=False)
            
            # Display in 2 columns: Top 10 and Complete Table
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("<div class='table-container'>", unsafe_allow_html=True)
                st.markdown("#### Top 10 Products")
                
                top_10_products = product_data.head(10).copy()
                top_10_products.columns = ['Style ID', 'Sales Quantity']
                
                st.dataframe(
                    top_10_products,
                    hide_index=True,
                    use_container_width=True,
                    height=400
                )
                st.markdown("</div>", unsafe_allow_html=True)
            
            with col2:
                st.markdown("<div class='table-container'>", unsafe_allow_html=True)
                st.markdown("#### Complete Products List")
                
                all_products = product_data.copy()
                all_products.columns = ['Style ID', 'Sales Quantity']
                
                st.dataframe(
                    all_products,
                    hide_index=True,
                    use_container_width=True,
                    height=400
                )
                
                total_products = len(all_products)
                st.caption(f"Total Products: {total_products}")
                st.markdown("</div>", unsafe_allow_html=True)
            
            st.markdown("---")
            
            # Data validation section
            with st.expander("🔍 Data Validation Details"):
                st.write("**Sample Data (First 10 Rows):**")
                display_cols = ['STYLE_ID', 'YEAR', 'MONTH', 'SALES_QTY']
                if 'Maketplace' in filtered_df.columns:
                    display_cols.append('Maketplace')
                if 'Subcategory' in filtered_df.columns:
                    display_cols.append('Subcategory')
                if 'Brand' in filtered_df.columns:
                    display_cols.append('Brand')
                
                # Display sample data
                sample_df = filtered_df[display_cols].head(10).copy()
                sample_df.columns = ['Style ID', 'Year', 'Month', 'Sales Quantity'] + \
                                   (['Marketplace'] if 'Maketplace' in filtered_df.columns else []) + \
                                   (['Subcategory'] if 'Subcategory' in filtered_df.columns else []) + \
                                   (['Brand'] if 'Brand' in filtered_df.columns else [])
                
                st.dataframe(sample_df, use_container_width=True)
                
                st.write("**Data Quality Check:**")
                quality_df = pd.DataFrame({
                    'Metric': [
                        'Total Records',
                        'Unique Products', 
                        'Records with Sales > 0',
                        'Average Sales per Record',
                        'Maximum Sales (Single Record)',
                        'Time Period Covered'
                    ],
                    'Value': [
                        len(filtered_df),
                        filtered_df['STYLE_ID'].nunique(),
                        (filtered_df['SALES_QTY'] > 0).sum(),
                        f"{filtered_df['SALES_QTY'].mean():.0f}",
                        f"{filtered_df['SALES_QTY'].max():,.0f}",
                        f"{filtered_df['YEAR'].min()} - {filtered_df['YEAR'].max()}"
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
    with st.expander("📋 Required Excel File Structure"):
        st.markdown("""
        ### **Sales Sheet Columns (Required):**
        - `Style_ID` - Product identifier
        - `YEAR` - Sales year (e.g., 2024)
        - `MONTH` - Sales month (1-12)
        - `Qty` - Quantity sold during period
        
        ### **Optional Columns:**
        - `Maketplace` - Sales marketplace
        - `Subcategory` - Product subcategory
        - `Heel_Type 1` - Type of heel
        - `Season` - Product season
        - `Brand` - Product brand
        - `Color` - Product color
        
        ### **Key Features:**
        - **Marketplace Bar Chart** - Visual sales by marketplace
        - **Category Tables** - Displayed side by side (2 per row)
        - **Monthly Trend Chart** - Line chart showing sales over time
        - **Products Analysis** - Top 10 and complete list side by side
        - **Interactive Sorting** - Click column headers to sort
        - **Year/Month Filters** - Filter data by time period
        """)
