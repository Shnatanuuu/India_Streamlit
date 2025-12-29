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
    .positive-percent {
        color: green;
        font-weight: bold;
    }
    .negative-percent {
        color: red;
        font-weight: bold;
    }
    .high-percent {
        background-color: #ffcccc;
    }
    .medium-percent {
        background-color: #ffffcc;
    }
    .low-percent {
        background-color: #ccffcc;
    }
    </style>
    """, unsafe_allow_html=True)

# Title
st.markdown("<h1 style='text-align: center; color: #1f77b4;'>📊 Sales Analytics Dashboard</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center; color: #666;'>Sales Performance Analysis with Stock Metrics</h3>", unsafe_allow_html=True)
st.markdown("---")

# File uploader
uploaded_file = st.file_uploader("Upload Excel File with 'Sales' sheet", type=['xlsx', 'xls'])

@st.cache_data(ttl=3600)
def load_and_process_data(uploaded_file):
    """Load and process the Excel file with sales data including opening stock"""
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
        sales_style_col = find_column(sales_df, ['Style_ID', 'STYLE_ID', 'StyleID', 'SKU', 'STYLE CODE', 'STYLE-CODE', 'Style Code'])
        sales_year_col = find_column(sales_df, ['YEAR', 'Year', 'SALES_YEAR', 'Sale Year'])
        sales_month_col = find_column(sales_df, ['MONTH', 'Month', 'SALES_MONTH', 'Sale Month'])
        sales_qty_col = find_column(sales_df, ['Qty', 'QTY', 'sales Qty', 'sales_Qty', 'Sales_Qty', 'Sales Qty', 'Quantity', 'Sales_QTY', 'SALES'])
        opening_stock_col = find_column(sales_df, ['Opening_stock', 'Opening Stock', 'OPENING_STOCK', 'Opening_Stock', 
                                                  'opening stock', 'OpeningStock', 'OP_STOCK', 'Opening_Stock_Qty'])
        
        # Verify required columns exist
        required_cols_sales = {
            'Style': sales_style_col,
            'Year': sales_year_col,
            'Month': sales_month_col,
            'Sales Qty': sales_qty_col,
            'Opening Stock': opening_stock_col
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
            'SALES_QTY': pd.to_numeric(sales_df[sales_qty_col], errors='coerce').fillna(0),
            'OPENING_STOCK': pd.to_numeric(sales_df[opening_stock_col], errors='coerce').fillna(0)
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
            agg_dict = {'SALES_QTY': 'sum', 'OPENING_STOCK': 'first'}
            for col in sales_clean.columns:
                if col not in duplicate_subset + ['SALES_QTY', 'OPENING_STOCK']:
                    agg_dict[col] = 'first'  # Take first value for categorical columns
            
            sales_clean = sales_clean.groupby(duplicate_subset, as_index=False).agg(agg_dict)
        
        # Add month name for display
        month_names = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 
                      6: 'June', 7: 'July', 8: 'August', 9: 'September', 
                      10: 'October', 11: 'November', 12: 'December'}
        sales_clean['MONTH_NAME'] = sales_clean['MONTH'].map(month_names)
        
        # Calculate sales percentage (Sales Qty / Opening Stock)
        # Handle division by zero and cases where opening stock is 0
        sales_clean['SALES_PERCENTAGE'] = np.where(
            sales_clean['OPENING_STOCK'] > 0,
            (sales_clean['SALES_QTY'] / sales_clean['OPENING_STOCK']) * 100,
            0
        )
        
        # Add a column for sales efficiency classification
        def classify_sales_percentage(percent):
            if percent == 0:
                return 'No Opening Stock'
            elif percent <= 30:
                return 'Low (<30%)'
            elif percent <= 60:
                return 'Medium (30-60%)'
            elif percent <= 100:
                return 'High (60-100%)'
            else:
                return 'Very High (>100%)'
        
        sales_clean['SALES_EFFICIENCY'] = sales_clean['SALES_PERCENTAGE'].apply(classify_sales_percentage)
        
        # Data validation summary
        total_sales = sales_clean['SALES_QTY'].sum()
        total_opening_stock = sales_clean['OPENING_STOCK'].sum()
        avg_sales_percentage = sales_clean[sales_clean['OPENING_STOCK'] > 0]['SALES_PERCENTAGE'].mean()
        
        st.sidebar.success(f"""
        **Data Processing Complete:**
        - ✅ Total Sales Qty: {total_sales:,.0f}
        - ✅ Total Opening Stock: {total_opening_stock:,.0f}
        - ✅ Avg Sales %: {avg_sales_percentage:.1f}%
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
            total_opening_stock = df['OPENING_STOCK'].sum()
            st.metric("Total Opening Stock", f"{total_opening_stock:,.0f}")
        with col3:
            avg_sales_percentage = df[df['OPENING_STOCK'] > 0]['SALES_PERCENTAGE'].mean()
            st.metric("Avg Sales %", f"{avg_sales_percentage:.1f}%", 
                     delta=f"{avg_sales_percentage - 50:.1f}%" if avg_sales_percentage else "N/A")
        with col4:
            st.metric("Unique Products", f"{df['STYLE_ID'].nunique():,.0f}")
        
        # Additional metrics
        col5, col6, col7, col8 = st.columns(4)
        with col5:
            sales_ratio = (df['SALES_QTY'].sum() / df[df['OPENING_STOCK'] > 0]['OPENING_STOCK'].sum() * 100) \
                if df[df['OPENING_STOCK'] > 0]['OPENING_STOCK'].sum() > 0 else 0
            st.metric("Overall Sales %", f"{sales_ratio:.1f}%")
        with col6:
            high_efficiency = len(df[df['SALES_PERCENTAGE'] > 60])
            st.metric("High Efficiency (>60%)", f"{high_efficiency:,}")
        with col7:
            avg_monthly_sales = df.groupby(['YEAR', 'MONTH'])['SALES_QTY'].sum().mean()
            st.metric("Avg Monthly Sales", f"{avg_monthly_sales:,.0f}")
        with col8:
            time_period = f"{df['YEAR'].min()} - {df['YEAR'].max()}"
            st.metric("Time Period", time_period)
        
        st.markdown("---")
        
        # Sidebar filters
        st.sidebar.header("🔍 Filter Options")
        
        # Get unique years and months
        years = sorted([int(y) for y in df['YEAR'].dropna().unique() if not pd.isna(y)])
        months = sorted([int(m) for m in df['MONTH'].dropna().unique() if not pd.isna(m)])
        month_names_dict = {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 
                          6: 'June', 7: 'July', 8: 'August', 9: 'September', 
                          10: 'October', 11: 'November', 12: 'December'}
        
        selected_year = st.sidebar.selectbox("Select Year", ['All'] + years)
        selected_month = st.sidebar.selectbox("Select Month", ['All'] + [month_names_dict[m] for m in months if m in month_names_dict])
        
        # Add sales efficiency filter
        efficiency_options = ['All', 'Very High (>100%)', 'High (60-100%)', 'Medium (30-60%)', 'Low (<30%)', 'No Opening Stock']
        selected_efficiency = st.sidebar.selectbox("Sales Efficiency", efficiency_options)
        
        # Filter data
        filtered_df = df.copy()
        
        if selected_year != 'All':
            filtered_df = filtered_df[filtered_df['YEAR'] == selected_year]
        
        if selected_month != 'All':
            month_num = [k for k, v in month_names_dict.items() if v == selected_month][0]
            filtered_df = filtered_df[filtered_df['MONTH'] == month_num]
        
        if selected_efficiency != 'All':
            filtered_df = filtered_df[filtered_df['SALES_EFFICIENCY'] == selected_efficiency]
        
        # Display filter summary
        st.sidebar.markdown("---")
        st.sidebar.info(f"""
        **Filter Applied:**
        - 📅 Year: {selected_year}
        - 📆 Month: {selected_month}
        - 📈 Efficiency: {selected_efficiency}
        - 📊 Records: {len(filtered_df):,}
        - 💰 Sales: {filtered_df['SALES_QTY'].sum():,.0f}
        - 📦 Opening Stock: {filtered_df['OPENING_STOCK'].sum():,.0f}
        - 📊 Sales %: {filtered_df[filtered_df['OPENING_STOCK'] > 0]['SALES_PERCENTAGE'].mean():.1f}%
        """)
        
        if len(filtered_df) == 0:
            st.warning("⚠️ No data available for the selected filters.")
        else:
            # Helper function for grouped analysis with stock metrics
            def analyze_with_stock(df, group_col, group_name):
                if group_col not in df.columns:
                    return pd.DataFrame()
                
                # Group by category with stock metrics
                grouped = df.groupby(group_col, observed=True).agg({
                    'SALES_QTY': 'sum',
                    'OPENING_STOCK': 'sum'
                }).reset_index()
                
                # Calculate sales percentage
                grouped['SALES_PERCENTAGE'] = np.where(
                    grouped['OPENING_STOCK'] > 0,
                    (grouped['SALES_QTY'] / grouped['OPENING_STOCK']) * 100,
                    0
                )
                
                # Sort by sales percentage descending by default
                grouped = grouped.sort_values('SALES_PERCENTAGE', ascending=False)
                grouped.rename(columns={group_col: group_name}, inplace=True)
                
                return grouped
            
            # Marketplace Bar Chart with Stock Metrics
            if 'Maketplace' in filtered_df.columns:
                st.markdown("### 📊 Marketplace Performance with Stock Metrics")
                
                # Group by marketplace
                marketplace_data = analyze_with_stock(filtered_df, 'Maketplace', 'Marketplace')
                
                if not marketplace_data.empty:
                    # Create two bar charts side by side
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Sales Quantity chart
                        fig_sales = px.bar(
                            marketplace_data, 
                            x='Marketplace', 
                            y='SALES_QTY',
                            color='SALES_QTY',
                            color_continuous_scale='viridis',
                            title=f"Sales Quantity by Marketplace",
                            text='SALES_QTY'
                        )
                        fig_sales.update_traces(
                            texttemplate='%{text:,.0f}',
                            textposition='outside'
                        )
                        fig_sales.update_layout(
                            height=400,
                            showlegend=False,
                            xaxis_title="Marketplace",
                            yaxis_title="Sales Quantity",
                            xaxis={'categoryorder': 'total descending', 'tickangle': 45},
                            title_x=0.5,
                            hovermode='x unified'
                        )
                        st.plotly_chart(fig_sales, use_container_width=True)
                    
                    with col2:
                        # Sales Percentage chart
                        fig_percent = px.bar(
                            marketplace_data, 
                            x='Marketplace', 
                            y='SALES_PERCENTAGE',
                            color='SALES_PERCENTAGE',
                            color_continuous_scale='RdYlGn',
                            title=f"Sales % by Marketplace",
                            text='SALES_PERCENTAGE'
                        )
                        fig_percent.update_traces(
                            texttemplate='%{text:.1f}%',
                            textposition='outside'
                        )
                        fig_percent.update_layout(
                            height=400,
                            showlegend=False,
                            xaxis_title="Marketplace",
                            yaxis_title="Sales Percentage (%)",
                            xaxis={'categoryorder': 'total descending', 'tickangle': 45},
                            title_x=0.5,
                            hovermode='x unified'
                        )
                        st.plotly_chart(fig_percent, use_container_width=True)
                    
                    # Marketplace data table with all metrics
                    with st.expander("📋 Marketplace Data Table with Stock Metrics"):
                        market_table = marketplace_data.copy()
                        market_table.columns = ['Marketplace', 'Sales Quantity', 'Opening Stock', 'Sales %']
                        market_table['Sales Quantity'] = market_table['Sales Quantity'].apply(lambda x: f"{x:,.0f}")
                        market_table['Opening Stock'] = market_table['Opening Stock'].apply(lambda x: f"{x:,.0f}")
                        market_table['Sales %'] = market_table['Sales %'].apply(lambda x: f"{x:.1f}%")
                        
                        # Apply conditional formatting
                        def highlight_sales_percentage(val):
                            try:
                                percent = float(val.replace('%', '').replace(',', ''))
                                if percent > 100:
                                    return 'background-color: #ffcccc'
                                elif percent > 60:
                                    return 'background-color: #ffffcc'
                                else:
                                    return 'background-color: #ccffcc'
                            except:
                                return ''
                        
                        styled_table = market_table.style.applymap(highlight_sales_percentage, subset=['Sales %'])
                        st.dataframe(styled_table, hide_index=True, use_container_width=True)
                
                st.markdown("---")
            
            # Category Analysis Tables with Stock Metrics
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
                st.markdown("### 📈 Sales Analysis by Category with Stock Metrics")
                
                # Display tables in rows of 2
                for i in range(0, len(available_categories), 2):
                    cols = st.columns(2)
                    
                    for j in range(2):
                        if i + j < len(available_categories):
                            col_name, display_name = available_categories[i + j]
                            
                            with cols[j]:
                                st.markdown(f"<div class='table-container'>", unsafe_allow_html=True)
                                st.markdown(f"#### {display_name}")
                                
                                category_data = analyze_with_stock(filtered_df, col_name, display_name)
                                
                                if not category_data.empty:
                                    # Create table with all metrics
                                    category_table = category_data.copy()
                                    category_table.columns = [display_name, 'Sales Qty', 'Opening Stock', 'Sales %']
                                    
                                    # Format the table
                                    display_table = category_table.copy()
                                    display_table['Sales Qty'] = display_table['Sales Qty'].apply(lambda x: f"{x:,.0f}")
                                    display_table['Opening Stock'] = display_table['Opening Stock'].apply(lambda x: f"{x:,.0f}")
                                    display_table['Sales %'] = display_table['Sales %'].apply(lambda x: f"{x:.1f}%")
                                    
                                    # Display table with sorting
                                    st.dataframe(
                                        display_table,
                                        hide_index=True,
                                        use_container_width=True,
                                        height=300
                                    )
                                    
                                    # Show summary below table
                                    total_sales_cat = category_table['Sales Qty'].sum()
                                    total_stock_cat = category_table['Opening Stock'].sum()
                                    avg_percent_cat = category_table[category_table['Opening Stock'] > 0]['Sales %'].mean()
                                    
                                    st.caption(f"**Summary:** Sales: {total_sales_cat:,.0f} | Stock: {total_stock_cat:,.0f} | Avg %: {avg_percent_cat:.1f}%")
                                else:
                                    st.info(f"No data available for {display_name}")
                                
                                st.markdown(f"</div>", unsafe_allow_html=True)
                
                st.markdown("---")
            
            # Monthly Trend Chart with Stock Metrics
            st.markdown("### 📅 Monthly Sales Trend with Stock Metrics")
            
            # Group by month for trend analysis
            monthly_data = filtered_df.groupby(['YEAR', 'MONTH', 'MONTH_NAME']).agg({
                'SALES_QTY': 'sum',
                'OPENING_STOCK': 'sum'
            }).reset_index()
            
            # Calculate sales percentage for monthly data
            monthly_data['SALES_PERCENTAGE'] = np.where(
                monthly_data['OPENING_STOCK'] > 0,
                (monthly_data['SALES_QTY'] / monthly_data['OPENING_STOCK']) * 100,
                0
            )
            
            # Sort by year and month
            monthly_data = monthly_data.sort_values(['YEAR', 'MONTH'])
            
            # Create X-axis labels
            monthly_data['Period'] = monthly_data['MONTH_NAME'] + ' ' + monthly_data['YEAR'].astype(str)
            
            # Create dual-axis chart
            fig_monthly = go.Figure()
            
            # Add Sales Quantity bars
            fig_monthly.add_trace(go.Bar(
                x=monthly_data['Period'],
                y=monthly_data['SALES_QTY'],
                name='Sales Quantity',
                marker_color='#1f77b4',
                yaxis='y1'
            ))
            
            # Add Sales Percentage line
            fig_monthly.add_trace(go.Scatter(
                x=monthly_data['Period'],
                y=monthly_data['SALES_PERCENTAGE'],
                name='Sales %',
                mode='lines+markers',
                line=dict(color='#ff7f0e', width=3),
                marker=dict(size=10),
                yaxis='y2'
            ))
            
            # Update layout for dual y-axes
            fig_monthly.update_layout(
                height=500,
                xaxis_title="Month",
                yaxis_title="Sales Quantity",
                yaxis2=dict(
                    title="Sales Percentage (%)",
                    overlaying='y',
                    side='right',
                    range=[0, max(monthly_data['SALES_PERCENTAGE'].max() * 1.2, 100)],
                    tickformat='.0f%'
                ),
                hovermode='x unified',
                template='plotly_white',
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                xaxis={'tickangle': 45}
            )
            
            st.plotly_chart(fig_monthly, use_container_width=True)
            
            # Monthly data table with all metrics
            with st.expander("📋 Monthly Trend Data Table with Stock Metrics"):
                trend_table = monthly_data[['Period', 'SALES_QTY', 'OPENING_STOCK', 'SALES_PERCENTAGE']].copy()
                trend_table.columns = ['Period', 'Sales Quantity', 'Opening Stock', 'Sales %']
                trend_table['Sales Quantity'] = trend_table['Sales Quantity'].apply(lambda x: f"{x:,.0f}")
                trend_table['Opening Stock'] = trend_table['Opening Stock'].apply(lambda x: f"{x:,.0f}")
                trend_table['Sales %'] = trend_table['Sales %'].apply(lambda x: f"{x:.1f}%")
                st.dataframe(trend_table, hide_index=True, use_container_width=True)
            
            st.markdown("---")
            
            # Top Products Analysis with Stock Metrics
            st.markdown("### 🏆 Top Products Analysis with Stock Metrics")
            
            # Group by product with stock metrics
            product_data = filtered_df.groupby('STYLE_ID').agg({
                'SALES_QTY': 'sum',
                'OPENING_STOCK': 'sum'
            }).reset_index()
            
            # Calculate sales percentage for products
            product_data['SALES_PERCENTAGE'] = np.where(
                product_data['OPENING_STOCK'] > 0,
                (product_data['SALES_QTY'] / product_data['OPENING_STOCK']) * 100,
                0
            )
            
            # Sort options
            sort_option = st.radio(
                "Sort products by:",
                ["Sales Percentage (Highest)", "Sales Quantity (Highest)", "Opening Stock (Highest)", "Sales Percentage (Lowest)"],
                horizontal=True,
                key="product_sort"
            )
            
            if "Sales Percentage" in sort_option and "Highest" in sort_option:
                product_data = product_data.sort_values('SALES_PERCENTAGE', ascending=False)
            elif "Sales Quantity" in sort_option:
                product_data = product_data.sort_values('SALES_QTY', ascending=False)
            elif "Opening Stock" in sort_option:
                product_data = product_data.sort_values('OPENING_STOCK', ascending=False)
            else:
                product_data = product_data.sort_values('SALES_PERCENTAGE', ascending=True)
            
            # Display in 2 columns: Top Products and Complete Table
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("<div class='table-container'>", unsafe_allow_html=True)
                st.markdown(f"#### Top 10 Products ({sort_option})")
                
                top_10_products = product_data.head(10).copy()
                top_10_products.columns = ['Style ID', 'Sales Qty', 'Opening Stock', 'Sales %']
                
                # Format for display
                display_top = top_10_products.copy()
                display_top['Sales Qty'] = display_top['Sales Qty'].apply(lambda x: f"{x:,.0f}")
                display_top['Opening Stock'] = display_top['Opening Stock'].apply(lambda x: f"{x:,.0f}")
                display_top['Sales %'] = display_top['Sales %'].apply(lambda x: f"{x:.1f}%")
                
                st.dataframe(
                    display_top,
                    hide_index=True,
                    use_container_width=True,
                    height=400
                )
                st.markdown("</div>", unsafe_allow_html=True)
            
            with col2:
                st.markdown("<div class='table-container'>", unsafe_allow_html=True)
                st.markdown(f"#### Complete Products List ({sort_option})")
                
                all_products = product_data.copy()
                all_products.columns = ['Style ID', 'Sales Qty', 'Opening Stock', 'Sales %']
                
                # Format for display
                display_all = all_products.copy()
                display_all['Sales Qty'] = display_all['Sales Qty'].apply(lambda x: f"{x:,.0f}")
                display_all['Opening Stock'] = display_all['Opening Stock'].apply(lambda x: f"{x:,.0f}")
                display_all['Sales %'] = display_all['Sales %'].apply(lambda x: f"{x:.1f}%")
                
                st.dataframe(
                    display_all,
                    hide_index=True,
                    use_container_width=True,
                    height=400
                )
                
                total_products = len(all_products)
                total_sales_all = all_products['Sales Qty'].sum()
                total_stock_all = all_products['Opening Stock'].sum()
                avg_percent_all = all_products[all_products['Opening Stock'] > 0]['Sales %'].mean()
                
                st.caption(f"**Summary:** Products: {total_products} | Sales: {total_sales_all:,.0f} | Stock: {total_stock_all:,.0f} | Avg %: {avg_percent_all:.1f}%")
                st.markdown("</div>", unsafe_allow_html=True)
            
            # Sales Efficiency Distribution Pie Chart
            st.markdown("### 📊 Sales Efficiency Distribution")
            
            efficiency_data = filtered_df.groupby('SALES_EFFICIENCY').agg({
                'STYLE_ID': 'nunique',
                'SALES_QTY': 'sum'
            }).reset_index()
            
            col1, col2 = st.columns(2)
            
            with col1:
                fig_pie = px.pie(
                    efficiency_data,
                    names='SALES_EFFICIENCY',
                    values='STYLE_ID',
                    title='Products by Sales Efficiency',
                    color='SALES_EFFICIENCY',
                    color_discrete_map={
                        'Very High (>100%)': '#ff0000',
                        'High (60-100%)': '#ff9900',
                        'Medium (30-60%)': '#ffff00',
                        'Low (<30%)': '#00ff00',
                        'No Opening Stock': '#cccccc'
                    }
                )
                fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                fig_pie.update_layout(height=400)
                st.plotly_chart(fig_pie, use_container_width=True)
            
            with col2:
                # Efficiency summary table
                efficiency_summary = efficiency_data.copy()
                efficiency_summary.columns = ['Sales Efficiency', 'Product Count', 'Total Sales']
                efficiency_summary['% of Products'] = (efficiency_summary['Product Count'] / efficiency_summary['Product Count'].sum() * 100).round(1)
                efficiency_summary['% of Sales'] = (efficiency_summary['Total Sales'] / efficiency_summary['Total Sales'].sum() * 100).round(1)
                
                # Format for display
                display_eff = efficiency_summary.copy()
                display_eff['Product Count'] = display_eff['Product Count'].apply(lambda x: f"{x:,}")
                display_eff['Total Sales'] = display_eff['Total Sales'].apply(lambda x: f"{x:,.0f}")
                display_eff['% of Products'] = display_eff['% of Products'].apply(lambda x: f"{x}%")
                display_eff['% of Sales'] = display_eff['% of Sales'].apply(lambda x: f"{x}%")
                
                st.dataframe(
                    display_eff,
                    hide_index=True,
                    use_container_width=True,
                    height=400
                )
            
            st.markdown("---")
            
            # Data validation section
            with st.expander("🔍 Data Validation Details"):
                st.write("**Sample Data (First 10 Rows):**")
                display_cols = ['STYLE_ID', 'YEAR', 'MONTH', 'SALES_QTY', 'OPENING_STOCK', 'SALES_PERCENTAGE', 'SALES_EFFICIENCY']
                if 'Maketplace' in filtered_df.columns:
                    display_cols.append('Maketplace')
                if 'Subcategory' in filtered_df.columns:
                    display_cols.append('Subcategory')
                if 'Brand' in filtered_df.columns:
                    display_cols.append('Brand')
                
                # Display sample data
                sample_df = filtered_df[display_cols].head(10).copy()
                sample_df.columns = ['Style ID', 'Year', 'Month', 'Sales Qty', 'Opening Stock', 'Sales %', 'Efficiency'] + \
                                   (['Marketplace'] if 'Maketplace' in filtered_df.columns else []) + \
                                   (['Subcategory'] if 'Subcategory' in filtered_df.columns else []) + \
                                   (['Brand'] if 'Brand' in filtered_df.columns else [])
                
                # Format sample data
                sample_display = sample_df.copy()
                sample_display['Sales Qty'] = sample_display['Sales Qty'].apply(lambda x: f"{x:,.0f}")
                sample_display['Opening Stock'] = sample_display['Opening Stock'].apply(lambda x: f"{x:,.0f}")
                sample_display['Sales %'] = sample_display['Sales %'].apply(lambda x: f"{x:.1f}%")
                
                st.dataframe(sample_display, use_container_width=True)
                
                st.write("**Data Quality Check:**")
                quality_df = pd.DataFrame({
                    'Metric': [
                        'Total Records',
                        'Unique Products', 
                        'Records with Sales > 0',
                        'Records with Opening Stock > 0',
                        'Average Sales per Record',
                        'Average Opening Stock per Record',
                        'Average Sales Percentage',
                        'Maximum Sales %',
                        'Maximum Sales (Single Record)',
                        'Time Period Covered'
                    ],
                    'Value': [
                        len(filtered_df),
                        filtered_df['STYLE_ID'].nunique(),
                        (filtered_df['SALES_QTY'] > 0).sum(),
                        (filtered_df['OPENING_STOCK'] > 0).sum(),
                        f"{filtered_df['SALES_QTY'].mean():.0f}",
                        f"{filtered_df['OPENING_STOCK'].mean():.0f}",
                        f"{filtered_df[filtered_df['OPENING_STOCK'] > 0]['SALES_PERCENTAGE'].mean():.1f}%",
                        f"{filtered_df['SALES_PERCENTAGE'].max():.1f}%",
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
        - `Opening_stock` - Opening stock quantity
        
        ### **Optional Columns:**
        - `Maketplace` - Sales marketplace
        - `Subcategory` - Product subcategory
        - `Heel_Type 1` - Type of heel
        - `Season` - Product season
        - `Brand` - Product brand
        - `Color` - Product color
        
        ### **Key Features Added:**
        - **Opening Stock Metrics** - Displayed alongside sales
        - **Sales Percentage** - Calculated as (Sales Qty / Opening Stock) * 100
        - **Sales Efficiency Classification** - Categories for sales performance
        - **Dual-axis Charts** - Sales quantity and percentage together
        - **Conditional Formatting** - Color-coded sales percentages
        - **Sorting Options** - Sort products by different metrics
        - **Efficiency Analysis** - Pie chart showing sales efficiency distribution
        
        ### **New Metrics in Tables:**
        1. **Sales Quantity** - Total units sold
        2. **Opening Stock** - Starting inventory
        3. **Sales %** - Percentage of stock sold
        4. **Efficiency** - Performance category
        """)
