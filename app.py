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
    /* Hide Streamlit warning/duplicate messages */
    .stAlert {
        display: none;
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
        
        # Clean column names - preserve original case but strip spaces
        def clean_columns(df):
            df.columns = df.columns.astype(str).str.strip()
            return df
        
        # Clean column names
        sales_df = clean_columns(sales_df)
        
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
        
        # Add additional columns from sales if they exist - with PROPER TRIMMING
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
                # PROPERLY TRIM TEXT COLUMNS to remove leading/trailing spaces
                if sales_df[found_col].dtype == 'object':  # Text columns
                    sales_clean[standard_name] = sales_df[found_col].astype(str).str.strip()
                else:
                    sales_clean[standard_name] = sales_df[found_col]
        
        # Handle duplicate sales records silently
        duplicate_subset = ['STYLE_ID', 'YEAR', 'MONTH']
        
        # Check if Maketplace column exists and add it to subset
        if 'Maketplace' in sales_clean.columns:
            duplicate_subset.append('Maketplace')
        
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
        
        # Display success message only
        st.success(f"✅ Data loaded successfully! {len(df):,} records processed")
        
        # Display data quality check for text columns
        with st.expander("🔍 Data Quality Check (Text Columns)"):
            text_columns = ['Brand', 'Subcategory', 'Season', 'Color', 'Heel_Type_1', 'Maketplace']
            available_text_cols = [col for col in text_columns if col in df.columns]
            
            quality_info = []
            for col in available_text_cols:
                unique_count = df[col].nunique()
                sample_values = df[col].dropna().unique()[:5].tolist()
                has_spaces = any(str(val).strip() != str(val) for val in df[col].dropna().unique()[:20])
                
                quality_info.append({
                    'Column': col,
                    'Unique Values': unique_count,
                    'Sample Values': ', '.join(str(v) for v in sample_values),
                    'Has Leading/Trailing Spaces': 'Yes' if has_spaces else 'No'
                })
            
            quality_df = pd.DataFrame(quality_info)
            st.dataframe(quality_df, use_container_width=True)
            
            if any(q['Has Leading/Trailing Spaces'] == 'Yes' for q in quality_info):
                st.info("ℹ️ Leading/trailing spaces have been automatically trimmed from all text columns.")
        
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
        
        # Filter data
        filtered_df = df.copy()
        
        if selected_year != 'All':
            filtered_df = filtered_df[filtered_df['YEAR'] == selected_year]
        
        if selected_month != 'All':
            month_num = [k for k, v in month_names_dict.items() if v == selected_month][0]
            filtered_df = filtered_df[filtered_df['MONTH'] == month_num]
        
        # Display filter summary
        st.sidebar.markdown("---")
        st.sidebar.info(f"""
        **Filter Applied:**
        - 📅 Year: {selected_year}
        - 📆 Month: {selected_month}
        - 📊 Records: {len(filtered_df):,}
        """)
        
        if len(filtered_df) == 0:
            st.warning("⚠️ No data available for the selected filters.")
        else:
            # Helper function for grouped analysis with stock metrics
            def analyze_with_stock(df, group_col, group_name):
                if group_col not in df.columns:
                    return pd.DataFrame()
                
                # IMPORTANT: Trim text column values before grouping
                if df[group_col].dtype == 'object':
                    df = df.copy()
                    df[group_col] = df[group_col].astype(str).str.strip()
                
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
                
                # Sort by Sales Qty descending by default
                grouped = grouped.sort_values('SALES_QTY', ascending=False)
                grouped.rename(columns={group_col: group_name}, inplace=True)
                
                return grouped
            
            # Function to create styled dataframe with proper sorting
            def create_sortable_dataframe(data, columns_mapping):
                """
                Create a dataframe with proper formatting that maintains sortability
                
                Parameters:
                data: DataFrame with numeric values
                columns_mapping: dict mapping original column names to display names
                """
                # Create a copy for display
                display_df = data.copy()
                
                # Configure column display formats
                column_config = {}
                
                for col in display_df.columns:
                    if col in columns_mapping:
                        display_name = columns_mapping[col]
                        
                        if 'QTY' in col.upper() or 'STOCK' in col.upper():
                            # Format as numbers with commas
                            column_config[col] = st.column_config.NumberColumn(
                                display_name,
                                help="Click to sort",
                                format="%d"
                            )
                        elif 'PERCENTAGE' in col.upper():
                            # Format as percentages
                            column_config[col] = st.column_config.NumberColumn(
                                display_name,
                                help="Click to sort",
                                format="%.1f%%"
                            )
                        else:
                            # For text columns
                            column_config[col] = st.column_config.TextColumn(
                                display_name,
                                help="Click to sort"
                            )
                
                return display_df, column_config
            
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
                        # Prepare column mapping for display
                        column_mapping = {
                            'Marketplace': 'Marketplace',
                            'SALES_QTY': 'Sales Qty',
                            'OPENING_STOCK': 'Opening Stock',
                            'SALES_PERCENTAGE': 'Sales %'
                        }
                        
                        # Create sortable dataframe
                        display_df, column_config = create_sortable_dataframe(marketplace_data, column_mapping)
                        
                        # Display with column configuration
                        st.dataframe(
                            display_df,
                            column_config=column_config,
                            hide_index=True,
                            use_container_width=True
                        )
                
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
                                    # Prepare column mapping for display
                                    column_mapping = {
                                        display_name: display_name,
                                        'SALES_QTY': 'Sales Qty',
                                        'OPENING_STOCK': 'Opening Stock',
                                        'SALES_PERCENTAGE': 'Sales %'
                                    }
                                    
                                    # Create sortable dataframe
                                    display_df, column_config = create_sortable_dataframe(category_data, column_mapping)
                                    
                                    # Display with column configuration
                                    st.dataframe(
                                        display_df,
                                        column_config=column_config,
                                        hide_index=True,
                                        use_container_width=True,
                                        height=300
                                    )
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
                # Create a copy for display
                trend_table = monthly_data[['Period', 'SALES_QTY', 'OPENING_STOCK', 'SALES_PERCENTAGE']].copy()
                
                # Prepare column mapping for display
                column_mapping = {
                    'Period': 'Period',
                    'SALES_QTY': 'Sales Quantity',
                    'OPENING_STOCK': 'Opening Stock',
                    'SALES_PERCENTAGE': 'Sales %'
                }
                
                # Create sortable dataframe
                display_df, column_config = create_sortable_dataframe(trend_table, column_mapping)
                
                # Display with column configuration
                st.dataframe(
                    display_df,
                    column_config=column_config,
                    hide_index=True,
                    use_container_width=True
                )
            
            st.markdown("---")
            
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
        
        ### **Key Features:**
        - **Opening Stock Metrics** - Displayed alongside sales
        - **Sales Percentage** - Calculated as (Sales Qty / Opening Stock) * 100
        - **Dual-axis Charts** - Sales quantity and percentage together
        - **Interactive Sorting** - Click column headers to sort tables
        
        ### **Metrics in Tables:**
        1. **Sales Quantity** - Total units sold
        2. **Opening Stock** - Starting inventory
        3. **Sales %** - Percentage of stock sold
        """)
