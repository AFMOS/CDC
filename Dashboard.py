import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import numpy as np

# Define all possible months
all_months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def month_sort_key(month):
    if pd.isna(month):
        return -1
    if isinstance(month, str):
        try:
            return all_months.index(month)
        except ValueError:
            try:
                # Try to convert string to int (e.g., "1" to 1)
                return int(month) - 1
            except ValueError:
                return -1  # Put unknown string months at the beginning
    elif isinstance(month, (int, float)):
        return int(month) - 1  # Assuming 1-based numeric months
    else:
        return -1  # Put unknown types at the beginning

def safe_month_sort(x):
    try:
        return month_sort_key(x)
    except Exception:
        return -1  # Return -1 for any unexpected types

# Set the page configuration to wide mode (only once)
st.set_page_config(layout="wide")

# Main variable selector
main_variable_options = {"Total": "total", "Weight": "weight"}
main_variable = st.sidebar.radio(
    "Select Main Variable",
    options=list(main_variable_options.keys()),
    index=list(main_variable_options.keys()).index("Total")
)
main_variable = main_variable_options[main_variable]

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

# Load the data
@st.cache_data
def load_data(file):
    if file is not None:
        master_data = pd.read_excel(file, sheet_name='Master')
        Target_data = pd.read_excel(file, sheet_name='Target')
    else:
        try:
            master_data = pd.read_excel('Master.xlsx', sheet_name='Master')
            Target_data = pd.read_excel('Master.xlsx', sheet_name='Target')
        except FileNotFoundError:
            st.error("The file 'Master.xlsx' was not found.")
            return None, None
    
    # Ensure column names are correct
    master_data.columns = [
        "document #", "customer_code", "customer_name", "customer_category", "payment_type",
        "salesman", "date", "item_code", "item_description", "item_category", "unit", "qty",
        "price", "total", "weight", "customer_type", "month", "system", "area"
    ]
    
    Target_data.columns = [
        "customer_code", "customer_name", "Target", "2022_sales", "2023_sales", "2024_sales",
        "area", "system", "salesman", "payment_type", "customer_category"
    ]
    
    # Convert 'month' column to string and ensure it's in the correct format
    master_data['month'] = master_data['month'].astype(str)
    master_data['month'] = master_data['month'].apply(lambda x: all_months[int(x)-1] if x.isdigit() and 1 <= int(x) <= 12 else x)
    
    return master_data, Target_data

data, Target_data = load_data(uploaded_file)

if data is None:
    st.stop()

# Check for empty data
if data.empty:
    st.error("The loaded data is empty. Please check your data source.")
    st.stop()

# Get the months present in the data
present_months = sorted(data['month'].dropna().unique(), key=safe_month_sort)

# Define quarter mapping
quarter_map = {
    'Q1': ["Jan", "Feb", "Mar"],
    'Q2': ["Apr", "May", "Jun"],
    'Q3': ["Jul", "Aug", "Sep"],
    'Q4': ["Oct", "Nov", "Dec"]
}

def apply_master_filter(df, search_term):
    if not search_term:
        return df, True

    search_term = search_term.lower()
    mask = pd.Series([False] * len(df))
    
    search_columns = ['customer_code', 'customer_name', 'customer_category', 'salesman', 
                      'item_code', 'item_description', 'item_category', 'month', 'area']
    
    for col in search_columns:
        if col in df.columns:
            mask |= df[col].astype(str).str.lower().str.contains(search_term, regex=True, na=False)
    
    filtered_df = df[mask]
    search_found = len(filtered_df) > 0
    return filtered_df, search_found

def update_filter_options(merged_data):
    options = {}
    for column in ['area', 'month', 'customer_category', 'salesman', 'item_category', 'customer_code', 'customer_name']:
        if column in merged_data.columns:
            unique_values = merged_data[column].dropna().unique()
            # Convert all values to strings before sorting
            sorted_values = sorted(unique_values, key=lambda x: str(x))
            options[column] = ['None'] + sorted_values  # Remove .tolist() here
    options['quarter'] = ['None'] + list(quarter_map.keys())
    return options

def generate_dashboard_title(search_term, selected_filters):
    title_parts = []
    
    if search_term:
        title_parts.append(f'<span style="color: maroon;">"{search_term}"</span>')
    
    for filter_name, filter_value in selected_filters.items():
        if filter_value and filter_value != "None":
            title_parts.append(f'<span style="color: maroon;">{filter_value}</span>')
    
    if title_parts:
        return f"{' - '.join(title_parts)} Breakdown"
    else:
        return "Sales Dashboard"

def update_dashboard(selected_area, selected_month, selected_quarter, 
                     selected_customer_category, selected_salesman, selected_item_category,
                     selected_customer_code, selected_customer_name, main_variable, search_term):
    filtered_data, search_found = apply_master_filter(data, search_term)
    
    if not search_found or filtered_data.empty:
        return None, False

    # Apply filters to Master data
    if selected_area != "None":
        filtered_data = filtered_data[filtered_data['area'] == selected_area]
    if selected_salesman != "None":
        filtered_data = filtered_data[filtered_data['salesman'] == selected_salesman]
    if selected_customer_code != "None":
        filtered_data = filtered_data[filtered_data['customer_code'] == selected_customer_code]
    if selected_customer_name != "None":
        filtered_data = filtered_data[filtered_data['customer_name'] == selected_customer_name]
    if selected_customer_category != "None":
        filtered_data = filtered_data[filtered_data['customer_category'] == selected_customer_category]
    
    if selected_month != "None" and selected_quarter == "None":
        filtered_data = filtered_data[filtered_data['month'] == selected_month]
    elif selected_quarter != "None" and selected_month == "None":
        months = quarter_map[selected_quarter]
        filtered_data = filtered_data[filtered_data['month'].isin(months)]
    elif selected_month != "None" and selected_quarter != "None":
        if selected_month not in quarter_map[selected_quarter]:
            st.warning(f"Selected month {selected_month} is not in the selected quarter {selected_quarter}. Applying only quarter filter.")
            months = quarter_map[selected_quarter]
            filtered_data = filtered_data[filtered_data['month'].isin(months)]
        else:
            filtered_data = filtered_data[filtered_data['month'] == selected_month]
    
    if selected_item_category != "None":
        filtered_data = filtered_data[filtered_data['item_category'] == selected_item_category]

    if filtered_data.empty:
        return None, False

    # Get unique months in the filtered data
    unique_months = sorted(filtered_data['month'].dropna().unique(), key=safe_month_sort)

    # Calculate summary statistics based on Master sheet
    total_sales = filtered_data['total'].sum()
    total_weight = filtered_data['weight'].sum()
    customer_count = filtered_data['customer_code'].nunique()
    
    total_customers = filtered_data['customer_code'].nunique()
    Cash_count = filtered_data[filtered_data['payment_type'] == 'Cash']['customer_code'].nunique()
    Credit_count = filtered_data[filtered_data['payment_type'] == 'Credit']['customer_code'].nunique()
    Cash_percentage = Cash_count / total_customers if total_customers > 0 else 0
    Credit_percentage = Credit_count / total_customers if total_customers > 0 else 0

    # Create charts using Plotly
    # Sales by Area Pie Chart
    sales_by_area = filtered_data.groupby('area')[main_variable].sum().reset_index()
    fig_area = px.pie(sales_by_area, values=main_variable, names='area', title='Sales by Area')

    # Combined Time Graph (Monthly and Quarterly)
    fig_time = go.Figure()

    sales_by_month = filtered_data.groupby('month')[main_variable].sum().reset_index()
    sales_by_month['month'] = pd.Categorical(sales_by_month['month'], categories=unique_months, ordered=True)
    sales_by_month = sales_by_month.sort_values('month', key=lambda x: x.map({m: safe_month_sort(m) for m in unique_months}))
    
    fig_time.add_trace(go.Bar(
        x=sales_by_month['month'],
        y=sales_by_month[main_variable],
        name='Monthly Sales',
        marker_color='rgba(55, 83, 109, 0.7)',
        text=sales_by_month[main_variable].apply(lambda x: f'{x:,.0f}'),
        textposition='inside',
        textfont=dict(color='white'),
    ))

    if len(sales_by_month) > 1:
        sales_by_month['pct_change'] = sales_by_month[main_variable].pct_change() * 100
        for i, row in sales_by_month.iterrows():
            fig_time.add_annotation(
                x=row['month'],
                y=row[main_variable],
                text=f"{row['pct_change']:.1f}%" if not pd.isna(row['pct_change']) else "",
                showarrow=False,
                yshift=20,
                font=dict(size=10, color='rgba(55, 83, 109, 1)')
            )

    # Quarterly sales calculation
    sales_by_quarter = filtered_data.copy()
    sales_by_quarter['quarter'] = sales_by_quarter['month'].apply(lambda x: next((q for q, months in quarter_map.items() if x in months), None))
    sales_by_quarter = sales_by_quarter.groupby('quarter')[main_variable].sum().reset_index()
    quarter_positions = {q: unique_months.index(months[0]) if months[0] in unique_months else 0 for q, months in quarter_map.items()}

    valid_quarters = [q for q in sales_by_quarter['quarter'] if q in quarter_positions]
    
    if valid_quarters:
        try:
            quarterly_x = [unique_months[quarter_positions[q]] for q in valid_quarters]
            quarterly_y = sales_by_quarter[sales_by_quarter['quarter'].isin(valid_quarters)][main_variable].tolist()
            
            fig_time.add_trace(go.Scatter(
                x=quarterly_x,
                y=quarterly_y,
                mode='lines+markers',
                name='Quarterly Sales',
                line=dict(color='rgba(255, 0, 0, 0.8)', width=3),
                marker=dict(size=12, symbol='star', color='rgba(255, 0, 0, 0.8)'),
            ))

            max_y = max(quarterly_y)

            for i, (x, y) in enumerate(zip(quarterly_x, quarterly_y)):
                fig_time.add_annotation(
                    x=x,
                    y=max_y * 1.1,
                    text=f"{y:,.0f}",
                    showarrow=False,
                    font=dict(size=10, color='rgba(255, 0, 0, 1)'),
                    bgcolor='rgba(255, 255, 255, 0.8)',
                    yanchor='bottom'
                )
                
                if i > 0:
                    pct_change = (y - quarterly_y[i-1]) / quarterly_y[i-1] * 100
                    fig_time.add_annotation(
                        x=x,
                        y=max_y * 1.2,
                        text=f"{pct_change:.1f}%",
                        showarrow=False,
                        font=dict(size=10, color='rgba(255, 0, 0, 1)'),
                        bgcolor='rgba(255, 255, 255, 0.8)',
                        yanchor='bottom'
                    )

            # Adjust the y-axis range to accommodate the annotations
            fig_time.update_layout(
                yaxis=dict(range=[0, max_y * 1.3])
            )

        except Exception as e:
            st.error(f"Error in quarterly sales calculation: {str(e)}")
            st.write("Debug info:")
            st.write(f"valid_quarters: {valid_quarters}")
            st.write(f"unique_months: {unique_months}")
            st.write(f"quarter_positions: {quarter_positions}")
            st.write(f"sales_by_quarter: {sales_by_quarter.to_dict()}")
    else:
        st.warning("No valid quarters found in the filtered data.")

    # Sales by Salesman and Area Stacked Bar Chart
    sales_by_salesman_area = filtered_data.groupby(['salesman', 'area'])[main_variable].sum().reset_index()
    salesman_order = sales_by_salesman_area.groupby('salesman')[main_variable].sum().sort_values(ascending=False).index
    fig_salesman = px.bar(sales_by_salesman_area, 
                          x='salesman', 
                          y=main_variable, 
                          color='area',
                          title='Sales by Salesman and Area',
                          labels={main_variable: 'Total Sales', 'salesman': 'Salesman', 'area': 'Area'},
                          category_orders={'salesman': salesman_order},
                          text=main_variable)
    
    fig_salesman.update_traces(texttemplate='%{text:.0s}', textposition='inside')
    fig_salesman.update_layout(
        xaxis_title='',
        yaxis_title='',
        xaxis=dict(tickangle=-45),
        legend_title='Area',
        barmode='stack'
    )

    # Generate heatmaps
    def create_heatmap(pivot_table, title):
        if not pivot_table.empty:
            # Sort columns based on month_sort_key
            pivot_table = pivot_table.reindex(columns=sorted(pivot_table.columns, key=safe_month_sort))
            
            rounded_values = pivot_table.round(0).fillna(0)
            text_values = rounded_values.astype(str)
            text_values = text_values.replace('0.0', '', regex=False)
            
            fig = go.Figure(data=go.Heatmap(
                z=pivot_table.values,
                x=pivot_table.columns,
                y=pivot_table.index,
                colorscale='Brwnyl',
                hoverongaps=False,
                text=text_values,
                texttemplate="%{text}",
                showscale=False
            ))
            
            fig.update_layout(
                title=title,
                xaxis_title='',
                yaxis_title='',
                height=max(25 * len(pivot_table.index), 400),  # Ensure minimum height
                xaxis=dict(side='top')
            )
        else:
            fig = go.Figure()
            fig.update_layout(title=f'No data available for {title}')
        return fig

    # Item Category Heatmap
    pivot_table_item = filtered_data.pivot_table(index='item_category', columns='month', values=main_variable, aggfunc='sum', fill_value=0)
    pivot_table_item = pivot_table_item.reindex(columns=unique_months)
    pivot_table_item['total'] = pivot_table_item.sum(axis=1)
    pivot_table_item = pivot_table_item.sort_values('total', ascending=True)
    pivot_table_item = pivot_table_item.drop('total', axis=1)
    fig_heatmap_item = create_heatmap(pivot_table_item, 'Item Category Heatmap')

    # Item Description Heatmap
    pivot_table_item_description = filtered_data.pivot_table(index='item_description', columns='month', values=main_variable, aggfunc='sum', fill_value=0)
    pivot_table_item_description = pivot_table_item_description.reindex(columns=unique_months)
    pivot_table_item_description['total'] = pivot_table_item_description.sum(axis=1)
    pivot_table_item_description = pivot_table_item_description.sort_values('total', ascending=True)
    pivot_table_item_description = pivot_table_item_description.drop('total', axis=1)
    fig_heatmap_item_description = create_heatmap(pivot_table_item_description, 'Item Description Heatmap')

    # Customer Heatmap
    pivot_table_customer = filtered_data.pivot_table(index='customer_name', columns='month', values=main_variable, aggfunc='sum', fill_value=0)
    pivot_table_customer = pivot_table_customer.reindex(columns=unique_months)
    pivot_table_customer['total'] = pivot_table_customer.sum(axis=1)
    pivot_table_customer = pivot_table_customer.sort_values('total', ascending=True)
    pivot_table_customer = pivot_table_customer.drop('total', axis=1)
    fig_heatmap_customer = create_heatmap(pivot_table_customer, 'Customer Heatmap')

    return (total_sales, total_weight, customer_count, Cash_count, Credit_count,
            Cash_percentage, Credit_percentage, fig_area, fig_time, fig_salesman, 
            fig_heatmap_item, fig_heatmap_item_description, fig_heatmap_customer), True

# Streamlit app
st.sidebar.header('Filters')

# Generate autocomplete suggestions
def get_autocomplete_suggestions(data):
    suggestions = []
    for col in ['customer_code', 'customer_name', 'customer_category', 'salesman', 
                'item_code', 'item_description', 'item_category', 'month', 'area']:
        suggestions.extend(data[col].astype(str).unique().tolist())
    return list(set(suggestions))

autocomplete_suggestions = get_autocomplete_suggestions(data)

# Add master search filter with autocomplete
search_term = st.sidebar.text_input(
    "Master Search Filter",
    key="master_search",
    help="Start typing to search..."
)

# Display autocomplete suggestions
st.sidebar.markdown(
    f"""
    <datalist id="suggestions">
        {"".join(f"<option value='{item}'>" for item in autocomplete_suggestions)}
    </datalist>
    <script>
        var input = document.querySelector('input[aria-label="Master Search Filter"]');
        input.setAttribute('list', 'suggestions');
    </script>
    """,
    unsafe_allow_html=True
)

# Initial filter options
filter_options = update_filter_options(data)

# Apply filters and update options
filtered_data, _ = apply_master_filter(data, search_term)
filter_options = update_filter_options(filtered_data)

selected_area = st.sidebar.selectbox('Select Area', options=filter_options.get('area', ['None']))
selected_month = st.sidebar.selectbox('Select Month', options=filter_options.get('month', ['None']))
selected_quarter = st.sidebar.selectbox('Select Quarter', options=filter_options['quarter'])
selected_customer_category = st.sidebar.selectbox('Select Customer Category', options=filter_options.get('customer_category', ['None']))
selected_salesman = st.sidebar.selectbox('Select Salesman', options=filter_options.get('salesman', ['None']))
selected_item_category = st.sidebar.selectbox('Select Item Category', options=filter_options.get('item_category', ['None']))
selected_customer_code = st.sidebar.selectbox('Select Customer Code', options=filter_options.get('customer_code', ['None']))
selected_customer_name = st.sidebar.selectbox('Select Customer Name', options=filter_options.get('customer_name', ['None']))

# Generate dynamic dashboard title
selected_filters = {
    'Area': selected_area,
    'Month': selected_month,
    'Quarter': selected_quarter,
    'Customer Category': selected_customer_category,
    'Salesman': selected_salesman,
    'Item Category': selected_item_category,
    'Customer Code': selected_customer_code,
    'Customer Name': selected_customer_name
}
dashboard_title = generate_dashboard_title(search_term, selected_filters)

# Display the dynamic dashboard title
st.markdown(f"""<h1 style="text-align: center;">📊 {dashboard_title}</h1>""", unsafe_allow_html=True)

# In the main Streamlit app section, add a debug print for the filtered data:
try:
    dashboard_data, search_found = update_dashboard(
        selected_area, selected_month, selected_quarter, 
        selected_customer_category, selected_salesman, selected_item_category,
        selected_customer_code, selected_customer_name,
        main_variable, search_term
    )
    
    if not search_found:
        st.warning("No data found matching the search criteria. Please adjust your filters or search term.")
    elif dashboard_data is None:
        st.warning("No data available for the selected filters. Please adjust your selections.")
    else:
        # Display dashboard components
        (total_sales, total_weight, customer_count, Cash_count, Credit_count,
         Cash_percentage, Credit_percentage, fig_area, fig_time, fig_salesman, 
         fig_heatmap_item, fig_heatmap_item_description, fig_heatmap_customer) = dashboard_data

        # Layout with columns for summary statistics
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.markdown(f"""
            <div style="background-color: #f0f2f6; padding: 10px; border-radius: 5px; font-size: 12px; text-align: center; height: 100%;">
                <h4 style="margin: 0; font-size: 14px;">Total</h4>
                <h2 style="margin: 0; font-size: 16px;">
                SAR {total_sales:,.0f}
                </h2>
            </div>
            """, unsafe_allow_html=True)

        with col2:
            st.markdown(f"""
            <div style="background-color: #f0f2f6; padding: 10px; border-radius: 5px; font-size: 12px; text-align: center; height: 100%;">
                <h4 style="margin: 0; font-size: 14px;">Weight</h4>
                <h2 style="margin: 0; font-size: 16px;">
                {total_weight:,.0f} Kg
                </h2>
            </div>
            """, unsafe_allow_html=True)

        with col3:
            st.markdown(f"""
            <div style="background-color: #f0f2f6; padding: 10px; border-radius: 5px; font-size: 12px; text-align: center; height: 100%;">
                <h4 style="margin: 0; font-size: 14px;">Customer Count</h4>
                <h2 style="margin: 0; font-size: 16px;">{customer_count:,}</h2>
            </div>
            """, unsafe_allow_html=True)

        with col4:
            st.markdown(f"""
            <div style="background-color: #f0f2f6; padding: 10px; border-radius: 5px; font-size: 12px; text-align: center; height: 100%;">
                <h4 style="margin: 0; font-size: 14px;">Payment Type</h4>
                <h2 style="margin: 0; font-size: 16px;">Cash: {Cash_count:,} ({Cash_percentage:.1%})</h2>
                <h2 style="margin: 0; font-size: 16px;">Credit: {Credit_count:,} ({Credit_percentage:.1%})</h2>
            </div>
            """, unsafe_allow_html=True)

        # Create tabs for charts
        tab1, tab2, tab3 = st.tabs(["Sales Overview", "Time Graphs", "Heatmaps"])

        with tab1:
            col1, col2 = st.columns(2)
            with col1:
                st.plotly_chart(fig_area, use_container_width=True)
            with col2:
                st.plotly_chart(fig_salesman, use_container_width=True)

        with tab2:
            st.plotly_chart(fig_time, use_container_width=True)

        with tab3:
            heatmap_option = st.selectbox(
                "Select Heatmap",
                ["Item Category", "Item Description", "Customer"]
            )
            
            if heatmap_option == "Item Category":
                st.plotly_chart(fig_heatmap_item, use_container_width=True)
            elif heatmap_option == "Item Description":
                st.plotly_chart(fig_heatmap_item_description, use_container_width=True)
            else:  # Customer
                st.plotly_chart(fig_heatmap_customer, use_container_width=True)

except Exception as e:
    st.error(f"An error occurred while updating the dashboard: {str(e)}")
    st.write("Error details:", e)
    st.write("Data types in 'month' column:", data['month'].dtype)
    st.write("Unique values in 'month' column:", data['month'].unique())