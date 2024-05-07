import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import seaborn as sns
import plotly.express as px
import numpy as np

#Dashboard Title
st.title("Minger Sales Analysis")

df = pd.read_excel('cleaned_data.xlsx', engine='openpyxl')

#Logo
col1, col2, col3= st.columns(3)
with col1:
    st.image("Minger Logo.png", width=450)

#Visualization 01 & 02- Sum of sales and profit
total_sales = df['Sales'].sum()
total_profit = df['Profit'].sum()

col1, col2 = st.columns(2)

with col1:
    st.subheader('Total Sales')
    st.metric(label="Sales", value=f"${total_sales:,.0f}")

with col2:
    st.subheader('Total Profit')
    st.metric(label="Profit", value=f"${total_profit:,.0f}")


#Visualization 03 & 04 - Top products with sales and profitability
Top_products = df.groupby('Product Name')['Sales'].sum().nlargest(10).round()
Profitable_products = df.groupby("Product Name")["Profit"].sum().nlargest(10).round()

st.subheader('Top 10 Best-Selling Products')
st.bar_chart(Top_products.sort_values() , color= "#00008B")
st.subheader('Top 10 Profitable Products')
st.bar_chart(Profitable_products.sort_values(), color= "#00008B")


#Visualization 05 & 06 - Pie chart showing Sales by sub category and customer segment

df['Order Date'] = pd.to_datetime(df['Order Date']).dt.date
df['Year'] = pd.to_datetime(df['Order Date']).dt.year

# Set up filters and selections
years = list(df['Year'].unique())
years.sort()
year_options = ["All Years"] + years
selected_year = st.selectbox('Select Year', year_options, key='year_select')

# Filter data based on selection
if selected_year == "All Years":
    filtered_data = df
else:
    filtered_data = df[df['Year'] == selected_year]

# Set up columns for layout
col1, col2 = st.columns(2)

with col1:
    st.subheader(f'Sales by Customer Segment in {selected_year}')
    segment_sales = filtered_data.groupby('Segment')['Sales'].sum().round()
    fig1, ax1 = plt.subplots()
    colors = ['#AEC6CF', '#77DD77', '#FDFD96'] 
    ax1.pie(segment_sales, labels=segment_sales.index, autopct=lambda p: f'{p * sum(segment_sales) / 100:,.0f}', startangle=90, colors=colors)
    ax1.axis('equal') 
    st.pyplot(plt)

with col2:
    st.subheader(f'Sales by Sub-Category in {selected_year}')
    sales_by_subcategory = filtered_data.groupby('Sub-Category')['Sales'].sum().round()
    plt.figure(figsize=(12, 10))
    colors = ['#B39EB5', '#FFB347', '#FF6961', '#AEC6CF']
    bars = plt.bar(sales_by_subcategory.index, sales_by_subcategory.values, color=colors[:len(sales_by_subcategory)])
    plt.xlabel('Sub-Category', fontsize=20)
    plt.ylabel('Total Sales', fontsize=20)
    plt.xticks(rotation=45, fontsize=18)
    plt.yticks(fontsize=18)
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval, f'{yval:,.0f}', ha='center', va='bottom', fontsize=16)  # Display numeric values
    st.pyplot(plt)


#Visualization 07 & 08 - Choropleth of sales and profit by country

country_data = df.groupby('Country').agg({'Sales': 'sum', 'Profit': 'sum'}).reset_index()
country_data['Sales'] = country_data['Sales'].round() 
country_data['Profit'] = country_data['Profit'].round() 

#Choropleth map for Sales
fig_sales = px.choropleth(
    country_data,
    locations="Country",
    locationmode='country names', 
    color="Sales",
    hover_name="Country",
    hover_data={"Sales": True, "Profit": True},
    color_continuous_scale=px.colors.sequential.Plasma
)

#Choropleth map for Profit
fig_profit = px.choropleth(
    country_data,
    locations="Country",
    locationmode='country names',
    color="Profit",
    hover_name="Country",
    hover_data={"Sales": True, "Profit": True},
    color_continuous_scale=px.colors.sequential.Viridis
)

# Create tabs for Sales and Profit maps
tab1, tab2 = st.tabs(["Sales by Country", "Profit by Country"])

with tab1:
    st.subheader("Sales by Country")
    st.plotly_chart(fig_sales)

with tab2:
    st.subheader("Profit by Country")
    st.plotly_chart(fig_profit)


#Visualization 09 - Monthly Sales Trends
st.subheader("Monthly Sales Trends")

df['Order Date'] = pd.to_datetime(df['Order Date'])
df['Order Date'] = df['Order Date'].dt.date 

# Slider setup
min_date = df['Order Date'].min()
max_date = df['Order Date'].max()

st.subheader("Select Date Range for Sales Trends")
start_date, end_date = st.slider(
    "Select a date range",
    value=(min_date, max_date),
    format="YYYY-MM-DD",
    min_value=min_date,
    max_value=max_date
)

# Filtering data based on the selection
MonthlyTrend_date = df[(df['Order Date'] >= start_date) & (df['Order Date'] <= end_date)].copy()

# Aggregating the filtered data by month
MonthlyTrend_date['YearMonth'] = MonthlyTrend_date['Order Date'].apply(lambda x: x.strftime('%Y-%m'))
monthly_sales = MonthlyTrend_date.groupby('YearMonth')['Sales'].sum()

st.line_chart(monthly_sales)


#Visualization 10 - Heatmap of Sales by Region and Category
st.subheader("Sales by Region and Category")

#Pivot table
pivot_table = df.pivot_table(values='Sales', index='Region', columns='Category', aggfunc='sum', fill_value=0)

#Heatmap
plt.figure(figsize=(10, 8))
sns.heatmap(pivot_table, annot=True, fmt=".0f", cmap='viridis') 
plt.ylabel('Region')
plt.xlabel('Category')

st.pyplot(plt)


#Visualization 11 - Profit by region
st.subheader("Profit by Region")

region_profit = df.groupby('Region')['Profit'].sum().sort_values(ascending=False)

#Column chart
plt.figure(figsize=(10, 6))
colors = ['#B39EB5', '#FFB347', '#FF6961', '#AEC6CF']  
plt.bar(region_profit.index, region_profit.values, color=colors)
plt.xlabel('Region')
plt.ylabel('Total Profit')
plt.xticks(rotation=45) 

st.pyplot(plt)


#Visualization 12 -Orders by ship mode
st.subheader("Number of Orders by Ship Mode")
orders_by_ship_mode = df['Ship Mode'].value_counts()

#Total sales by Ship Mode
sales_by_ship_mode = df.groupby('Ship Mode')['Sales'].sum()

# Bar chart
plt.figure(figsize=(10, 5))
colors = ['#008000', '#FFC0CB', '#FF0000', '#A52A2A']  
plt.bar(orders_by_ship_mode.index, orders_by_ship_mode.values, color=colors)
plt.xlabel('Ship Mode')
plt.ylabel('Number of Orders')
plt.xticks(rotation=45) 

st.pyplot(plt) 

#Visualization 13 - Ship Mode Performance: Sales vs. Profit with Quantity as Size

st.subheader("Ship Mode Performance: Sales vs. Profit with Quantity as Size")
ship_mode_stats = df.groupby('Ship Mode').agg({
    'Sales': 'sum',
    'Profit': 'sum',
    'Quantity': 'sum'
}).reset_index()

#Bubble chart
fig = px.scatter(
    ship_mode_stats,
    x='Sales', 
    y='Profit',
    size='Quantity',  
    color='Profit', 
    hover_name='Ship Mode',
    size_max=60,
    labels={"Sales": "Total Sales", "Profit": "Total Profit", "Quantity": "Total Quantity"},
    text='Ship Mode'
)

fig.update_traces(textposition='top center')

st.plotly_chart(fig)


#Visualization 14 - Sales vs Discount
st.subheader("Sales vs Discount according to category")

#Slicer setup
product_categories = df['Category'].unique()
selected_category = st.selectbox('Select Product Category', product_categories)

# Filtering data based on selected category
filtered_data = df[df['Category'] == selected_category]

#Scatter Plot
fig = px.scatter(filtered_data, x='Discount', y='Sales',
                 title=f'Sales vs. Discount for {selected_category}',
                 labels={'Discount': 'Discount (%)', 'Sales': 'Sales ($)'},
                 hover_data=['Product Name'],
                 color='Discount', 
                 template='simple_white') 

st.plotly_chart(fig)
