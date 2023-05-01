import pandas as pd
import os
import plotly.express as px
import streamlit as st

# Define the path to the data folder
data_folder = 'data'

# Get a list of all the Excel files in the data folder
excel_files = [f for f in os.listdir(data_folder) if f.endswith('.xlsx')]

# Create an empty dataframe to hold the combined data
combined_data = pd.DataFrame()

# Loop through each Excel file and append its data to the combined_data dataframe
for file in excel_files:
    # Load the Excel file into a dataframe
    df = pd.read_excel(os.path.join(data_folder, file))
    
    # Add a column to the dataframe to indicate the source file
    df['Source'] = file
    
    # Append the data to the combined_data dataframe
    combined_data = combined_data.append(df, ignore_index=True)

# Drop duplicate rows
combined_data.drop_duplicates(inplace=True)

# Handle missing values
combined_data.dropna(inplace=True)

# Calculate total revenue, expenses, and profit for each category
category_data = combined_data.groupby('Category').agg({'Revenue': 'sum', 'Expenses': 'sum'})
category_data['Profit'] = category_data['Revenue'] - category_data['Expenses']

# Create a Streamlit dashboard
st.set_page_config(page_title='Total Revenue, Expenses, and Profit by Category', page_icon=':chart_with_upwards_trend:')
st.title('Total Revenue, Expenses, and Profit by Category')
st.subheader('Grouped Data')

# Create an interactive chart using Plotly and display it in the dashboard
fig = px.bar(category_data, x=category_data.index, y=['Revenue', 'Expenses', 'Profit'], barmode='group')
st.plotly_chart(fig)
