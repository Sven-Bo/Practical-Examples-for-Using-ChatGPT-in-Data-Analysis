import pandas as pd
import xlsxwriter
from pathlib import Path


def load_excel_files(folder):
    """Load all Excel files in a folder into a list of dataframes"""
    excel_files = list(folder.glob('*.xlsx'))
    return [pd.read_excel(file) for file in excel_files]


def combine_dataframes(dataframes):
    """Combine multiple dataframes into a single dataframe"""
    combined_data = pd.concat(dataframes)
    return combined_data.drop_duplicates().dropna()


def calculate_category_data(combined_data):
    """Calculate total revenue, expenses, and profit for each category"""
    category_data = combined_data.groupby('Category').agg({'Revenue': 'sum', 'Expenses': 'sum'})
    category_data['Profit'] = category_data['Revenue'] - category_data['Expenses']
    return category_data


def create_category_report(category_data, report_name):
    """Create an Excel report summarizing category data"""
    workbook = xlsxwriter.Workbook(report_name)
    worksheet = workbook.add_worksheet()

    # Write the category data to the report
    worksheet.write_row(0, 0, ['Category', 'Revenue', 'Expenses', 'Profit'])
    for i, (index, row_data) in enumerate(category_data.iterrows()):
        worksheet.write_row(i+1, 0, [index, row_data['Revenue'], row_data['Expenses'], row_data['Profit']])

    # Create a chart representing the data
    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'name': 'Revenue',
        'categories': ['Sheet1', 1, 0, len(category_data), 0],
        'values': ['Sheet1', 1, 1, len(category_data), 1],
    })
    chart.add_series({
        'name': 'Expenses',
        'categories': ['Sheet1', 1, 0, len(category_data), 0],
        'values': ['Sheet1', 1, 2, len(category_data), 2],
    })
    chart.add_series({
        'name': 'Profit',
        'categories': ['Sheet1', 1, 0, len(category_data), 0],
        'values': ['Sheet1', 1, 3, len(category_data), 3],
    })
    chart.set_title({'name': 'Total Revenue, Expenses, and Profit by Category'})
    chart.set_x_axis({'name': 'Category'})
    chart.set_y_axis({'name': 'Amount'})
    chart.set_legend({'position': 'bottom'})
    worksheet.insert_chart('E1', chart)

    # Close the workbook
    workbook.close()
    print(f'Report saved to {report_name}.')


if __name__ == '__main__':
    data_folder = Path('data')
    excel_dataframes = load_excel_files(data_folder)
    combined_data = combine_dataframes(excel_dataframes)
    category_data = calculate_category_data(combined_data)
    create_category_report(category_data, 'category_report.xlsx')