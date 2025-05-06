import openpyxl
import pandas as pd

# Load the Excel file
wb = openpyxl.load_workbook('./data/iris_data.xlsx')

# Select the sheet
sheet = wb['versicolor']

# Extract the values (including headers)
sheet_data_raw = sheet.values

# Separate the headers into a variable
headers = next(sheet_data_raw)[0:]

# Create a DataFrame based on the second and subsequent lines of data with the header as column names
sheet_data = pd.DataFrame(sheet_data_raw, columns=headers)

print(sheet_data.head())