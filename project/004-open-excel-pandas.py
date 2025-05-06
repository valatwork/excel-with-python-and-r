import pandas as pd

# Load the Excel file
df = pd.read_excel('./data/iris_data.xlsx', sheet_name='setosa')

# Display the first few rows of the DataFrame
print(df.head())