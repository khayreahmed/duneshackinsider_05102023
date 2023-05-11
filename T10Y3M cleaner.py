import pandas as pd

# Load the data from CSV
df = pd.read_csv('T10Y3M.csv')

# Convert the 'DATE' column to datetime and set it as the index
df['DATE'] = pd.to_datetime(df['DATE'])
df.set_index('DATE', inplace=True)

# Convert the 'T10Y3M' column to numeric, coercing invalid values to NaN
df['T10Y3M'] = pd.to_numeric(df['T10Y3M'], errors='coerce')

# Calculate the quarterly averages and round to two decimal places
df_quarterly = df.resample('Q')['T10Y3M'].mean().round(2)

# Filter for Q1 2007 onwards
df_quarterly = df_quarterly[df_quarterly.index >= '2007-01-01']

# Convert index to "Q1 2007" format
df_quarterly.index = df_quarterly.index.to_period('Q').astype(str)

# Save the dataframe to an Excel file
df_quarterly.to_excel('Yield_Curve_quarterly_average.xlsx')
