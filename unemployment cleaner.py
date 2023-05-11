import pandas as pd

# Load the data from CSV, skipping the first 11 rows
df = pd.read_excel('Unemployment (1997 - 2023).xlsx', skiprows=range(11))

# Melt the dataframe to long format and rename the columns
df_melted = df.melt(id_vars='Year', var_name='Month', value_name='Unemployment Rate')

# Create a datetime index
df_melted['Date'] = pd.to_datetime(df_melted['Year'].astype(str) + '-' + df_melted['Month'])
df_melted.set_index('Date', inplace=True)

# Drop the Year and Month columns
df_melted.drop(['Year', 'Month'], axis=1, inplace=True)

# Calculate the quarterly averages and round to two decimal places
df_quarterly = df_melted.resample('Q')['Unemployment Rate'].mean().round(2)

# Filter for Q1 2007 onwards
df_quarterly = df_quarterly.loc['2007-01-01':'2023-06-30']

# Convert index to "Q1 2007" format
df_quarterly.index = df_quarterly.index.to_period('Q').astype(str)

df_quarterly.at['2023Q2'] = 3.4

# Save the dataframe to an Excel file
df_quarterly.to_excel('Unemployment_Rate_quarterly_average.xlsx')
