import pandas as pd
import numpy as np
from sklearn import linear_model

######################################## Data preparation #########################################

# Load data from the Excel files
df_deal_size = pd.read_excel('Early_Stage_quarterly_average.xlsx')
df_yield_curve = pd.read_excel('Yield_Curve_quarterly_average.xlsx')
df_unemployment = pd.read_excel('Unemployment_Rate_quarterly_average.xlsx')

# Rename columns for merging
df_deal_size.rename(columns={'Announced Date': 'Date'}, inplace=True)
df_yield_curve.rename(columns={'DATE': 'Date', 'T10Y3M': 'Yield Curve'}, inplace=True)
df_unemployment.rename(columns={'Date': 'Date', 'Unemployment Rate': 'Unemployment Rate'}, inplace=True)

# Merge the dataframes on the 'Date' column
df = pd.merge(df_deal_size, df_yield_curve, on='Date')
df = pd.merge(df, df_unemployment, on='Date')

# Convert 'Money Raised Currency (in USD)', 'Yield Curve', 'Unemployment Rate' to numeric
df[['Money Raised Currency (in USD)', 'Yield Curve', 'Unemployment Rate']] = df[['Money Raised Currency (in USD)', 'Yield Curve', 'Unemployment Rate']].apply(pd.to_numeric)

X = df[['Yield Curve', 'Unemployment Rate']].values.reshape(-1,2)
Y = df['Money Raised Currency (in USD)']

################################################ Train #############################################

ols = linear_model.LinearRegression()
model = ols.fit(X, Y)

############################################## Evaluate ############################################

r2 = model.score(X, Y)

# Print regression results
print('Intercept: \n', model.intercept_)
print('Coefficients: \n', model.coef_)
print('R-squared: \n', r2)
