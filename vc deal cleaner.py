import pandas as pd

# Load the data from CSV
df = pd.read_csv('1997-2023 United States Pre-Seed, Seed, and Series A Deals (Crunchbase) - 1997-2023 Pre-Seed, Seed, and Series A Deals.csv')

# Convert the 'Announced Date' column to datetime and set it as the index
df['Announced Date'] = pd.to_datetime(df['Announced Date'])
df.set_index('Announced Date', inplace=True)

# Create a list of funding types
funding_types = ['Pre-Seed', 'Seed', 'Series A']

# Initialize an empty dictionary to store dataframes
dfs = {}

# Loop through the funding types
for funding in funding_types:
    # Filter the dataframe for the current funding type
    df_funding = df[df['Funding Type'] == funding]
    
    # Calculate the quarterly averages and round to nearest whole number
    df_funding = df_funding.resample('Q')['Money Raised Currency (in USD)'].mean().round()
    
    # Filter for Q1 2007 onwards
    df_funding = df_funding[df_funding.index >= '2007-01-01']
    
    # Convert index to "Q1 2007" format
    df_funding.index = df_funding.index.to_period('Q').astype(str)
    
    # Store the dataframe in the dictionary
    dfs[funding] = df_funding

# Now dfs is a dictionary where the keys are the funding types and the values are
# dataframes with the quarterly averages for each funding type

# Save each dataframe to an Excel file
for funding, df_funding in dfs.items():
    df_funding.to_excel(f'{funding}_quarterly_average.xlsx')
    
# Now let's create the 'Early Stage' dataframe, which is the average of all funding types
df_early_stage = pd.concat(dfs.values(), axis=1).mean(axis=1).round()
df_early_stage = df_early_stage[df_early_stage.index >= '2007Q1']
df_early_stage.to_excel('Early_Stage_quarterly_average.xlsx')
