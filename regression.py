import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn import linear_model
from mpl_toolkits.mplot3d import Axes3D
import os

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

######################## Prepare model data point for visualization ###############################

x = X[:, 0]
y = X[:, 1]
z = Y

x_pred = np.linspace(x.min(), x.max(), 30)   # range of yield curve values
y_pred = np.linspace(y.min(), y.max(), 30)   # range of unemployment rate values
xx_pred, yy_pred = np.meshgrid(x_pred, y_pred)
model_viz = np.array([xx_pred.flatten(), yy_pred.flatten()]).T

################################################ Train #############################################

ols = linear_model.LinearRegression()
model = ols.fit(X, Y)
predicted = model.predict(model_viz)

############################################## Evaluate ############################################

r2 = model.score(X, Y)

############################################## Plot ################################################

plt.style.use('default')

fig = plt.figure(figsize=(8, 6))

ax = fig.add_subplot(111, projection='3d')

ax.plot(x, y, z, color='k', zorder=15, linestyle='none', marker='o', alpha=0.5)
ax.scatter(xx_pred.flatten(), yy_pred.flatten(), predicted, facecolor=(0,0,0,0), s=20, edgecolor='#70b3f0')
ax.set_xlabel('Yield Curve', fontsize=12)
ax.set_ylabel('Unemployment Rate', fontsize=12)
ax.set_zlabel('Deal Size (USD)', fontsize=12)
ax.locator_params(nbins=4, axis='x')
ax.locator_params(nbins=5, axis='x')

# Create images directory if it does not exist
if not os.path.exists('images'):
    os.makedirs('images')

for ii in np.arange(0, 360, 1):
    ax.view_init(elev=32, azim=ii)
    fig.suptitle('$R^2 = %.2f$' % r2, fontsize=20)
    fig.tight_layout()
    plt.savefig('images/gif_image%d.png' % ii)

