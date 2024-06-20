import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy import stats

# File paths
file_path1 = 'veh0171.xlsx'

# Read the sheet into DataFrame
df4 = pd.read_excel(file_path1, sheet_name='VEH0171b_GenModels')

# Correct column names based on inspection
body_type_col = 'Unnamed: 2'
make_col = 'Unnamed: 3'

# Filter new registrations for cars
car_body_types = ['Cars']
df4_cars = df4[df4[body_type_col].isin(car_body_types)]

# Select columns from index 6 onwards (excluding the first 6 columns)
sales_columns = df4_cars.columns[6:]

# Sum the sales across all years
total_sales = df4_cars[sales_columns].sum(axis=1)
df4_cars['Total Sales'] = total_sales

# Probability Distribution Analysis
mean_sales = total_sales.mean()
median_sales = total_sales.median()
std_dev_sales = total_sales.std()
skewness = stats.skew(total_sales)
kurtosis = stats.kurtosis(total_sales)

print(f"Mean Sales: {mean_sales:.2f}")
print(f"Median Sales: {median_sales:.2f}")
print(f"Standard Deviation: {std_dev_sales:.2f}")
print(f"Skewness: {skewness:.2f}")
print(f"Kurtosis: {kurtosis:.2f}")

# Create histogram
plt.figure(figsize=(10, 6))
plt.hist(total_sales, bins=30, edgecolor='black')
plt.title('Distribution of Total Sales')
plt.xlabel('Total Sales')
plt.ylabel('Frequency')
plt.savefig('total_sales_histogram.png')
plt.close()

# Find outliers using IQR method
Q1 = total_sales.quantile(0.25)
Q3 = total_sales.quantile(0.75)
IQR = Q3 - Q1
lower_bound = Q1 - 1.5 * IQR
upper_bound = Q3 + 1.5 * IQR

outliers = df4_cars[(df4_cars['Total Sales'] < lower_bound) | (df4_cars['Total Sales'] > upper_bound)]

print("\nOutliers:")
print(outliers[[make_col, 'Total Sales']])

# Save the filtered and analyzed data to Excel
output_path = 'cars_veh0171b_GenModels_with_stats.xlsx'
df4_cars.to_excel(output_path, index=False)

print(f"\nAnalyzed data saved to {output_path}")
print("Histogram saved as total_sales_histogram.png")
