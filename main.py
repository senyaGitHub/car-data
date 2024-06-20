import pandas as pd

# File paths
file_path1 = 'veh0171.xlsx'  # Provided file path
# file_path2 = 'veh0181.xlsx'  # Placeholder for the second file

# Read the sheets into DataFrames
df4 = pd.read_excel(file_path1, sheet_name='VEH0171b_GenModels')
# df3 = pd.read_excel(file_path2, sheet_name='VEH0181b_GenModels')

# Display the column names to identify the correct ones
print("Column names:")
print(df4.columns)

# Correct column names based on inspection
body_type_col = 'Unnamed: 2'
make_col = 'Unnamed: 3'

# Filter new registrations for cars
car_body_types = ['Cars']  # List of body types that correspond to cars
df4_cars = df4[df4[body_type_col].isin(car_body_types)]

# Select columns from index 6 onwards (excluding the first 6 columns)
sales_columns = df4_cars.columns[6:]

# Sum the sales across all years
total_sales = df4_cars[sales_columns].sum(axis=1)
df4_cars['Total Sales'] = total_sales
grouped = df4_cars.groupby('Unnamed: 3')


mean_sales = grouped[sales_columns].mean()
median_sales = grouped[sales_columns].median()
range_sales = grouped[sales_columns].max() - grouped[sales_columns].min()
variance_sales = grouped[sales_columns].var()
std_dev_sales = grouped[sales_columns].std()




# Save the filtered and analyzed data to Excel
output_path = 'cars_veh0171b_GenModels_with_stats.xlsx'
df4_cars.to_excel(output_path, index=False)

print("Mean Sales:")
print(mean_sales)
print("\nMedian Sales:")
print(median_sales)
print("\nRange of Sales:")
print(range_sales)
print("\nVariance of Sales:")
print(variance_sales)
print("\nStandard Deviation of Sales:")
print(std_dev_sales)

