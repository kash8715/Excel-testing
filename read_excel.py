import pandas as pd

# Read the Excel file
df = pd.read_excel('DHS1.xlsx')

# Display the data in a more readable format
print("\nContents of the Excel file:")
print("=" * 80)
print(df.to_string())
print("=" * 80)

# Display the shape of the dataset
print(f"\nNumber of rows: {df.shape[0]}")
print(f"Number of columns: {df.shape[1]}")

# Display column names
print("\nColumn names:")
for i, col in enumerate(df.columns):
    print(f"{i+1}. {col}")

# Display basic information about the dataset
print("\nDataset Info:")
print(df.info())

print("\nFirst few rows of the data:")
print(df.head())

print("\nBasic statistics of the data:")
print(df.describe())

print("\nColumns in the dataset:")
print(df.columns.tolist()) 