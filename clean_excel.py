import pandas as pd
import re

def clean_text(text):
    if pd.isna(text):  # Check if value is NaN
        return text
    # Convert to string if not already
    text = str(text)
    # Remove special characters but keep spaces, numbers and %
    cleaned = re.sub(r'[^a-zA-Z0-9\s%\.]', '', text)
    # Remove extra spaces
    cleaned = ' '.join(cleaned.split())
    return cleaned

# Read the Excel file
print("Reading Excel file...")
df = pd.read_excel('DHS1.xlsx')

# Get and display the original headers
print("\nOriginal Headers:")
headers = df.columns.tolist()
for i, header in enumerate(headers):
    print(f"{i+1}. {header}")

# Clean the headers
cleaned_headers = [clean_text(header) for header in headers]

# Rename columns with cleaned headers
df.columns = cleaned_headers

# Clean the data
print("\nCleaning data...")
# Create a copy of the dataframe with cleaned data
cleaned_df = df.copy()
for column in cleaned_df.columns:
    cleaned_df[column] = cleaned_df[column].apply(clean_text)

# Create a new Excel file with cleaned data
print("\nSaving cleaned data to new Excel file...")
with pd.ExcelWriter('cleaned_dhs_data.xlsx', engine='openpyxl') as writer:
    # Write the data
    cleaned_df.to_excel(writer, sheet_name='Cleaned Data', index=False)
    
    # Auto-adjust column widths
    worksheet = writer.sheets['Cleaned Data']
    for idx, col in enumerate(cleaned_df.columns):
        max_length = max(
            cleaned_df[col].astype(str).apply(len).max(),
            len(str(col))
        )
        # Add a little extra space to the width
        worksheet.column_dimensions[chr(65 + idx)].width = max_length + 4

print("\nCleaned Excel file has been created as 'cleaned_dhs_data.xlsx'")

# Display sample of cleaned data
print("\nSample of cleaned data (first 5 rows):")
pd.set_option('display.max_columns', None)  # Show all columns
pd.set_option('display.width', None)        # Don't wrap to multiple lines
print(cleaned_df.head().to_string()) 